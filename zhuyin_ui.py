# zhuyin_ui.py
# 提供：get_zhuyin(name), load_zhuyin_config(), save_zhuyin_config(), ZhuyinSettingsDialog
# 依賴：pypinyin（若可用則顯示注音），並使用 additions.py 中的 db_config_get_raw/db_config_set_raw 儲存設定

import json
import tkinter as tk
from tkinter import messagebox
try:
    from tkinter import ttk
except Exception:
    ttk = None

# 嘗試使用 pypinyin 產生注音（Bopomofo）
try:
    from pypinyin import lazy_pinyin, Style
    _PYPINYIN_AVAILABLE = True
except Exception:
    _PYPINYIN_AVAILABLE = False

# DB helpers（依賴 additions.py 提供）
try:
    from additions import db_config_get_raw, db_config_set_raw
except Exception:
    # fallback: provide no-op DB functions to avoid crash if additions not loaded.
    def db_config_get_raw(key): return None
    def db_config_set_raw(key, value): pass

ZHUYIN_CONFIG_KEY = "zhuyin_config"
DEFAULT_ZHUYIN_CONFIG = {
    "enabled": False,     # 是否顯示注音（預設關閉）
    "sample_fontsize": 12
}

def load_zhuyin_config():
    raw = db_config_get_raw(ZHUYIN_CONFIG_KEY)
    if not raw:
        return DEFAULT_ZHUYIN_CONFIG.copy()
    try:
        cfg = json.loads(raw)
        out = DEFAULT_ZHUYIN_CONFIG.copy()
        out.update(cfg)
        out["enabled"] = bool(out.get("enabled", DEFAULT_ZHUYIN_CONFIG["enabled"]))
        out["sample_fontsize"] = int(out.get("sample_fontsize", DEFAULT_ZHUYIN_CONFIG["sample_fontsize"]))
        return out
    except Exception:
        return DEFAULT_ZHUYIN_CONFIG.copy()

def save_zhuyin_config(cfg):
    try:
        db_config_set_raw(ZHUYIN_CONFIG_KEY, json.dumps(cfg, ensure_ascii=False))
    except Exception:
        pass

def get_zhuyin(name: str) -> str:
    """
    回傳注音標示 (Bopomofo) 並以空格分隔每個字的注音部份。
    若 pypinyin 不可用則回傳空字串。
    範例: "ㄩㄢˋ ㄔㄨㄣˊ"
    """
    if not name:
        return ""
    if not _PYPINYIN_AVAILABLE:
        return ""
    try:
        parts = lazy_pinyin(name, style=Style.BOPOMOFO, errors='default')
        # join with space for readability
        return " ".join(parts)
    except Exception:
        return ""

class ZhuyinSettingsDialog(tk.Toplevel):
    """
    簡單的注音設定視窗（是否顯示注音 + 範例字型大小）
    on_save(cfg) optional callback
    """
    def __init__(self, master, on_save=None):
        super().__init__(master)
        self.transient(master)
        self.grab_set()
        self.title("注音設定")
        self.geometry("360x160")
        self.on_save = on_save

        self.cfg = load_zhuyin_config()

        frm = tk.Frame(self)
        frm.pack(fill="both", expand=True, padx=12, pady=10)

        self.enabled_var = tk.BooleanVar(self, value=self.cfg.get("enabled", False))
        tk.Checkbutton(frm, text="顯示注音 (注音符號 ㄅㄆㄇ)", variable=self.enabled_var).pack(anchor="w", pady=(0,8))

        tk.Label(frm, text="範例字型大小 (注音顯示):").pack(anchor="w")
        self.fontsize_var = tk.IntVar(self, value=self.cfg.get("sample_fontsize", 12))
        fs_frame = tk.Frame(frm)
        fs_frame.pack(anchor="w", pady=(2,8))
        tk.Entry(fs_frame, textvariable=self.fontsize_var, width=6).pack(side="left")
        tk.Label(fs_frame, text="px").pack(side="left", padx=(6,0))

        # sample display
        sample_frame = tk.Frame(frm)
        sample_frame.pack(fill="x", pady=(4,6))
        tk.Label(sample_frame, text="範例：").pack(side="left")
        self.sample_label = tk.Label(sample_frame, text=self._sample_text(), fg="gray")
        self.sample_label.pack(side="left", padx=(6,0))

        btnf = tk.Frame(self)
        btnf.pack(fill="x", pady=(6,6))
        tk.Button(btnf, text="保存", bg="#4CAF50", fg="white", command=self._on_save).pack(side="left", padx=6)
        tk.Button(btnf, text="取消", command=self.destroy).pack(side="left")

        # update sample when fontsize changes
        self.fontsize_var.trace_add("write", lambda *a: self._update_sample())

    def _sample_text(self):
        # sample two-chinese chars and zhuyin if available
        sample_name = "媛純"  # example; not critical
        z = get_zhuyin(sample_name) if _PYPINYIN_AVAILABLE else "(pypinyin not installed)"
        return z

    def _update_sample(self):
        try:
            size = int(self.fontsize_var.get())
        except Exception:
            size = 12
        self.sample_label.config(text=self._sample_text(), font=("Microsoft JhengHei", size))

    def _on_save(self):
        try:
            cfg = {
                "enabled": bool(self.enabled_var.get()),
                "sample_fontsize": int(self.fontsize_var.get())
            }
        except Exception as e:
            messagebox.showwarning("輸入錯誤", f"請檢查輸入值格式: {e}")
            return
        save_zhuyin_config(cfg)
        if self.on_save:
            try:
                self.on_save(cfg)
            except Exception:
                pass
        messagebox.showinfo("保存成功", "注音設定已儲存。")
        self.destroy()