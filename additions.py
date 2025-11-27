# additions.py
# 提供三項功能：
# A) TTS 設定 UI（節流 throttle / 開關 / 語速 / 音量）並存 DB (config.key='tts_config')
# B) 字屬性 GUI 編輯器（CharAttributesEditor）支援匯入/匯出 JSON 並修改 CHAR_ATTRS
# C) 快捷鍵註冊函式（空白抽取、t 發音、u 撤銷）
#
# 使用方式（在 姓名產生器.py）：
#   from additions import TTSSettingsDialog, CharAttributesEditor, register_shortcuts, load_tts_config, save_tts_config
#   然後在 _setup_ui 之後呼叫 register_shortcuts(self.master, app_instance)
#   並在適當處放置按鈕或選單來打開 TTSSettingsDialog(self.master) 與 CharAttributesEditor(self.master)
#
# 這個模組會嘗試找到 name_generator_data 資料夾並使用該目錄下的 sqlite DB 與 char_attributes.json。
# 如果你把資料夾改名字，請在 main 程式中傳入正確路徑（簡單改動即可）。

import json
import os
import sqlite3
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog, scrolledtext
from datetime import datetime

# 嘗試使用與主程式相同的資料夾名稱（與你的主程式保持一致）
DATA_DIR = "name_generator_data"
DB_FILE = os.path.join(DATA_DIR, "name_generator.sqlite3")
CHAR_ATTR_FILE = os.path.join(DATA_DIR, "char_attributes.json")

# ---------- DB config helpers (獨立於主檔，可用 sqlite 直接讀寫 config 表) ----------
def _db_connect():
    if not os.path.exists(DB_FILE):
        # caller should ensure DB 已建立 (主程式會呼叫 init_db)
        # 仍然嘗試建立一個檔案連線
        conn = sqlite3.connect(DB_FILE, timeout=10, isolation_level=None)
        try:
            conn.execute("PRAGMA journal_mode=WAL;")
        except Exception:
            pass
        return conn
    conn = sqlite3.connect(DB_FILE, timeout=10, isolation_level=None)
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
    except Exception:
        pass
    return conn

def db_config_get_raw(key):
    try:
        with _db_connect() as conn:
            cur = conn.cursor()
            cur.execute("SELECT value FROM config WHERE key = ?;", (key,))
            row = cur.fetchone()
            return row[0] if row else None
    except Exception:
        return None

def db_config_set_raw(key, value):
    try:
        with _db_connect() as conn:
            cur = conn.cursor()
            cur.execute("INSERT INTO config(key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value;", (key, value))
    except Exception:
        pass

# ---------- TTS config load/save ----------
DEFAULT_TTS_CONFIG = {
    "enabled": True,        # 是否允許 TTS
    "interrupt": True,      # 抽取時是否中斷當前播放並播放最新
    "throttle_ms": 300,     # cooldown(ms)；在此時間內若再次抽取，根據 mode 行為決定
    "throttle_mode": "interrupt",  # "interrupt" / "skip" / "debounce"
    "rate": 160,
    "volume": 1.0
}

def load_tts_config():
    raw = db_config_get_raw("tts_config")
    if not raw:
        return DEFAULT_TTS_CONFIG.copy()
    try:
        cfg = json.loads(raw)
        # 確保欄位都有
        out = DEFAULT_TTS_CONFIG.copy()
        out.update(cfg)
        # normalize types
        out["throttle_ms"] = int(out.get("throttle_ms", DEFAULT_TTS_CONFIG["throttle_ms"]))
        out["rate"] = int(out.get("rate", DEFAULT_TTS_CONFIG["rate"]))
        out["volume"] = float(out.get("volume", DEFAULT_TTS_CONFIG["volume"]))
        out["enabled"] = bool(out.get("enabled", DEFAULT_TTS_CONFIG["enabled"]))
        out["interrupt"] = bool(out.get("interrupt", DEFAULT_TTS_CONFIG["interrupt"]))
        out["throttle_mode"] = str(out.get("throttle_mode", DEFAULT_TTS_CONFIG["throttle_mode"]))
        return out
    except Exception:
        return DEFAULT_TTS_CONFIG.copy()

def save_tts_config(cfg):
    try:
        # ensure valid json serializable
        db_config_set_raw("tts_config", json.dumps(cfg, ensure_ascii=False))
    except Exception:
        pass

# ---------- TTS Settings Dialog ----------
class TTSSettingsDialog(tk.Toplevel):
    """TTS 設定視窗：enable, interrupt, throttle_ms, throttle_mode, rate, volume"""
    def __init__(self, master, on_save=None):
        super().__init__(master)
        self.title("TTS 設定")
        self.geometry("420x320")
        self.transient(master)
        self.grab_set()
        self.on_save = on_save

        self.cfg = load_tts_config()

        frm = tk.Frame(self)
        frm.pack(fill="both", expand=True, padx=12, pady=10)

        # Enabled
        self.enabled_var = tk.BooleanVar(self, value=self.cfg.get("enabled", True))
        tk.Checkbutton(frm, text="啟用發音 (TTS)", variable=self.enabled_var).pack(anchor="w", pady=(0,6))

        # Interrupt default
        self.interrupt_var = tk.BooleanVar(self, value=self.cfg.get("interrupt", True))
        tk.Checkbutton(frm, text="抽取時中斷目前播放並播放最新 (interrupt)", variable=self.interrupt_var).pack(anchor="w", pady=(0,6))

        # Throttle ms
        tk.Label(frm, text="節流 / 冷卻 (ms)：").pack(anchor="w")
        self.throttle_var = tk.StringVar(self, value=str(self.cfg.get("throttle_ms", 300)))
        tk.Entry(frm, textvariable=self.throttle_var, width=8).pack(anchor="w", pady=(0,6))

        # Throttle mode
        tk.Label(frm, text="節流模式：").pack(anchor="w")
        self.mode_var = tk.StringVar(self, value=self.cfg.get("throttle_mode", "interrupt"))
        modes = [("中斷並播放最新 (interrupt)", "interrupt"),
                 ("跳過發音 (skip)", "skip"),
                 ("延遲播放最新 (debounce)", "debounce")]
        for text, val in modes:
            tk.Radiobutton(frm, text=text, variable=self.mode_var, value=val).pack(anchor="w")

        # rate / volume
        tk.Label(frm, text="語速 (rate)：").pack(anchor="w", pady=(8,0))
        self.rate_var = tk.StringVar(self, value=str(self.cfg.get("rate", 160)))
        tk.Entry(frm, textvariable=self.rate_var, width=8).pack(anchor="w", pady=(0,6))

        tk.Label(frm, text="音量 (0.0 - 1.0)：").pack(anchor="w")
        self.volume_var = tk.StringVar(self, value=str(self.cfg.get("volume", 1.0)))
        tk.Entry(frm, textvariable=self.volume_var, width=8).pack(anchor="w", pady=(0,6))

        btnf = tk.Frame(self)
        btnf.pack(fill="x", pady=(8,6))
        tk.Button(btnf, text="保存", bg="#4CAF50", fg="white", command=self._save).pack(side="left", padx=6)
        tk.Button(btnf, text="取消", command=self.destroy).pack(side="left")

    def _save(self):
        try:
            cfg = {
                "enabled": bool(self.enabled_var.get()),
                "interrupt": bool(self.interrupt_var.get()),
                "throttle_ms": int(self.throttle_var.get()),
                "throttle_mode": self.mode_var.get(),
                "rate": int(self.rate_var.get()),
                "volume": float(self.volume_var.get())
            }
        except Exception as e:
            messagebox.showwarning("輸入錯誤", f"請檢查輸入值格式: {e}")
            return

        save_tts_config(cfg)
        if self.on_save:
            try:
                self.on_save(cfg)
            except Exception:
                pass
        messagebox.showinfo("保存成功", "TTS 設定已儲存。")
        self.destroy()

# ---------- Char Attributes Editor (簡單表格式) ----------
class CharAttributesEditor(tk.Toplevel):
    """
    字屬性編輯器：讀取 CHAR_ATTR_FILE（若不存在會建立），
    顯示字列表，選字後可編輯筆劃、五行、權重與註解，並支持匯入/匯出 JSON。
    """
    def __init__(self, master, char_attrs_path=CHAR_ATTR_FILE, on_save=None):
        super().__init__(master)
        self.title("字屬性編輯器")
        self.geometry("720x520")
        self.char_attrs_path = char_attrs_path
        self.on_save = on_save

        # load or create
        if not os.path.exists(self.char_attrs_path):
            # create empty
            try:
                with open(self.char_attrs_path, "w", encoding="utf-8") as f:
                    json.dump({}, f, ensure_ascii=False, indent=2)
            except Exception:
                pass

        self._load_attrs()

        # left: listbox of chars
        left = tk.Frame(self)
        left.pack(side="left", fill="y", padx=(10,6), pady=10)
        tk.Label(left, text="字列表：").pack(anchor="w")
        self.listbox = tk.Listbox(left, font=("Microsoft JhengHei", 14), width=6)
        self.listbox.pack(expand=True, fill="y")
        for ch in sorted(self.attrs.keys()):
            self.listbox.insert(tk.END, ch)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        # right: editor fields
        right = tk.Frame(self)
        right.pack(side="left", fill="both", expand=True, padx=(6,10), pady=10)

        tk.Label(right, text="字：").grid(row=0, column=0, sticky="w")
        self.char_label = tk.Label(right, text="", font=("Microsoft JhengHei", 18, "bold"))
        self.char_label.grid(row=0, column=1, sticky="w")

        tk.Label(right, text="筆劃：").grid(row=1, column=0, sticky="w")
        self.strokes_var = tk.StringVar(right, value="")
        tk.Entry(right, textvariable=self.strokes_var, width=10).grid(row=1, column=1, sticky="w")

        tk.Label(right, text="五行：").grid(row=2, column=0, sticky="w")
        self.wuxing_var = tk.StringVar(right, value="")
        tk.Entry(right, textvariable=self.wuxing_var, width=10).grid(row=2, column=1, sticky="w")

        tk.Label(right, text="權重 (越大越偏好)：").grid(row=3, column=0, sticky="w")
        self.weight_var = tk.StringVar(right, value="1")
        tk.Entry(right, textvariable=self.weight_var, width=10).grid(row=3, column=1, sticky="w")

        tk.Label(right, text="註解 / 字義：").grid(row=4, column=0, sticky="nw")
        self.meaning_area = scrolledtext.ScrolledText(right, width=40, height=6, wrap="word")
        self.meaning_area.grid(row=4, column=1, sticky="w", pady=(4,0))

        # buttons
        bf = tk.Frame(right)
        bf.grid(row=10, column=0, columnspan=2, pady=12)
        tk.Button(bf, text="新增字到列表", command=self._add_char).pack(side="left", padx=6)
        tk.Button(bf, text="保存變更", command=self._save_selected).pack(side="left", padx=6)
        tk.Button(bf, text="匯出 JSON", command=self._export_json).pack(side="left", padx=6)
        tk.Button(bf, text="匯入 JSON", command=self._import_json).pack(side="left", padx=6)
        tk.Button(bf, text="關閉", command=self._on_close).pack(side="left", padx=6)

        # selection state
        self.selected_char = None

    def _load_attrs(self):
        try:
            with open(self.char_attrs_path, "r", encoding="utf-8") as f:
                self.attrs = json.load(f) or {}
        except Exception:
            self.attrs = {}

    def _on_select(self, evt=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        ch = self.listbox.get(sel[0])
        self.selected_char = ch
        info = self.attrs.get(ch, {})
        self.char_label.config(text=ch)
        self.strokes_var.set(str(info.get("strokes", "")))
        self.wuxing_var.set(info.get("wuxing", ""))
        self.weight_var.set(str(info.get("weight", 1)))
        self.meaning_area.delete("1.0", tk.END)
        self.meaning_area.insert(tk.END, info.get("meaning", ""))

    def _add_char(self):
        txt = simpledialog.askstring("新增字", "請輸入要新增的字（或數個字，每個換行）:")
        if not txt:
            return
        for ch in txt.strip().splitlines():
            ch = ch.strip()
            if not ch:
                continue
            if ch in self.attrs:
                continue
            self.attrs[ch] = {"strokes": None, "wuxing": "", "weight": 1, "meaning": ""}
            self.listbox.insert(tk.END, ch)

    def _save_selected(self):
        if not self.selected_char:
            messagebox.showwarning("未選取", "請先在清單中選擇一個字")
            return
        ch = self.selected_char
        try:
            strokes = int(self.strokes_var.get()) if self.strokes_var.get().strip() else None
        except Exception:
            strokes = None
        wux = self.wuxing_var.get().strip()
        try:
            weight = float(self.weight_var.get())
        except Exception:
            weight = 1
        meaning = self.meaning_area.get("1.0", tk.END).strip()
        self.attrs[ch] = {"strokes": strokes, "wuxing": wux, "weight": weight, "meaning": meaning}
        # write back to file
        try:
            with open(self.char_attrs_path, "w", encoding="utf-8") as f:
                json.dump(self.attrs, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存成功", f"字屬性已保存至 {self.char_attrs_path}")
            if self.on_save:
                try:
                    self.on_save(self.attrs)
                except Exception:
                    pass
        except Exception as e:
            messagebox.showerror("保存失敗", f"無法保存檔案: {e}")

    def _export_json(self):
        fn = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files","*.json")], initialfile=f"char_attributes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        if not fn:
            return
        try:
            with open(fn, "w", encoding="utf-8") as f:
                json.dump(self.attrs, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("匯出成功", f"已匯出到 {fn}")
        except Exception as e:
            messagebox.showerror("匯出失敗", f"寫入錯誤: {e}")

    def _import_json(self):
        fn = filedialog.askopenfilename(filetypes=[("JSON files","*.json")])
        if not fn:
            return
        try:
            with open(fn, "r", encoding="utf-8") as f:
                data = json.load(f)
            # merge: 新字加入或覆寫
            for k, v in (data or {}).items():
                self.attrs[k] = v
            # refresh listbox
            self.listbox.delete(0, tk.END)
            for ch in sorted(self.attrs.keys()):
                self.listbox.insert(tk.END, ch)
            messagebox.showinfo("匯入成功", "已成功匯入並更新列表")
        except Exception as e:
            messagebox.showerror("匯入失敗", f"讀取錯誤: {e}")

    def _on_close(self):
        # 最後保存一次
        try:
            with open(self.char_attrs_path, "w", encoding="utf-8") as f:
                json.dump(self.attrs, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        self.destroy()

# ---------- 快捷鍵註冊函式 ----------
def register_shortcuts(root_window, app_instance):
    """
    註冊快捷鍵：
      空白鍵 => 抽取 (draw_name)
      t 或 T => 發音當前名字 (speak_current_name)
      u 或 U => 撤銷 (undo_last_draw_gui)
    root_window: tk.Tk() 或 frame
    app_instance: NameGeneratorApp 實例
    """
    def on_space(e):
        # 防止在輸入欄位中按下空白誤觸
        widget = e.widget
        if isinstance(widget, tk.Entry) or isinstance(widget, tk.Text) or isinstance(widget, scrolledtext.ScrolledText):
            return
        try:
            app_instance.draw_name()
        except Exception:
            pass

    def on_t(e):
        try:
            app_instance.speak_current_name()
        except Exception:
            pass

    def on_u(e):
        try:
            app_instance.undo_last_draw_gui()
        except Exception:
            pass

    # binding (use bind_all 方便在各種 widget focus 下都有效)
    root_window.bind_all("<space>", on_space)
    root_window.bind_all("t", on_t)
    root_window.bind_all("T", on_t)
    root_window.bind_all("u", on_u)
    root_window.bind_all("U", on_u)

# End of additions.py