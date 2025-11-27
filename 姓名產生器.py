# -*- coding: utf-8 -*-

import random
import json
import os
import sys
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog, filedialog
import shutil
import time
import re
import sqlite3
import threading
import subprocess
import platform
from tts import speak_text, stop_worker
from additions import TTSSettingsDialog, CharAttributesEditor, register_shortcuts, load_tts_config, save_tts_config
from zhuyin_ui import ZhuyinSettingsDialog, load_zhuyin_config, get_zhuyin

# --- pypinyin 可選 ---
try:
    import pypinyin
    PINYIN_ENABLED = True
    def get_pinyin_with_tone(name):
        pinyin_display_result = pypinyin.pinyin(name, style=pypinyin.Style.TONE)
        display_pinyin = " ".join([p[0] for p in pinyin_display_result])
        pinyin_num_result = pypinyin.pinyin(name, style=pypinyin.Style.TONE3)
        tones = []
        for p in pinyin_num_result:
            p_str = p[0]
            tone_num = int(p_str[-1]) if p_str and p_str[-1].isdigit() else 5
            tones.append(tone_num)
        return display_pinyin, tuple(tones)
except Exception:
    PINYIN_ENABLED = False

# ----------------- pyttsx3 支援 -----------------

def speak_current_name(self):
    text = self.current_name or self.name_var.get() or ""
    if not text or "請點擊抽取" in text or "已全部抽取完畢" in text:
        messagebox.showwarning("無法發音", "目前沒有可發音的名字，請先抽取或選擇一個名字。")
        return
    speak_text(text)  # 呼叫 pyttsx3 的非阻塞發音

# ----------------- 配置 -----------------
WORDS_FILE = 'words_list.txt'
DATA_DIR = 'name_generator_data'
STATE_FILE = None
HISTORY_FILE = None
FAVORITES_FILE = None
STATUS_FILE = None
DB_FILE = None
CHAR_ATTR_FILE = None

MASTER_WORDS = []
POOL_SIZE = 0
WORD_COUNT = 0
NAME_INDICES_CACHE = []
WORD_TO_INDEX = {}
HISTORY_RE = re.compile(r"\] - (.+?)(?: \[|$)")

# ----------------- 工具 -----------------
def atomic_write(path, data, mode='w', encoding='utf-8'):
    tmp = f"{path}.tmp"
    with open(tmp, mode, encoding=encoding) as f:
        f.write(data)
    os.replace(tmp, path)

def atomic_write_json(path, obj):
    atomic_write(path, json.dumps(obj, ensure_ascii=False, indent=2))

# ----------------- DB 支援 -----------------
def db_connect():
    global DB_FILE
    if DB_FILE is None:
        raise RuntimeError("DB_FILE unknown. Call setup_data_paths() first.")
    conn = sqlite3.connect(DB_FILE, timeout=10, isolation_level=None)
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
    except Exception:
        pass
    return conn

def init_db():
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS remaining_indices (
                idx INTEGER PRIMARY KEY
            );
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                name TEXT NOT NULL,
                tones TEXT
            );
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS favorites (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                name TEXT NOT NULL
            );
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS excluded (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                name TEXT NOT NULL
            );
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS config (
                key TEXT PRIMARY KEY,
                value TEXT
            );
        """)

def db_replace_remaining(indices):
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("BEGIN;")
        cur.execute("DELETE FROM remaining_indices;")
        cur.executemany("INSERT INTO remaining_indices(idx) VALUES (?);", ((i,) for i in indices))
        cur.execute("COMMIT;")

def db_get_remaining():
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("SELECT idx FROM remaining_indices;")
        rows = cur.fetchall()
        return [r[0] for r in rows]

def db_delete_remaining_index(idx):
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM remaining_indices WHERE idx = ?;", (idx,))

def db_insert_remaining_index(idx):
    with db_connect() as conn:
        cur = conn.cursor()
        try:
            cur.execute("INSERT OR IGNORE INTO remaining_indices(idx) VALUES (?);", (idx,))
        except Exception:
            pass

def db_insert_history(timestamp, name, tones_text=None):
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("INSERT INTO history(timestamp, name, tones) VALUES (?, ?, ?);", (timestamp, name, tones_text))

def db_get_history(limit=None):
    with db_connect() as conn:
        cur = conn.cursor()
        q = "SELECT timestamp, name, tones FROM history ORDER BY id ASC"
        if limit:
            q += f" LIMIT {limit}"
        cur.execute(q)
        return cur.fetchall()

def db_pop_last_history():
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, timestamp, name, tones FROM history ORDER BY id DESC LIMIT 1;")
        row = cur.fetchone()
        if not row:
            return None
        cur.execute("DELETE FROM history WHERE id = ?;", (row[0],))
        return row

def db_insert_favorite(timestamp, name):
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("INSERT INTO favorites(timestamp, name) VALUES (?, ?);", (timestamp, name))

def db_get_favorites():
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("SELECT timestamp, name FROM favorites ORDER BY id ASC;")
        return cur.fetchall()

def db_insert_excluded(timestamp, name):
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("INSERT INTO excluded(timestamp, name) VALUES (?, ?);", (timestamp, name))

def db_get_excluded():
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, timestamp, name FROM excluded ORDER BY id DESC;")
        return cur.fetchall()

def db_delete_excluded_by_id(excluded_id):
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM excluded WHERE id = ?;", (excluded_id,))

def db_config_get(key, default=None):
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("SELECT value FROM config WHERE key = ?;", (key,))
        row = cur.fetchone()
        return row[0] if row else default

def db_config_set(key, value):
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("INSERT INTO config(key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value;", (key, value))

# ----------------- 過濾設定 & 屬性檔 -----------------
DEFAULT_FILTER_CONFIG = {
    "unsmooth_blacklist": [(3,3), (4,4), (1,1), (2,2)],
    "probabilistic_blacklist": [],
    "reject_chance": 50
}

def load_filter_config():
    raw = db_config_get("filter_config", None)
    if raw:
        try:
            cfg = json.loads(raw)
            cfg["unsmooth_blacklist"] = [tuple(x) for x in cfg.get("unsmooth_blacklist", [])]
            cfg["probabilistic_blacklist"] = [tuple(x) for x in cfg.get("probabilistic_blacklist", [])]
            cfg["reject_chance"] = int(cfg.get("reject_chance", DEFAULT_FILTER_CONFIG["reject_chance"]))
            return cfg
        except Exception:
            pass
    return DEFAULT_FILTER_CONFIG.copy()

def save_filter_config(cfg):
    copy = {
        "unsmooth_blacklist": [list(t) for t in cfg.get("unsmooth_blacklist", [])],
        "probabilistic_blacklist": [list(t) for t in cfg.get("probabilistic_blacklist", [])],
        "reject_chance": int(cfg.get("reject_chance", DEFAULT_FILTER_CONFIG["reject_chance"]))
    }
    db_config_set("filter_config", json.dumps(copy, ensure_ascii=False))

# 字詞屬性 (char attributes)
CHAR_ATTRS = {}  # char -> {strokes:int, wuxing:str, weight:int, meaning:str}

def load_char_attributes():
    """從 CHAR_ATTR_FILE 載入字屬性；若不存在則建立範例檔案。"""
    global CHAR_ATTRS
    if CHAR_ATTR_FILE is None:
        return
    if not os.path.exists(CHAR_ATTR_FILE):
        # create a sample attribute set for the existing MASTER_WORDS if possible
        sample = {}
        for ch in MASTER_WORDS[:20]:
            sample[ch] = {"strokes": random.randint(5,15), "wuxing": random.choice(["木","火","土","金","水"]), "weight": 1, "meaning": ""}
        try:
            atomic_write_json(CHAR_ATTR_FILE, sample)
        except Exception:
            pass
        CHAR_ATTRS = sample
        return
    try:
        with open(CHAR_ATTR_FILE, 'r', encoding='utf-8') as f:
            CHAR_ATTRS = json.load(f)
    except Exception:
        CHAR_ATTRS = {}

def save_char_attributes():
    if CHAR_ATTR_FILE:
        try:
            atomic_write_json(CHAR_ATTR_FILE, CHAR_ATTRS)
        except Exception:
            pass

# ----------------- 評分系統 -----------------
def score_name(name):
    """
    簡單評分範例（可擴充）：
    - 權重（weight）
    - 筆劃平衡
    - 五行配對
    - 聲調影響（若能取得）
    """
    base = 0.0
    a = name[0]; b = name[1]
    wa = CHAR_ATTRS.get(a, {}).get("weight", 1)
    wb = CHAR_ATTRS.get(b, {}).get("weight", 1)
    base += (wa + wb) * 1.0

    # strokes
    sa = CHAR_ATTRS.get(a, {}).get("strokes")
    sb = CHAR_ATTRS.get(b, {}).get("strokes")
    if sa is not None and sb is not None:
        diff = abs(sa - sb)
        base += max(0, 3 - diff) * 0.6

    # wuxing
    wa_x = CHAR_ATTRS.get(a, {}).get("wuxing")
    wb_x = CHAR_ATTRS.get(b, {}).get("wuxing")
    if wa_x and wb_x:
        if wa_x == wb_x:
            base -= 0.5
        else:
            base += 0.4

    # pinyin/tones
    if PINYIN_ENABLED:
        try:
            _, tones = get_pinyin_with_tone(name)
            cfg = load_filter_config()
            unsmooth = [tuple(x) for x in cfg.get("unsmooth_blacklist", [])]
            prob_list = [tuple(x) for x in cfg.get("probabilistic_blacklist", [])]
            chance = int(cfg.get("reject_chance", 50))
            if tuple(tones) in unsmooth:
                base -= 5.0
            elif tuple(tones) in prob_list:
                base -= (chance / 100.0) * 2.0
            else:
                if len(tones) >=2 and tones[0] != tones[1]:
                    base += 1.2
                else:
                    base -= 0.2
        except Exception:
            pass

    base += max(0, 1.5 - ((wa + wb) / 2.0)) * 0.7

    return base

# ----------------- name/index 與核心邏輯 -----------------
def name_to_index(name):
    if not name or len(name) < 2:
        return None
    a, b = name[0], name[1]
    idx_a = WORD_TO_INDEX.get(a)
    idx_b = WORD_TO_INDEX.get(b)
    if idx_a is None or idx_b is None:
        return None
    return idx_a * WORD_COUNT + idx_b

def analyze_words_from_text(content):
    final_words = []
    for ch in content:
        if ch.strip() and not ch.isspace() and ch not in [',','，','#','\n']:
            final_words.append(ch)
    word_count = len(final_words)
    pool_size = word_count * word_count
    return word_count, pool_size

def get_drawn_indices_from_history():
    drawn_indices = set()
    rows = db_get_history()
    for timestamp, name, tones in rows:
        if name and len(name) == 2:
            idx_a = WORD_TO_INDEX.get(name[0])
            idx_b = WORD_TO_INDEX.get(name[1])
            if idx_a is not None and idx_b is not None:
                drawn_indices.add(idx_a * WORD_COUNT + idx_b)
    return drawn_indices

def get_word_frequency_stats():
    frequency = {word:0 for word in MASTER_WORDS}
    rows = db_get_history()
    for ts, name, tones in rows:
        if name and len(name)==2:
            a,b = name[0], name[1]
            if a in frequency: frequency[a]+=1
            if b in frequency: frequency[b]+=1
    return frequency

def initialize_database(reset_history=True, exclude_drawn=False):
    global NAME_INDICES_CACHE
    if POOL_SIZE == 0:
        return
    init_db()
    all_indices = list(range(POOL_SIZE))
    if exclude_drawn:
        drawn = get_drawn_indices_from_history()
        remaining = [i for i in all_indices if i not in drawn]
    else:
        remaining = all_indices
    random.shuffle(remaining)
    NAME_INDICES_CACHE = remaining.copy()
    try:
        db_replace_remaining(remaining)
    except Exception:
        with db_connect() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM remaining_indices;")
            for idx in remaining:
                try:
                    cur.execute("INSERT INTO remaining_indices(idx) VALUES (?);", (idx,))
                except Exception:
                    pass
    if reset_history:
        with db_connect() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM history;")
            cur.execute("DELETE FROM favorites;")
            cur.execute("DELETE FROM excluded;")
        reset_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            atomic_write_json(STATUS_FILE, {"last_reset": reset_time})
        except Exception:
            pass

def load_indices_cache():
    global NAME_INDICES_CACHE
    init_db()
    remaining = db_get_remaining()
    if remaining:
        NAME_INDICES_CACHE = remaining
        random.shuffle(NAME_INDICES_CACHE)
    else:
        NAME_INDICES_CACHE = []

def save_indices_cache():
    global NAME_INDICES_CACHE
    try:
        db_replace_remaining(NAME_INDICES_CACHE)
    except Exception as e:
        print("警告：無法保存索引到 DB:", e)

def get_unique_name():
    global NAME_INDICES_CACHE
    while NAME_INDICES_CACHE:
        next_index = NAME_INDICES_CACHE.pop()
        idx_a = next_index // WORD_COUNT
        idx_b = next_index % WORD_COUNT
        if idx_a >= WORD_COUNT or idx_b >= WORD_COUNT:
            messagebox.showerror("數據錯誤", "索引超出範圍，請重置數據庫。")
            return None, len(NAME_INDICES_CACHE)
        name = MASTER_WORDS[idx_a] + MASTER_WORDS[idx_b]
        tones = None
        if PINYIN_ENABLED:
            try:
                _, tones = get_pinyin_with_tone(name)
                cfg = load_filter_config()
                unsmooth = [tuple(x) for x in cfg.get("unsmooth_blacklist", [])]
                prob_list = [tuple(x) for x in cfg.get("probabilistic_blacklist", [])]
                chance = int(cfg.get("reject_chance", 50))
                if tuple(tones) in unsmooth:
                    try:
                        db_delete_remaining_index(next_index)
                    except Exception:
                        pass
                    continue
                if tuple(tones) in prob_list:
                    roll = random.randint(1,100)
                    if roll <= chance:
                        try:
                            db_delete_remaining_index(next_index)
                        except Exception:
                            pass
                        continue
            except Exception:
                pass
        try:
            db_delete_remaining_index(next_index)
        except Exception:
            pass
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            tones_text = json.dumps(list(tones)) if tones else None
            db_insert_history(timestamp, name, tones_text)
        except Exception:
            pass
        return name, len(NAME_INDICES_CACHE)
    return None, 0

def get_progress_bar(remaining):
    total = POOL_SIZE
    drawn = total - remaining
    if total == 0:
        progress_ratio = 1.0
    else:
        progress_ratio = drawn / total
    bar_length = 25
    filled_length = int(bar_length * progress_ratio)
    bar = '█' * filled_length + '░' * (bar_length - filled_length)
    percentage = f"{progress_ratio:.2%}"
    drawn_formatted = f"{drawn:,}"
    total_formatted = f"{total:,}"
    return f"進度: {drawn_formatted} / {total_formatted} ({percentage}) [{bar}]"

# ----------------- Preview Dialog (即時預覽) -----------------
class PreviewCandidatesDialog(tk.Toplevel):
    def __init__(self, master_app, sample_size=800, top_n=50):
        super().__init__(master_app.master)
        self.title("預覽高分候選名字")
        self.geometry("600x700")
        self.master_app = master_app
        self.sample_size = sample_size
        self.top_n = top_n

        tk.Label(self, text=f"從剩餘候選中隨機抽樣 {sample_size} 個，顯示 Top {top_n}（按分數排序）").pack(pady=6)

        self.text = scrolledtext.ScrolledText(self, wrap=tk.WORD, font=('Courier New', 12))
        self.text.pack(expand=True, fill='both', padx=8, pady=6)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=6)
        tk.Button(btn_frame, text="刷新", command=self.refresh, bg="#03A9F4", fg="white").pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="發音", command=self.speak_selected, bg="#9C27B0", fg="white").pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="使用選定名字", command=self.use_selected, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="關閉", command=self.destroy).pack(side=tk.LEFT, padx=6)

        self.listbox = tk.Listbox(self, height=12, font=('Courier New', 12))
        self.listbox.pack(expand=False, fill='x', padx=8, pady=(0,6))

        self.candidates = []
        self.refresh()

    def refresh(self):
        self.listbox.delete(0, tk.END)
        self.text.config(state=tk.NORMAL)
        self.text.delete('1.0', tk.END)
        remaining = db_get_remaining()
        if not remaining:
            self.text.insert(tk.END, "剩餘候選為空。請先重置數據庫。")
            self.text.config(state=tk.DISABLED)
            return
        sample = random.sample(remaining, min(self.sample_size, len(remaining)))
        scored = []
        for idx in sample:
            ia = idx // WORD_COUNT
            ib = idx % WORD_COUNT
            if ia >= WORD_COUNT or ib >= WORD_COUNT:
                continue
            name = MASTER_WORDS[ia] + MASTER_WORDS[ib]
            sc = score_name(name)
            tones_display = ""
            if PINYIN_ENABLED:
                try:
                    pinyin_display, tones = get_pinyin_with_tone(name)
                    tones_display = f"{pinyin_display} {tones}"
                except Exception:
                    tones_display = ""
            scored.append((sc, name, idx, tones_display))
        scored.sort(key=lambda x: x[0], reverse=True)
        top = scored[:self.top_n]
        self.candidates = top
        for i, (sc, name, idx, tdisp) in enumerate(top, start=1):
            self.listbox.insert(tk.END, f"{i:02d}. {name}  (score:{sc:.2f})")
            line = f"{i:02d}. {name}  score:{sc:.2f}\n    {tdisp}\n"
            self.text.insert(tk.END, line)
        self.text.config(state=tk.DISABLED)

    def speak_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning("請選擇", "請先從列表中選擇一個名字")
            return
        idx = sel[0]
        sc, name, index, tones = self.candidates[idx]
        speak_text(name)

    def use_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning("請選擇", "請先從列表中選擇一個名字")
            return
        idx = sel[0]
        sc, name, index, tones = self.candidates[idx]
        if not messagebox.askyesno("確認使用", f"您確定要使用名字 '{name}' 嗎？\n(此動作會將該組合從待抽取清單移除並記錄到歷史)"):
            return
        try:
            if index in NAME_INDICES_CACHE:
                try:
                    NAME_INDICES_CACHE.remove(index)
                except Exception:
                    pass
            db_delete_remaining_index(index)
        except Exception:
            pass
        try:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            db_insert_history(ts, name, json.dumps([]))
        except Exception:
            pass
        messagebox.showinfo("已使用", f"名字 '{name}' 已被使用並記錄。")
        self.master_app._update_progress_display(remaining=len(db_get_remaining()))
        self.refresh()

# ----------------- FilterSettingsDialog (unchanged) -----------------
class FilterSettingsDialog(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("聲調過濾設定")
        self.geometry("520x420")
        self.transient(master)
        self.grab_set()

        self.cfg = load_filter_config()

        tk.Label(self, text="不流暢（確定拒絕）的聲調組合 (每行一組，格式: 3,3)：").pack(anchor="w", padx=10, pady=(8, 0))
        self.unsmooth_box = scrolledtext.ScrolledText(self, height=6, font=('Microsoft JhengHei', 11))
        self.unsmooth_box.pack(fill="both", padx=10, pady=(0, 8))
        self.unsmooth_box.insert(tk.END, "\n".join(f"{a},{b}" for a,b in self.cfg.get("unsmooth_blacklist", [])))

        tk.Label(self, text="機率拒絕（當遇到以下組合，依拒絕機率拒絕） 每行一組：").pack(anchor="w", padx=10, pady=(0, 0))
        self.prob_box = scrolledtext.ScrolledText(self, height=4, font=('Microsoft JhengHei', 11))
        self.prob_box.pack(fill="both", padx=10, pady=(0, 8))
        self.prob_box.insert(tk.END, "\n".join(f"{a},{b}" for a,b in self.cfg.get("probabilistic_blacklist", [])))

        chance_frame = tk.Frame(self)
        chance_frame.pack(fill="x", padx=10, pady=(0, 8))
        tk.Label(chance_frame, text="拒絕機率 (%)：").pack(side=tk.LEFT)
        self.chance_var = tk.StringVar(self, value=str(self.cfg.get("reject_chance", 50)))
        tk.Entry(chance_frame, textvariable=self.chance_var, width=6).pack(side=tk.LEFT, padx=(8, 0))

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=8)
        tk.Button(btn_frame, text="保存設定", bg="#4CAF50", fg="white", command=self.save).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="取消", command=self.destroy).pack(side=tk.LEFT, padx=8)

    def save(self):
        def parse_box(text):
            items = []
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue
                parts = [p.strip() for p in line.split(',')]
                if len(parts) != 2:
                    continue
                try:
                    a = int(parts[0]); b = int(parts[1])
                    items.append((a,b))
                except Exception:
                    continue
            return items
        unsmooth = parse_box(self.unsmooth_box.get("1.0", tk.END))
        prob = parse_box(self.prob_box.get("1.0", tk.END))
        try:
            chance = int(self.chance_var.get())
            if chance <0 or chance>100: raise ValueError()
        except Exception:
            messagebox.showwarning("輸入錯誤", "拒絕機率必須是 0 到 100 的整數。")
            return
        cfg = {"unsmooth_blacklist": unsmooth, "probabilistic_blacklist": prob, "reject_chance": chance}
        save_filter_config(cfg)
        messagebox.showinfo("保存成功", "過濾設定已儲存。")
        self.destroy()

# ----------------- GUI 主應用（加入發音按鈕） -----------------
class NameGeneratorApp:
    def __init__(self, master):
        self.master = master
        master.title(f"名字抽取器 | 總組合數: {POOL_SIZE:,}")
        # 設為預設較大的視窗（若需改回其他尺寸請修改此行）
        master.geometry("900x450")
        master.config(bg='#F0F0F0')

        self.name_var = tk.StringVar(master, value="準備就緒，請點擊抽取")
        self.progress_var = tk.StringVar(master)
        self.pinyin_var = tk.StringVar(master, value="")
        self.config_stats_var = tk.StringVar(master, value="")
        self.current_name = ""

        # 確保在建立視窗之後註冊關閉處理
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 建立 UI 與載入畫面
        self._setup_ui()
        self._update_progress_display()

        # 註冊快捷鍵（若 additions.register_shortcuts 已存在）
        try:
            register_shortcuts(self.master, self)
        except Exception:
            # 非致命錯誤：若 additions 尚未加入或註冊失敗就忽略
            print("Warning: register_shortcuts failed (continuing)")

    def on_closing(self):
        """視窗關閉時，優雅停止 TTS worker 並保存狀態後關閉視窗。"""
        try:
            stop_worker()
        except Exception:
            pass

        try:
            save_indices_cache()
        except Exception:
            pass

        try:
            save_char_attributes()
        except Exception:
            pass

        try:
            self.master.destroy()
        except Exception:
            pass

    # ---------------- UI 建構（四列按鈕） ----------------
    def _setup_ui(self):
        # debug
        print("[DEBUG] _setup_ui (merged) called")

        main_bg = '#F0F0F0'
        name_fg = 'black'

        # --- 1. 清除可能殘留的舊 widget，避免重複或混用 pack/grid 導致布局異常 ---
        for attr in ['name_display', 'pinyin_display', 'progress_label', 'draw_button',
                    'batch_frame', 'buttons_container', 'status_label', 'speak_button']:
            try:
                w = getattr(self, attr, None)
                if w is not None:
                    try:
                        w.destroy()
                    except Exception:
                        pass
                    try:
                        delattr(self, attr)
                    except Exception:
                        # delattr 可能在 class 沒有屬性時失敗，忽略
                        pass
            except Exception:
                pass

        # --- 2. 建立新的緊貼 name/pinyin 布局（Label + grid） ---
        # container for name + pinyin + speak button
        np_frame = tk.Frame(self.master, bg=main_bg)
        np_frame.pack(fill='x', padx=12, pady=(8,4))

        # grid config: col0 = center (stretch), col1 = speak button (fixed)
        np_frame.grid_columnconfigure(0, weight=1)
        np_frame.grid_columnconfigure(1, weight=0)

        # Use Label for the big name to have tighter control over height
        self.name_var.set(self.name_var.get())  # 確保有值
        self.name_display = tk.Label(np_frame,
                                    textvariable=self.name_var,
                                    font=('Microsoft JhengHei', 28, 'bold'),
                                    fg=name_fg, bg=main_bg,
                                    pady=0)
        # 放中央欄（col=0），minimal pady
        self.name_display.grid(row=0, column=0, sticky='n', pady=(0,2))

        # pinyin 下方緊貼
        self.pinyin_display = tk.Label(np_frame,
                                    textvariable=self.pinyin_var,
                                    font=('Microsoft JhengHei', 12, 'italic'),
                                    fg='gray', bg=main_bg)
        self.pinyin_display.grid(row=1, column=0, sticky='n', pady=(0,2))

        # speak button 右側，跨兩列以對齊 name + pinyin
        self.speak_button = tk.Button(np_frame, text="發音 (t)", command=self.speak_current_name,
                                    font=('Microsoft JhengHei', 10), bg="#9C27B0", fg='white', width=12)
        self.speak_button.grid(row=0, column=1, rowspan=2, sticky='ne', padx=(8,0), pady=(0,2))

        # --- 3. progress 與 draw button（較小的間距） ---
        self.progress_label = tk.Label(self.master, textvariable=self.progress_var,
                                    font=('Microsoft JhengHei', 10), bg=main_bg)
        self.progress_label.pack(pady=(0,6))

        self.draw_button = tk.Button(self.master, text="抽取下一個名字 (Click)", command=self.draw_name,
                                    font=('Microsoft JhengHei', 14), bg="#4CAF50", fg='white',
                                    width=36, height=2)
        self.draw_button.pack(pady=(0,8))

        # 批量抽取區
        batch_frame = tk.Frame(self.master, bg=main_bg)
        batch_frame.pack(pady=(4,8))
        tk.Label(batch_frame, text="批量數量:", bg=main_bg, font=('Microsoft JhengHei', 10)).pack(side=tk.LEFT)
        self.batch_count_var = tk.StringVar(self.master, value='50')
        self.batch_count_entry = tk.Entry(batch_frame, textvariable=self.batch_count_var, width=6, font=('Microsoft JhengHei', 10))
        self.batch_count_entry.pack(side=tk.LEFT, padx=(6,8))
        self.batch_draw_button = tk.Button(batch_frame, text="批量抽取並預覽", command=self.batch_draw_gui,
                                           font=('Microsoft JhengHei', 10), bg="#03A9F4", fg='white', width=18)
        self.batch_draw_button.pack(side=tk.LEFT)

        # 主要功能按鈕：4x4
        btn_defs = [
            ("檢視歷史 (l)", self.view_history_gui, "#ECEFF1"),
            ("檢視收藏 (v)", self.view_favorites_gui, "#E8F5E9"),
            ("匯出歷史 (e)", self.export_history_gui, "#FFF3E0"),
            ("系統資訊 (i)", self.display_info_gui, "#F3E5F5"),

            ("重置數據庫 (r)", self.reset_database, "#FFCDD2"),
            ("檢視排除列表", self.view_excluded_names_gui, "#FFEBEE"),
            ("恢復排除組合", self.view_and_restore_excluded_gui, "#F0F4C3"),
            ("字詞庫管理 (m)", self.manage_words_gui, "#FFE0B2"),

            ("過濾設定 (g)", self.open_filter_settings, "#B2DFDB"),
            ("預覽候選 (p)", self.open_preview_dialog, "#FFCC80"),
            ("查詢名字 (s)", self.search_name_gui, "#E1F5FE"),
            ("字詞頻率 (w)", self.display_frequency_stats_gui, "#B3E5FC"),

            ("排除此組合 (x)", self.exclude_current_name_gui, "#FFCDD2"),
            ("收藏名字 (f)", self.add_favorite_gui, "#FFF9C4"),
            ("撤銷抽取 (u)", self.undo_last_draw_gui, "#D1C4E9"),
            ("檢視字詞庫 (w)", self.view_word_list_gui, "#C8E6C9"),
        ]

        buttons_container = tk.Frame(self.master, bg=main_bg)
        buttons_container.pack(padx=12, pady=(6,12), fill='x')

        rows = 4
        cols = 4
        for r in range(rows):
            row_frame = tk.Frame(buttons_container, bg=main_bg)
            row_frame.pack(fill='x', pady=3)
            for c in range(cols):
                row_frame.grid_columnconfigure(c, weight=1, uniform=f"row{r}")
            for c in range(cols):
                idx = r * cols + c
                if idx < len(btn_defs):
                    text, cmd, color = btn_defs[idx]
                    b = tk.Button(row_frame, text=text, command=cmd, font=('Microsoft JhengHei', 10),
                                  bg=color)
                    b.grid(row=0, column=c, padx=6, sticky='ew')
                else:
                    filler = tk.Label(row_frame, text="", bg=main_bg)
                    filler.grid(row=0, column=c, padx=6, sticky='ew')

        # bottom status
        status_frame = tk.Frame(self.master, bg=main_bg)
        status_frame.pack(fill='x', padx=12, pady=(6,10))
        self.status_label = tk.Label(status_frame, text="歡迎使用名字抽取器", font=('Microsoft JhengHei', 9), bg=main_bg, fg='gray')
        self.status_label.pack(side=tk.LEFT)
        

    # ----------------- TTS / UI 操作相關方法 -----------------
    def speak_current_name(self):
        text = self.current_name or self.name_var.get() or ""
        if not text or "請點擊抽取" in text or "已全部抽取完畢" in text:
            messagebox.showwarning("無法發音", "目前沒有可發音的名字，請先抽取或選擇一個名字。")
            return
        speak_text(text)

    def open_filter_settings(self):
        FilterSettingsDialog(self.master)

    def open_preview_dialog(self):
        PreviewCandidatesDialog(self)

    def view_and_restore_excluded_gui(self):
        rows = db_get_excluded()
        if not rows:
            messagebox.showwarning("提示", "目前沒有被排除的組合可供恢復。")
            return
        RestoreExcludedDialog(self, rows)

    def view_word_list_gui(self):
        w = tk.Toplevel(self.master); w.title(f"當前字詞庫（總字數: {len(MASTER_WORDS)}）"); w.geometry("450x600")
        sorted_words = sorted(MASTER_WORDS)
        WORDS_PER_LINE = 10
        lines = [" | ".join(sorted_words[i:i+WORDS_PER_LINE]) for i in range(0, len(sorted_words), WORDS_PER_LINE)]
        text_area = scrolledtext.ScrolledText(w, wrap=tk.WORD, font=('Microsoft JhengHei', 12)); text_area.insert(tk.END, "\n".join(lines)); text_area.config(state=tk.DISABLED); text_area.pack(expand=True, fill='both'); tk.Button(w, text="關閉", command=w.destroy).pack(pady=10)

    def view_excluded_names_gui(self):
        rows = db_get_excluded()
        if not rows:
            messagebox.showinfo("排除清單", "尚無排除項目。")
            return
        w = tk.Toplevel(self.master); w.title("已被排除的組合"); w.geometry("550x400")
        text_area = scrolledtext.ScrolledText(w, wrap=tk.WORD, font=('Courier New', 11), padx=10, pady=10)
        for _id, ts, name in rows:
            text_area.insert(tk.END, f"[{ts}] - {name}\n")
        text_area.config(state=tk.DISABLED); text_area.pack(expand=True, fill='both'); tk.Button(w, text="關閉", command=w.destroy).pack(pady=5)

    def _get_remaining_count(self):
        return len(NAME_INDICES_CACHE)

    def _update_progress_display(self, name=None, remaining=None, pinyin_str=None):
        if remaining is None:
            remaining = self._get_remaining_count()
        self.progress_var.set(get_progress_bar(remaining))
        if name:
            self.name_var.set(name)
        elif remaining == 0:
            self.name_var.set("已全部抽取完畢！"); self.draw_button.config(state=tk.DISABLED)
        if pinyin_str is not None:
            self.pinyin_var.set(pinyin_str)
        else:
            self.pinyin_var.set("")

    # draw_name 保留你之前的完整實作（含 TTS throttle/interrutp/debounce）
    def draw_name(self):
        MAX_ATTEMPTS = min(POOL_SIZE if POOL_SIZE else 1000, 1000)
        for attempt in range(MAX_ATTEMPTS):
            name, remaining = get_unique_name()
            if not name:
                self.current_name = ""
                self._update_progress_display(name, remaining)
                messagebox.showinfo("提示", "所有名字已抽取完畢或無合適組合！")
                return

            # 設定目前名字並嘗試複製到剪貼簿
            self.current_name = name
            try:
                self.master.clipboard_clear()
                self.master.clipboard_append(name)
            except Exception:
                pass

            # TTS 控制（參考 additions.load_tts_config 取得設定）
            try:
                cfg = load_tts_config()
            except Exception:
                cfg = None

            now_ms = int(time.time() * 1000)
            last_ms = getattr(self, "_last_speak_ts", 0)
            throttle_ms = 300
            mode = "interrupt"
            rate = 160
            volume = 1.0
            interrupt_pref = True
            enabled = True

            if cfg:
                enabled = bool(cfg.get("enabled", True))
                interrupt_pref = bool(cfg.get("interrupt", True))
                try:
                    throttle_ms = int(cfg.get("throttle_ms", 300))
                except Exception:
                    throttle_ms = 300
                mode = str(cfg.get("throttle_mode", "interrupt"))
                try:
                    rate = int(cfg.get("rate", 160))
                except Exception:
                    rate = 160
                try:
                    volume = float(cfg.get("volume", 1.0))
                except Exception:
                    volume = 1.0

            if enabled:
                elapsed = now_ms - last_ms
                if elapsed >= throttle_ms:
                    do_interrupt = interrupt_pref or (mode == "interrupt")
                    try:
                        speak_text(name, rate=rate, volume=volume, interrupt=do_interrupt)
                    except Exception:
                        pass
                    self._last_speak_ts = int(time.time() * 1000)
                else:
                    if mode == "interrupt":
                        try:
                            speak_text(name, rate=rate, volume=volume, interrupt=True)
                        except Exception:
                            pass
                        self._last_speak_ts = int(time.time() * 1000)
                    elif mode == "skip":
                        pass
                    elif mode == "debounce":
                        try:
                            if hasattr(self, "_debounce_after_id") and self._debounce_after_id:
                                try:
                                    self.master.after_cancel(self._debounce_after_id)
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        self._debounce_pending_name = name
                        def _debounced_play():
                            pending = getattr(self, "_debounce_pending_name", None)
                            if pending:
                                try:
                                    speak_text(pending, rate=rate, volume=volume, interrupt=interrupt_pref)
                                except Exception:
                                    pass
                                self._last_speak_ts = int(time.time() * 1000)
                            self._debounce_after_id = None
                            self._debounce_pending_name = None
                        try:
                            self._debounce_after_id = self.master.after(throttle_ms, _debounced_play)
                        except Exception:
                            try:
                                speak_text(name, rate=rate, volume=volume, interrupt=interrupt_pref)
                                self._last_speak_ts = int(time.time() * 1000)
                            except Exception:
                                pass
                    else:
                        try:
                            speak_text(name, rate=rate, volume=volume, interrupt=True)
                        except Exception:
                            pass
                        self._last_speak_ts = int(time.time() * 1000)

            # 顯示拼音與更新 GUI
            pinyin_str = ""
            if PINYIN_ENABLED:
                try:
                    pinyin_str, _ = get_pinyin_with_tone(name)
                except Exception:
                    pinyin_str = ""
            self._update_progress_display(name, remaining, pinyin_str)
            return

        # 若嘗試耗盡仍未找到
        self.current_name = ""
        self.pinyin_var.set("")
        remaining = self._get_remaining_count()
        self._update_progress_display(name="連續過濾失敗", remaining=remaining)
        messagebox.showwarning("抽取失敗", f"連續 {MAX_ATTEMPTS} 次抽取都遇到過濾情形，請重置或調整過濾規則。")
        self.draw_button.config(state=tk.DISABLED)
    
    def display_info_gui(self):
        """
        顯示系統資訊視窗（字詞庫、進度、檔案狀態、TTS 狀態等）。
        將此方法貼到 NameGeneratorApp 類中（與其它 view_* 函式並列）。
        依賴外部全域變數/函式： WORDS_FILE, WORD_COUNT, POOL_SIZE, db_get_remaining, STATUS_FILE, DB_FILE, CHAR_ATTR_FILE
        """
        try:
            remaining_count = self._get_remaining_count()
        except Exception:
            try:
                remaining_count = len(db_get_remaining())
            except Exception:
                remaining_count = "N/A"

        drawn_count = "N/A"
        try:
            if isinstance(POOL_SIZE, int) and isinstance(remaining_count, int):
                drawn_count = POOL_SIZE - remaining_count
            else:
                drawn_count = "N/A"
        except Exception:
            drawn_count = "N/A"

        last_reset = "N/A"
        try:
            if os.path.exists(STATUS_FILE):
                with open(STATUS_FILE, 'r', encoding='utf-8') as sf:
                    status_data = json.load(sf)
                    last_reset = status_data.get("last_reset", last_reset)
        except Exception:
            pass

        db_exists = os.path.exists(DB_FILE)
        char_attr_exists = os.path.exists(CHAR_ATTR_FILE)
        words_exists = os.path.exists(WORDS_FILE)

        # TTS status (best-effort)
        tts_status = "未知"
        try:
            import pyttsx3
            tts_status = "pyttsx3 可用"
        except Exception:
            # fall back to check if speak_text is present
            if 'speak_text' in globals():
                tts_status = "TTS 函式可用"
            else:
                tts_status = "TTS 未安裝或不可用"

        info_lines = []
        info_lines.append("[一、字詞庫資訊]")
        info_lines.append(f"  - 字詞庫檔案: {WORDS_FILE} ({'存在' if words_exists else '遺失'})")
        info_lines.append(f"  - 總字數 (N): {WORD_COUNT:,}" if isinstance(WORD_COUNT, int) else f"  - 總字數 (N): {WORD_COUNT}")
        info_lines.append(f"  - 總組合數 (N x N): {POOL_SIZE:,}" if isinstance(POOL_SIZE, int) else f"  - 總組合數 (N x N): {POOL_SIZE}")

        info_lines.append("\n[二、抽取進度]")
        info_lines.append(f"  - 已抽取: {drawn_count if isinstance(drawn_count, int) else drawn_count}")
        info_lines.append(f"  - 剩餘數量: {remaining_count if isinstance(remaining_count, int) else remaining_count}")

        info_lines.append("\n[三、檔案狀態]")
        info_lines.append(f"  - DB: {'✅ 存在' if db_exists else '❌ 遺失'} ({DB_FILE})")
        info_lines.append(f"  - 字屬性檔: {'✅ 存在' if char_attr_exists else '❌ 遺失'} ({CHAR_ATTR_FILE})")
        info_lines.append(f"  - 上次重置時間: {last_reset}")

        info_lines.append("\n[四、系統環境 & TTS]")
        try:
            py_ver = sys.version.splitlines()[0]
        except Exception:
            py_ver = sys.version if 'sys' in globals() else "unknown"
        info_lines.append(f"  - Python: {py_ver}")
        info_lines.append(f"  - 平台: {platform.system()} {platform.release()}")
        info_lines.append(f"  - TTS: {tts_status}")

        # 顯示視窗
        info = "\n".join(info_lines)
        try:
            messagebox.showinfo("系統狀態與資訊 (INFO)", info)
        except Exception:
            # fallback: open a simple Toplevel with scrolledtext
            w = tk.Toplevel(self.master)
            w.title("系統狀態與資訊 (INFO)")
            txt = scrolledtext.ScrolledText(w, wrap=tk.WORD, width=80, height=24)
            txt.pack(expand=True, fill='both', padx=10, pady=10)
            txt.insert(tk.END, info)
            txt.config(state=tk.DISABLED)
            tk.Button(w, text="關閉", command=w.destroy).pack(pady=(0,10))
    
    def display_frequency_stats_gui(self):
        """
        顯示字詞被抽取次數統計視窗（從 get_word_frequency_stats() 取得資料）。
        將此方法貼到 NameGeneratorApp 類中（與其它 view_* 方法並列）。
        支援匯出為 CSV 檔案。
        """
        try:
            stats = get_word_frequency_stats()
        except Exception as e:
            messagebox.showerror("錯誤", f"無法計算字詞頻率: {e}")
            return

        if not stats or all(count == 0 for count in stats.values()):
            messagebox.showinfo("字詞抽取頻率", "尚未抽取任何名字，或歷史記錄中沒有與當前字詞庫匹配的字。")
            return

        sorted_stats = sorted(stats.items(), key=lambda item: item[1], reverse=True)
        total_draws = sum(stats.values()) // 2  # 每個名字有兩個字，所以這樣估算抽取名字數
        header = f"【字詞抽取頻率統計】\n\n總抽取名字數 (估算): {total_draws:,} 個（雙字計算）\n\n"

        # 建立視窗
        w = tk.Toplevel(self.master)
        w.title("字詞抽取頻率統計")
        w.geometry("420x560")
        w.transient(self.master)
        w.grab_set()

        # 文本區
        text_widget = scrolledtext.ScrolledText(w, wrap=tk.WORD, font=('Courier New', 10))
        text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=(10,6))

        # 組裝顯示內容（只列出有次數的字）
        lines = [header]
        for word, count in sorted_stats:
            if count > 0:
                lines.append(f"  - 字 '{word}': 被抽取 {count} 次")
        content = "\n".join(lines)
        text_widget.insert(tk.END, content)
        text_widget.config(state=tk.DISABLED)

        # 按鈕列（匯出 CSV / 關閉）
        def export_csv():
            try:
                fn = filedialog.asksaveasfilename(defaultextension=".csv",
                                                filetypes=[("CSV files", "*.csv"), ("Text files", "*.txt")],
                                                initialfile=f"frequency_stats_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                                title="匯出字詞頻率為 CSV")
                if not fn:
                    return
                import csv
                with open(fn, "w", encoding="utf-8", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Word", "Count"])
                    for word, count in sorted_stats:
                        if count > 0:
                            writer.writerow([word, count])
                messagebox.showinfo("匯出成功", f"字詞頻率已匯出至:\n{fn}")
            except Exception as e:
                messagebox.showerror("匯出失敗", f"匯出過程發生錯誤: {e}")

        btn_frame = tk.Frame(w)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0,10))
        tk.Button(btn_frame, text="匯出 CSV", command=export_csv, bg="#2196F3", fg="white").pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="關閉", command=w.destroy).pack(side=tk.RIGHT, padx=6)

        # 自動滾到頂部
        try:
            text_widget.see("1.0")
        except Exception:
            pass

    def search_name_gui(self):
        """
        查詢兩字名字的狀態並顯示相關資訊。
        將此方法貼到 NameGeneratorApp 類中（與其他 view_* 方法並列）。
        會檢查：
        - 名字是否為兩個字
        - 是否在字詞庫中（每個字是否存在 MASTER_WORDS）
        - 是否已被抽取（透過 history）
        - 若啟用 pypinyin，會顯示拼音與聲調
        """
        name = simpledialog.askstring("名字查詢", "請輸入要查詢的兩個漢字名字:")
        if not name:
            return
        name = name.strip()
        if len(name) != 2:
            messagebox.showwarning("查詢失敗", "名字必須為兩個漢字。")
            return

        a, b = name[0], name[1]
        in_pool = (a in MASTER_WORDS) and (b in MASTER_WORDS)
        idx = name_to_index(name) if in_pool else None

        # 檢查抽取狀態（從 history）
        drawn_status = "❌ 待抽取"
        try:
            rows = db_get_history()
            for ts, n, tones in rows:
                if n == name:
                    drawn_status = "✅ 已抽取"
                    break
        except Exception:
            # 若讀取失敗，設為未知
            drawn_status = "未知（歷史讀取失敗）"

        # 拼音/聲調資訊（若可用）
        pinyin_display = ""
        tones_display = ""
        if PINYIN_ENABLED:
            try:
                pinyin_display, tones = get_pinyin_with_tone(name)
                tones_display = f" 聲調: {tones}"
            except Exception:
                pinyin_display = ""
                tones_display = ""

        # 組出訊息
        title = f"名字查詢：{name}"
        msg_lines = []
        msg_lines.append(f"名字：{name}")
        msg_lines.append(f"是否在字詞庫中：{'✅ 是' if in_pool else '❌ 否'}")
        if in_pool and idx is not None:
            try:
                msg_lines.append(f"索引 (a_index * N + b_index)：{idx}")
                msg_lines.append(f"總字數: {WORD_COUNT:,}，總組合: {POOL_SIZE:,}")
            except Exception:
                pass
        msg_lines.append(f"抽取狀態：{drawn_status}")
        if pinyin_display:
            msg_lines.append(f"拼音：{pinyin_display}{tones_display}")

        # 顯示結果
        messagebox.showinfo(title, "\n".join(msg_lines))
    
    def export_history_gui(self):
        """
        匯出歷史紀錄到文字或 CSV 檔案。
        建議將此方法貼到 NameGeneratorApp 類中（與其它 view_* 函式平行）。
        會使用 db_get_history() 取得 (timestamp, name, tones) 三元組。
        """
        try:
            rows = db_get_history()
        except Exception as e:
            messagebox.showerror("匯出失敗", f"讀取歷史資料時發生錯誤：{e}")
            return

        if not rows:
            messagebox.showwarning("匯出失敗", "歷史記錄為空，無法匯出。")
            return

        # 提供兩種格式：.txt 或 .csv
        initial = f"Export_History_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        fn = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files","*.txt"), ("CSV files","*.csv")],
            initialfile=initial + ".txt",
            title="匯出歷史記錄為..."
        )
        if not fn:
            return

        try:
            lower = fn.lower()
            if lower.endswith(".csv"):
                # 匯出 CSV：欄位 Timestamp, Name, Tones (tones 做為 JSON 或逗號分隔)
                import csv
                with open(fn, "w", encoding="utf-8", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Timestamp", "Name", "Tones"])
                    for ts, name, tones in rows:
                        tcell = ""
                        if tones:
                            # tones 在 DB 中可能是 JSON 字串，嘗試解析
                            try:
                                parsed = json.loads(tones)
                                if isinstance(parsed, (list, tuple)):
                                    tcell = ",".join(str(x) for x in parsed)
                                else:
                                    tcell = str(parsed)
                            except Exception:
                                tcell = str(tones)
                        writer.writerow([ts, name, tcell])
            else:
                # 預設匯出為純文字，每行一筆
                with open(fn, "w", encoding="utf-8") as f:
                    for ts, name, tones in rows:
                        line = f"[{ts}] - {name}"
                        if tones:
                            try:
                                tlist = json.loads(tones)
                                line += " [" + ",".join(map(str, tlist)) + "]"
                            except Exception:
                                line += f" [{tones}]"
                        f.write(line + "\n")
            messagebox.showinfo("匯出成功", f"歷史記錄已成功匯出至:\n{fn}")
        except Exception as e:
            messagebox.showerror("匯出失敗", f"匯出過程發生錯誤: {e}")
        
    # 以下方法大致維持原先實作（保留行為）
    def undo_last_draw_gui(self):
        last = db_pop_last_history()
        if not last:
            messagebox.showwarning("無法撤銷", "歷史記錄為空或無法讀取。")
            return
        if len(last) >= 4:
            _id, ts, name, tones = last
        else:
            messagebox.showwarning("撤銷警告", "歷史解析錯誤，請手動檢查。"); return
        if not name or len(name) !=2:
            messagebox.showwarning("撤銷警告", f"名字長度異常：{name}"); return
        idx = name_to_index(name)
        if idx is None:
            messagebox.showwarning("撤銷警告", f"字詞不在庫中：{name}"); return
        if idx not in NAME_INDICES_CACHE:
            NAME_INDICES_CACHE.append(idx)
        try:
            db_insert_remaining_index(idx)
            db_replace_remaining(NAME_INDICES_CACHE)
        except Exception:
            pass
        messagebox.showinfo("成功", f"已撤銷抽取：{name}")

    def batch_draw_gui(self):
        try:
            count = int(self.batch_count_var.get())
            if count<=0 or count>1000:
                messagebox.showwarning("警告","批量抽取數量必須是 1 到 1000 之間的整數。"); return
        except ValueError:
            messagebox.showwarning("警告","請輸入有效的批量抽取數量。"); return
        drawn_names=[]
        draw_limit = min(count, self._get_remaining_count())
        if draw_limit==0:
            messagebox.showinfo("提示","剩餘待抽取名字數量為 0。"); return
        for _ in range(draw_limit):
            name, remaining = get_unique_name()
            if name:
                drawn_names.append(name)
            else:
                break
        final_remaining = self._get_remaining_count()
        self.current_name = drawn_names[-1] if drawn_names else ""
        self._update_progress_display(name=self.current_name, remaining=final_remaining)
        self._display_batch_results(drawn_names, draw_limit)

    def manage_words_gui(self):
        if hasattr(self, '_batch_dialog') and getattr(self, '_batch_dialog', None) and self._batch_dialog.winfo_exists():
            self._batch_dialog.lift(); return
        self._batch_dialog = BatchWordManagerDialog(self.master, self)
        self.master.wait_window(self._batch_dialog)

    def _display_batch_results(self, names, draw_count):
        results_window = tk.Toplevel(self.master); results_window.title(f"批量抽取結果 ({len(names)} 個)"); results_window.geometry("400x550")
        header_text = f"成功抽取 {len(names)} 個名字。\n"; header_label = tk.Label(results_window, text=header_text, font=('Microsoft JhengHei', 10, 'bold'), pady=5); header_label.pack()
        text_widget = scrolledtext.ScrolledText(results_window, wrap=tk.WORD, font=('Courier New', 12)); text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=(0,10))
        output_content = ""
        for i, name in enumerate(names):
            output_content += f"{i+1:03d}. {name}\n"
        text_widget.insert(tk.END, output_content); text_widget.config(state=tk.DISABLED)
        def copy_to_clipboard():
            full_text = "\n".join(names); results_window.clipboard_clear(); results_window.clipboard_append(full_text); messagebox.showinfo("複製成功", f"共 {len(names)} 個名字已複製到剪貼簿！")
        copy_button = tk.Button(results_window, text=f"複製 {len(names)} 個結果到剪貼簿", command=copy_to_clipboard, font=('Microsoft JhengHei', 10), bg="#2196F3", fg='white'); copy_button.pack(pady=(0,10), padx=10, fill=tk.X)

    def reset_database(self, show_message=True):
        if show_message:
            if not messagebox.askyesno("警告", "您確定要重置數據庫嗎？\n\n注意：重置將會清空當前未抽取的索引列表。"):
                return
            reset_type = messagebox.askquestion("選擇重置模式", "【標準重置 (Yes)】：清空所有歷史/收藏\n【智慧重置 (No)】：保留歷史/收藏，排除已抽組合", type=messagebox.YESNOCANCEL, default=messagebox.YES)
            if reset_type == messagebox.CANCEL: return
            is_standard_reset = (reset_type == messagebox.YES)
        else:
            is_standard_reset = True
        if is_standard_reset:
            initialize_database(reset_history=True, exclude_drawn=False); reset_message = "標準重置完成"
        else:
            initialize_database(reset_history=False, exclude_drawn=True); reset_message = "智慧重置完成"
        final_remaining_count = self._get_remaining_count(); self.current_name = ""; self._update_progress_display(remaining=final_remaining_count); self.draw_button.config(state=tk.NORMAL)
        if show_message:
            messagebox.showinfo(reset_message, f"數據庫已重置。\n\n總字數: {WORD_COUNT} 個\n總組合數: {POOL_SIZE:,} 個\n剩餘待抽取數量: {final_remaining_count:,} 個"); self.name_var.set("重置完成，請點擊抽取")

    def exclude_current_name_gui(self):
        name_to_exclude = self.current_name
        if not name_to_exclude or name_to_exclude=="已全部抽取完畢！" or len(name_to_exclude)!=2:
            messagebox.showwarning("無法排除","請先抽取一個名字，且名字必須為兩個漢字。"); return
        if not messagebox.askyesno("確認排除", f"您確定要將名字 '{name_to_exclude}' 從待抽取列表永久排除嗎？"):
            return
        try:
            idx = name_to_index(name_to_exclude)
            if idx is None: raise ValueError("字詞庫中不存在該字")
            try:
                NAME_INDICES_CACHE.remove(idx)
            except Exception:
                pass
            try:
                db_delete_remaining_index(idx)
            except Exception:
                pass
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); db_insert_excluded(ts, name_to_exclude)
            try: db_replace_remaining(NAME_INDICES_CACHE)
            except Exception: pass
            self.current_name=""; self._update_progress_display(name=f"'{name_to_exclude}' 已永久排除", remaining=len(NAME_INDICES_CACHE)); messagebox.showinfo("排除成功", f"名字 '{name_to_exclude}' 已從待抽取組合中永久移除。")
        except ValueError:
            messagebox.showerror("錯誤","當前字詞庫中不包含此名字的字詞，無法排除。")
        except Exception as e:
            messagebox.showerror("錯誤", f"執行排除操作時發生錯誤: {e}")

    def add_favorite_gui(self):
        if self.current_name:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                db_insert_favorite(ts, self.current_name)
                messagebox.showinfo("收藏成功", f"'{self.current_name}' 已加入收藏清單。")
            except Exception as e:
                messagebox.showerror("錯誤", f"無法寫入收藏: {e}")
        else:
            messagebox.showwarning("提示","請先抽取一個名字再進行收藏。")
    
    def view_favorites_gui(self):
        """
        顯示收藏清單的視窗（含播放、複製、刪除與匯出功能）。
        會使用資料庫中的 favorites 表（fetch id,timestamp,name）來展示並允許刪除單筆。
        將這個方法貼到你的 NameGeneratorApp 類中（與其他 view_* 函式平行）。
        """
        try:
            # 取得帶 id 的 favorites 資料
            with db_connect() as conn:
                cur = conn.cursor()
                cur.execute("SELECT id, timestamp, name FROM favorites ORDER BY id ASC;")
                rows = cur.fetchall()
        except Exception as e:
            messagebox.showerror("錯誤", f"無法讀取收藏資料庫：{e}")
            return

        w = tk.Toplevel(self.master)
        w.title("收藏名字清單")
        w.geometry("520x520")
        w.transient(self.master)
        w.grab_set()

        # listbox 與 scrollbar
        list_frame = tk.Frame(w)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=('Courier New', 10))
        scrollbar.config(command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        # 將 rows（id,ts,name）存到本地變數以便操作
        fav_rows = list(rows or [])

        def refresh_list():
            nonlocal fav_rows
            try:
                with db_connect() as conn:
                    cur = conn.cursor()
                    cur.execute("SELECT id, timestamp, name FROM favorites ORDER BY id ASC;")
                    fav_rows = cur.fetchall()
            except Exception as e:
                messagebox.showerror("錯誤", f"無法讀取收藏資料庫：{e}")
                fav_rows = []
            listbox.delete(0, tk.END)
            for i, (fid, ts, name) in enumerate(fav_rows, start=1):
                display = f"{i:03d}. [{ts}] - {name}"
                listbox.insert(tk.END, display)

        refresh_list()

        # 按鈕功能
        def play_selected():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("請選擇", "請先從列表中選擇一個收藏項目。")
                return
            idx = sel[0]
            _, ts, name = fav_rows[idx]
            try:
                speak_text(name)
            except Exception:
                messagebox.showwarning("發音失敗", "發音模組不可用或發生錯誤。")

        def copy_selected():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("請選擇", "請先從列表中選擇一個收藏項目。")
                return
            idx = sel[0]
            _, ts, name = fav_rows[idx]
            try:
                w.clipboard_clear()
                w.clipboard_append(name)
                messagebox.showinfo("複製成功", f"已複製：{name}")
            except Exception as e:
                messagebox.showerror("複製失敗", f"複製到剪貼簿時發生錯誤：{e}")

        def delete_selected():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("請選擇", "請先從列表中選擇一個收藏項目。")
                return
            idx = sel[0]
            fid, ts, name = fav_rows[idx]
            if not messagebox.askyesno("確認刪除", f"確定要刪除收藏：\n[{ts}] - {name}？"):
                return
            try:
                with db_connect() as conn:
                    cur = conn.cursor()
                    # 使用 id 刪除最安全
                    cur.execute("DELETE FROM favorites WHERE id = ?;", (fid,))
                refresh_list()
                messagebox.showinfo("刪除成功", "已刪除選定的收藏。")
                # 若在主畫面顯示收藏數量或其他資訊，可在此更新
            except Exception as e:
                messagebox.showerror("刪除失敗", f"刪除收藏時發生錯誤：{e}")

        def export_all():
            try:
                if not fav_rows:
                    messagebox.showinfo("匯出", "收藏列表為空，無資料可匯出。")
                    return
                fn = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")], initialfile=f"favorites_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
                if not fn:
                    return
                with open(fn, 'w', encoding='utf-8') as f:
                    for fid, ts, name in fav_rows:
                        f.write(f"[{ts}] - {name}\n")
                messagebox.showinfo("匯出成功", f"已匯出至：{fn}")
            except Exception as e:
                messagebox.showerror("匯出失敗", f"匯出過程發生錯誤：{e}")

        # 按鈕列
        btn_frame = tk.Frame(w)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0,10))
        tk.Button(btn_frame, text="發音選取", command=play_selected, bg="#9C27B0", fg="white").pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="複製選取", command=copy_selected, bg="#2196F3", fg="white").pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="刪除選取", command=delete_selected, bg="#D32F2F", fg="white").pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="匯出全部", command=export_all).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="關閉", command=w.destroy).pack(side=tk.RIGHT, padx=6)

        # 欄位鍵綁定（Enter 播放、Delete 刪除）
        def on_key(event):
            if event.keysym in ("Return", "KP_Enter"):
                play_selected()
            elif event.keysym == "Delete":
                delete_selected()

        listbox.bind("<Double-Button-1>", lambda e: play_selected())
        listbox.bind("<Key>", on_key)

        # 讓清單自動選到第一項（若存在）
        if fav_rows:
            try:
                listbox.selection_set(0)
                listbox.see(0)
            except Exception:
                pass
    
    

    def view_history_gui(self):
        rows = db_get_history()
        w = tk.Toplevel(self.master); w.title("抽取歷史紀錄"); w.geometry("450x600")
        text_widget = scrolledtext.ScrolledText(w, wrap=tk.WORD, font=('Courier New', 10)); text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        if not rows: text_widget.insert(tk.END, "尚無歷史記錄。")
        else:
            for ts, name, tones in rows:
                t_display = f"[{ts}] - {name}"
                if tones:
                    try:
                        t_list = json.loads(tones); t_display += f" [{','.join(map(str,t_list))}]"
                    except Exception:
                        pass

# ----------------- 啟動邏輯 -----------------
def setup_data_paths():
    global WORDS_FILE, STATE_FILE, HISTORY_FILE, FAVORITES_FILE, STATUS_FILE, DATA_DIR, DB_FILE, CHAR_ATTR_FILE
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    WORDS_FILE = os.path.join(DATA_DIR, 'words_list.txt')
    STATE_FILE = os.path.join(DATA_DIR, 'name_indices.json')
    HISTORY_FILE = os.path.join(DATA_DIR, 'drawn_history.txt')
    FAVORITES_FILE = os.path.join(DATA_DIR, 'favorites.txt')
    STATUS_FILE = os.path.join(DATA_DIR, 'system_status.json')
    DB_FILE = os.path.join(DATA_DIR, 'name_generator.sqlite3')
    CHAR_ATTR_FILE = os.path.join(DATA_DIR, 'char_attributes.json')

def load_master_words():
    global MASTER_WORDS, POOL_SIZE, WORD_COUNT, WORD_TO_INDEX
    if not os.path.exists(WORDS_FILE):
        try:
            with open(WORDS_FILE, 'w', encoding='utf-8') as f:
                f.write("愛\n麗\n雅\n靜\n")
                f.write("風\n雲\n月\n星\n")
            messagebox.showerror("錯誤：找不到字詞庫", f"找不到字詞庫檔案 '{WORDS_FILE}'。\n程式已在資料夾中為您創建範本檔案，請編輯後再次運行程式。")
            sys.exit(1)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法創建字詞庫檔案: {e}"); sys.exit(1)
    try:
        with open(WORDS_FILE, 'r', encoding='utf-8') as f: content = f.read()
        final_words = []
        for ch in content:
            if ch.strip() and not ch.isspace() and ch not in [',','，','#','\n']:
                final_words.append(ch)
        if not final_words:
            messagebox.showerror("錯誤", f"字詞庫檔案 '{WORDS_FILE}' 內容為空。請編輯後重新運行。"); sys.exit(1)
        MASTER_WORDS = final_words
        WORD_COUNT = len(MASTER_WORDS); POOL_SIZE = WORD_COUNT * WORD_COUNT
        WORD_TO_INDEX = {word:i for i,word in enumerate(MASTER_WORDS)}
    except Exception as e:
        messagebox.showerror("錯誤", f"加載字詞庫時發生錯誤: {e}"); sys.exit(1)

# ----------------- 補充：簡化的 RestoreExcludedDialog 和 BatchWordManagerDialog ------------
class RestoreExcludedDialog(tk.Toplevel):
    def __init__(self, master_app, excluded_rows):
        super().__init__(master_app.master)
        self.title("恢復已排除組合")
        self.geometry("600x420")
        self.master_app = master_app
        self.excluded_rows = excluded_rows
        self.listbox = tk.Listbox(self, selectmode=tk.MULTIPLE, width=80, height=18, font=('Courier New',11))
        self.listbox.pack(padx=10, pady=6, fill='both', expand=True)
        for _id, ts, name in excluded_rows:
            self.listbox.insert(tk.END, f"[{ts}] - {name}  (id:{_id})")
        bf = tk.Frame(self); bf.pack(pady=6)

        # 新增：發音按鈕（對選取的已排除名字發音）
        tk.Button(bf, text="發音", command=self.speak_selected, bg="#9C27B0", fg="white").pack(side=tk.LEFT, padx=8)

        tk.Button(bf, text="恢復選定組合", command=self.restore_selected, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=8)
        tk.Button(bf, text="關閉", command=self.destroy).pack(side=tk.LEFT, padx=8)

    def speak_selected(self):
        """對 Listbox 中選取的第一筆已排除名字進行發音（若有 tts 可用）。"""
        try:
            sel = self.listbox.curselection()
            if not sel:
                messagebox.showwarning("請選擇", "請先選擇一個已排除的名字以供發音。")
                return
            i = sel[0]
            _id, ts, name = self.excluded_rows[i]
            # speak_text 需在檔案頂部匯入： from tts import speak_text
            try:
                speak_text(name)
            except NameError:
                messagebox.showwarning("發音功能未載入", "找不到發音模組 (speak_text)。請確認已加入 tts.py 並在檔案頂端 import。")
        except Exception as e:
            messagebox.showerror("發音錯誤", f"嘗試發音時發生錯誤: {e}")

    def restore_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning("提示","請至少選擇一個組合進行恢復。"); return
        restored = 0
        for i in sel:
            _id, ts, name = self.excluded_rows[i]
            idx = name_to_index(name)
            if idx is None:
                continue
            if idx not in NAME_INDICES_CACHE:
                NAME_INDICES_CACHE.append(idx)
            try:
                db_insert_remaining_index(idx)
                db_delete_excluded_by_id(_id)
                restored += 1
            except Exception:
                pass
        try:
            db_replace_remaining(NAME_INDICES_CACHE)
        except Exception:
            pass
        messagebox.showinfo("成功", f"已恢復 {restored} 個組合。")
        self.master_app._update_progress_display(remaining=len(NAME_INDICES_CACHE))
        self.destroy()

class BatchWordManagerDialog(tk.Toplevel):
    def __init__(self, master, app_instance):
        super().__init__(master)
        self.title("字詞庫內容管理")
        self.geometry("500x600")
        self.app_instance = app_instance
        self.word_edit_area = scrolledtext.ScrolledText(self, wrap=tk.WORD, font=('Microsoft JhengHei',12), padx=10, pady=10)
        initial_content = "\n".join(MASTER_WORDS)
        self.word_edit_area.insert(tk.END, initial_content)
        self.word_edit_area.pack(expand=True, fill='both')
        bf = tk.Frame(self); bf.pack(pady=10)

        # 新增：對選取文字發音（方便試聽單字）
        tk.Button(bf, text="發音選取字", command=self.speak_selection, font=('Microsoft JhengHei',10), bg="#9C27B0", fg="white", width=12).pack(side=tk.LEFT, padx=6)

        tk.Button(bf, text="保存並重新啟動", command=self.save_changes, font=('Microsoft JhengHei',10), bg="#2196F3", fg="white", width=15).pack(side=tk.LEFT, padx=10)
        tk.Button(bf, text="取消", command=self.destroy, font=('Microsoft JhengHei',10), width=15).pack(side=tk.LEFT, padx=10)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.transient(master); self.grab_set()

    def speak_selection(self):
        """讀取編輯區目前選取文字並發音（通常選一個字或幾個字）。"""
        try:
            sel_text = self.word_edit_area.selection_get().strip()
            if not sel_text:
                messagebox.showwarning("無選取", "請先在編輯區選取一個字以發音。")
                return
            # speak_text 需在檔案頂部匯入： from tts import speak_text
            try:
                speak_text(sel_text)
            except NameError:
                messagebox.showwarning("發音功能未載入", "找不到發音模組 (speak_text)。請確認已加入 tts.py 並在檔案頂端 import。")
        except tk.TclError:
            messagebox.showwarning("無選取", "請先在編輯區選取一個字以發音。")
        except Exception as e:
            messagebox.showerror("發音錯誤", f"嘗試發音時發生錯誤: {e}")

    def save_changes(self):
        raw = self.word_edit_area.get('1.0', tk.END)
        clean_words = sorted({w.strip() for w in raw.split('\n') if w.strip()})
        if len(clean_words) < 2:
            messagebox.showerror("保存失敗", "字詞庫至少需要兩個字。"); return
        try:
            atomic_write(WORDS_FILE, "\n".join(clean_words) + "\n")
        except Exception as e:
            messagebox.showerror("保存失敗", f"寫入檔案時發生錯誤:\n{e}"); return
        messagebox.showinfo("保存成功", f"字詞庫已更新，共 {len(clean_words)} 個字，程式將重新啟動。")
        self.master.quit()
        python = sys.executable
        try:
            os.execl(python, python, *sys.argv)
        except Exception as e:
            messagebox.showerror("重啟失敗", f"無法自動重啟，請手動重新啟動：\n{e}")
            
if __name__ == "__main__":
    setup_data_paths()
    load_master_words()
    load_char_attributes()
    init_db()
    if not os.path.exists(DB_FILE) or (not db_get_remaining() and POOL_SIZE > 0):
        initialize_database(reset_history=True)
    else:
        load_indices_cache()
    root = tk.Tk()
    # NOTE: integrate complete NameGeneratorApp implementation (above is truncated with pass for brevity)
    app = NameGeneratorApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()