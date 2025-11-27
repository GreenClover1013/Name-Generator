# -*- coding: utf-8 -*-
import random
import json
import os
import sys
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog
import shutil
import time
import re
import sqlite3

# --- [新增] Pinyin 庫與語音學分析 ---
try:
    import pypinyin
    PINYIN_ENABLED = True

    def get_pinyin_with_tone(name):
        """轉換名字為拼音（帶聲調符號）和數字聲調元組。"""
        pinyin_display_result = pypinyin.pinyin(name, style=pypinyin.Style.TONE)
        display_pinyin = " ".join([p[0] for p in pinyin_display_result])

        pinyin_num_result = pypinyin.pinyin(name, style=pypinyin.Style.TONE3)
        tones = []
        for p in pinyin_num_result:
            p_str = p[0]
            tone_num = int(p_str[-1]) if p_str and p_str[-1].isdigit() else 5
            tones.append(tone_num)

        return display_pinyin, tuple(tones)

except ImportError:
    PINYIN_ENABLED = False
    print("Warning: pypinyin library not found. Pinyin analysis is disabled.")
# --- [結束新增] Pinyin 庫與語音學分析 ---

# ----------------- 配置 -----------------
WORDS_FILE = 'words_list.txt'
STATE_FILE = 'name_indices.json'  # legacy file still used for compatibility in case
HISTORY_FILE = 'drawn_history.txt'  # legacy export path
FAVORITES_FILE = 'favorites.txt'  # legacy export
STATUS_FILE = 'system_status.json'
DATA_DIR = 'name_generator_data'
DB_FILE = None  # will be set in setup_data_paths()

MASTER_WORDS = []
POOL_SIZE = 0
WORD_COUNT = 0
# 索引緩存在記憶體中
NAME_INDICES_CACHE = []
WORD_TO_INDEX = {}
HISTORY_RE = re.compile(r"\] - (.+?)(?: \[|$)")
# ------------------------------------------

# ----------------- 工具函數 -----------------
def atomic_write(path, data, mode='w', encoding='utf-8'):
    """安全地以原子方式寫入檔案（先寫入臨時檔再替換）。"""
    tmp = f"{path}.tmp"
    with open(tmp, mode, encoding=encoding) as f:
        f.write(data)
    os.replace(tmp, path)

def atomic_write_json(path, obj):
    atomic_write(path, json.dumps(obj, ensure_ascii=False), mode='w', encoding='utf-8')

# ----------------- SQLite 支援 -----------------
def db_connect():
    """返回 sqlite3 連線（自動建立資料庫檔案）。"""
    global DB_FILE
    if DB_FILE is None:
        raise RuntimeError("DB_FILE 尚未設定，請先呼叫 setup_data_paths()")
    conn = sqlite3.connect(DB_FILE, timeout=10, isolation_level=None)
    # 使用 WAL 模式提升並發寫入效能
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
    except Exception:
        pass
    return conn

def init_db():
    """建立必要的資料表（如果不存在）。"""
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
    """以 transaction 的方式替換 remaining_indices 表內容。"""
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
        return row  # (id, timestamp, name, tones)

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

# ----------------- Pinyin 過濾設定 -----------------
CONFIG_FILE = None  # 同樣在 setup_data_paths 設定

# 預設過濾設定
DEFAULT_FILTER_CONFIG = {
    "unsmooth_blacklist": [(3, 3), (4, 4), (1, 1), (2, 2)],
    "probabilistic_blacklist": [],  # list of tuples
    "reject_chance": 50  # percentage (0-100) applied to probabilistic_blacklist entries
}

def load_filter_config():
    """從 DB（或檔案）載入過濾設定並回傳 dict。"""
    # 優先從 DB config 表載入
    raw = db_config_get("filter_config", None)
    if raw:
        try:
            cfg = json.loads(raw)
            # convert lists of lists to tuples
            cfg["unsmooth_blacklist"] = [tuple(x) for x in cfg.get("unsmooth_blacklist", [])]
            cfg["probabilistic_blacklist"] = [tuple(x) for x in cfg.get("probabilistic_blacklist", [])]
            cfg["reject_chance"] = int(cfg.get("reject_chance", DEFAULT_FILTER_CONFIG["reject_chance"]))
            return cfg
        except Exception:
            pass

    # fallback: return defaults
    return DEFAULT_FILTER_CONFIG.copy()

def save_filter_config(cfg):
    """保存過濾設定到 DB config 表。"""
    # convert tuples to lists for JSON
    copy = {
        "unsmooth_blacklist": [list(t) for t in cfg.get("unsmooth_blacklist", [])],
        "probabilistic_blacklist": [list(t) for t in cfg.get("probabilistic_blacklist", [])],
        "reject_chance": int(cfg.get("reject_chance", DEFAULT_FILTER_CONFIG["reject_chance"]))
    }
    db_config_set("filter_config", json.dumps(copy, ensure_ascii=False))

def is_tone_combination_smooth(tones):
    """
    使用可配置的黑名單與機率過濾。
    - 如果 tones 在 unsmooth_blacklist -> 不流暢 (False)
    - 如果 tones 在 probabilistic_blacklist -> 以 reject_chance 的機率拒絕
    - 否則視為流暢
    """
    if not isinstance(tones, tuple) or len(tones) != 2:
        return True

    cfg = load_filter_config()
    unsmooth = [tuple(x) for x in cfg.get("unsmooth_blacklist", [])]
    prob_list = [tuple(x) for x in cfg.get("probabilistic_blacklist", [])]
    chance = int(cfg.get("reject_chance", 50))

    if tones in unsmooth:
        return False
    if tones in prob_list:
        # 以機率拒絕
        roll = random.randint(1, 100)
        return False if roll <= chance else True

    return True

# ==================== 核心數據函數 ====================
def name_to_index(name):
    """將名字 (例如 '雅靜') 轉換回其在 MASTER_WORDS 中的索引。"""
    if not name or len(name) < 2:
        return None
    a, b = name[0], name[1]
    idx_a = WORD_TO_INDEX.get(a)
    idx_b = WORD_TO_INDEX.get(b)
    if idx_a is None or idx_b is None:
        return None
    return idx_a * WORD_COUNT + idx_b

def analyze_words_from_text(content):
    """從文字內容中分析總字數和總組合數，不影響全域狀態。"""
    final_words = []
    for ch in content:
        if ch.strip() and not ch.isspace() and ch not in [',', '，', '#', '\n']:
            final_words.append(ch)
    word_count = len(final_words)
    pool_size = word_count * word_count
    return word_count, pool_size

def get_drawn_indices_from_history():
    """從 history table 解析已抽取索引（只計算仍在字詞庫中的字）。"""
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
    """計算每個字詞在歷史記錄中的抽取頻率。"""
    frequency = {word: 0 for word in MASTER_WORDS}
    rows = db_get_history()
    for timestamp, name, tones in rows:
        if name and len(name) == 2:
            a, b = name[0], name[1]
            if a in frequency:
                frequency[a] += 1
            if b in frequency:
                frequency[b] += 1
    return frequency

def initialize_database(reset_history=True, exclude_drawn=False):
    """初始化數據庫，將索引寫入緩存和 sqlite 資料庫。"""
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

    # 替換 DB 內容
    try:
        db_replace_remaining(remaining)
    except Exception:
        # 容錯：嘗試逐筆插入
        with db_connect() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM remaining_indices;")
            for idx in remaining:
                try:
                    cur.execute("INSERT INTO remaining_indices(idx) VALUES (?);", (idx,))
                except Exception:
                    pass

    if reset_history:
        # 清空 history/favorites/excluded
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
    """程式啟動時，從 sqlite 載入索引到記憶體緩存。"""
    global NAME_INDICES_CACHE
    init_db()
    remaining = db_get_remaining()
    if remaining:
        NAME_INDICES_CACHE = remaining
        random.shuffle(NAME_INDICES_CACHE)
    else:
        NAME_INDICES_CACHE = []

def save_indices_cache():
    """程式退出前，將記憶體緩存寫回 sqlite（以替換方式）。"""
    global NAME_INDICES_CACHE
    try:
        db_replace_remaining(NAME_INDICES_CACHE)
    except Exception as e:
        print(f"警告：無法保存狀態到 DB: {e}")

def get_unique_name():
    """
    從記憶體緩存中抽取索引，並在抽取時檢查拼音流暢度。
    抽取成功才會寫入 history 與從 DB 刪除 remaining_indices 中對應的索引。
    """
    global NAME_INDICES_CACHE

    while NAME_INDICES_CACHE:
        next_index = NAME_INDICES_CACHE.pop()
        idx_a = next_index // WORD_COUNT
        idx_b = next_index % WORD_COUNT

        if idx_a >= WORD_COUNT or idx_b >= WORD_COUNT:
            messagebox.showerror("數據錯誤", "索引超出範圍，請重置數據庫。")
            return None, len(NAME_INDICES_CACHE)

        name = MASTER_WORDS[idx_a] + MASTER_WORDS[idx_b]

        # 拼音檢查（若有安裝）
        tones = None
        if PINYIN_ENABLED:
            try:
                _, tones = get_pinyin_with_tone(name)
                if not is_tone_combination_smooth(tones):
                    # 不流暢：跳過（不記錄到歷史，也不重新加入）
                    print(f"Skipped name (unsmooth/prob filtered): {name} (tones: {tones})")
                    # 同時也從 DB 中移除該索引（如果尚未移除）
                    try:
                        db_delete_remaining_index(next_index)
                    except Exception:
                        pass
                    continue
            except Exception:
                # 若拼音轉換失敗，仍接受此名字
                pass

        # 接受：從 DB 刪除此索引（若尚未），並寫入 history
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
    """計算並返回格式化的進度條字串"""
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

# ----------------- GUI 對話框：調整過濾規則 -----------------
class FilterSettingsDialog(tk.Toplevel):
    """允許使用者自定義不流暢黑名單、機率性黑名單和拒絕機率。"""
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
        self.unsmooth_box.insert(tk.END, "\n".join(f"{a},{b}" for a, b in self.cfg.get("unsmooth_blacklist", [])))

        tk.Label(self, text="機率拒絕（當遇到以下組合，依拒絕機率拒絕） 每行一組：").pack(anchor="w", padx=10, pady=(0, 0))
        self.prob_box = scrolledtext.ScrolledText(self, height=4, font=('Microsoft JhengHei', 11))
        self.prob_box.pack(fill="both", padx=10, pady=(0, 8))
        self.prob_box.insert(tk.END, "\n".join(f"{a},{b}" for a, b in self.cfg.get("probabilistic_blacklist", [])))

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
                    items.append((a, b))
                except Exception:
                    continue
            return items

        unsmooth = parse_box(self.unsmooth_box.get("1.0", tk.END))
        prob = parse_box(self.prob_box.get("1.0", tk.END))
        try:
            chance = int(self.chance_var.get())
            if chance < 0 or chance > 100:
                raise ValueError()
        except Exception:
            messagebox.showwarning("輸入錯誤", "拒絕機率必須是 0 到 100 的整數。")
            return

        cfg = {
            "unsmooth_blacklist": unsmooth,
            "probabilistic_blacklist": prob,
            "reject_chance": chance
        }
        save_filter_config(cfg)
        messagebox.showinfo("保存成功", "過濾設定已儲存。")
        self.destroy()

# ==================== Tkinter 介面邏輯 ====================

class RestoreExcludedDialog(tk.Toplevel):
    # 這個對話框會列出 excluded 表中的項目，允許恢復至 remaining_indices
    def __init__(self, master_app, excluded_rows):
        super().__init__(master_app.master)
        self.title("恢復已排除組合")
        self.geometry("600x500")
        self.master_app = master_app
        self.excluded_rows = excluded_rows  # list of (id, timestamp, name)
        self.selected_ids = []

        tk.Label(self, text="請選擇要恢復的已排除組合：").pack(pady=5)

        self.listbox = tk.Listbox(self, selectmode=tk.MULTIPLE, height=20, width=80, font=('Courier New', 11))
        self.listbox.pack(padx=10, pady=5, fill="both", expand=True)

        for row in excluded_rows:
            eid, ts, name = row
            self.listbox.insert(tk.END, f"[{ts}] - {name}  (id:{eid})")

        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)
        tk.Button(button_frame, text="恢復選定組合", command=self.restore_selected, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="關閉", command=self.destroy).pack(side=tk.LEFT, padx=10)

    def restore_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "請至少選擇一個組合進行恢復。")
            return

        restored = 0
        for i in sel:
            eid, ts, name = self.excluded_rows[i]
            idx = name_to_index(name)
            if idx is None:
                continue
            # 加回記憶體緩存與 DB（若尚未存在）
            if idx not in NAME_INDICES_CACHE:
                NAME_INDICES_CACHE.append(idx)
            db_insert_remaining_index(idx)
            db_delete_excluded_by_id(eid)
            restored += 1

        # 同步 DB 狀態
        try:
            db_replace_remaining(NAME_INDICES_CACHE)
        except Exception:
            pass

        messagebox.showinfo("成功", f"已恢復 {restored} 個組合。")
        self.master_app._update_progress_display(remaining=len(NAME_INDICES_CACHE))
        self.destroy()

class BatchWordManagerDialog(tk.Toplevel):
    """字詞庫內容管理對話框 (模態)"""
    def __init__(self, master, app_instance):
        super().__init__(master)
        self.title("字詞庫內容管理")
        self.geometry("500x600")
        self.app_instance = app_instance

        self.word_edit_area = scrolledtext.ScrolledText(self,
                                                        wrap=tk.WORD,
                                                        font=('Microsoft JhengHei', 12),
                                                        padx=10, pady=10)

        initial_content = "\n".join(MASTER_WORDS)
        self.word_edit_area.insert(tk.END, initial_content)
        self.word_edit_area.pack(expand=True, fill='both')

        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="保存並重新啟動",
                  command=self.save_changes,
                  font=('Microsoft JhengHei', 10),
                  bg="#2196F3", fg="white", width=15).pack(side=tk.LEFT, padx=10)

        tk.Button(button_frame, text="取消",
                  command=self.destroy,
                  font=('Microsoft JhengHei', 10),
                  width=15).pack(side=tk.LEFT, padx=10)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.transient(master)
        self.grab_set()

    def save_changes(self):
        raw = self.word_edit_area.get('1.0', tk.END)
        clean_words = sorted({w.strip() for w in raw.split('\n') if w.strip()})

        if len(clean_words) < 2:
            messagebox.showerror("保存失敗", "字詞庫至少需要兩個字。")
            return

        try:
            atomic_write(WORDS_FILE, "\n".join(clean_words) + "\n")
        except Exception as e:
            messagebox.showerror("保存失敗", f"寫入檔案時發生錯誤:\n{e}")
            return

        messagebox.showinfo(
            "保存成功",
            f"字詞庫已更新，共 {len(clean_words)} 個字，程式將重新啟動。"
        )

        self.master.quit()
        python = sys.executable
        try:
            os.execl(python, python, *sys.argv)
        except Exception as e:
            messagebox.showerror(
                "重啟失敗",
                f"無法自動重啟，請手動重新啟動：\n{e}"
            )

def setup_data_paths():
    """設定數據檔案的實際路徑並確保資料夾存在。"""
    global WORDS_FILE, STATE_FILE, HISTORY_FILE, FAVORITES_FILE, STATUS_FILE, DATA_DIR, DB_FILE, CONFIG_FILE

    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    WORDS_FILE = os.path.join(DATA_DIR, 'words_list.txt')
    STATE_FILE = os.path.join(DATA_DIR, 'name_indices.json')
    HISTORY_FILE = os.path.join(DATA_DIR, 'drawn_history.txt')
    FAVORITES_FILE = os.path.join(DATA_DIR, 'favorites.txt')
    STATUS_FILE = os.path.join(DATA_DIR, 'system_status.json')
    DB_FILE = os.path.join(DATA_DIR, 'name_generator.sqlite3')
    CONFIG_FILE = os.path.join(DATA_DIR, 'config.json')

def get_excluded_names():
    """從 excluded 表讀取所有已排除的名字（返回顯示文字與數量）。"""
    rows = db_get_excluded()
    if not rows:
        return "尚無被排除的組合。", 0
    lines = [f"[{ts}] - {name}" for (_id, ts, name) in rows]
    return "\n".join(lines), len(rows)

# ==================== Tkinter 主應用 ====================

class NameGeneratorApp:
    def __init__(self, master):
        self.master = master
        master.title(f"名字抽取器 | 總組合數: {POOL_SIZE:,}")
        master.geometry("780x500")
        master.config(bg='#F0F0F0')

        self.name_var = tk.StringVar(master, value="準備就緒，請點擊抽取")
        self.progress_var = tk.StringVar(master)
        self.pinyin_var = tk.StringVar(master, value="")
        self.config_stats_var = tk.StringVar(master, value="")
        self.current_name = ""

        self._setup_ui()
        self._update_progress_display()

    def on_closing(self):
        save_indices_cache()
        self.master.destroy()

    def _setup_ui(self):
        main_bg = '#F0F0F0'
        name_fg = 'black'

        self.name_display = tk.Entry(self.master,
                                     textvariable=self.name_var,
                                     font=('Microsoft JhengHei', 24, 'bold'),
                                     fg=name_fg,
                                     bg=main_bg,
                                     justify='center',
                                     state='readonly',
                                     readonlybackground=main_bg,
                                     relief='flat',
                                     width=15)
        self.name_display.pack(pady=10, padx=10)

        self.pinyin_display = tk.Label(self.master,
                                      textvariable=self.pinyin_var,
                                      font=('Microsoft JhengHei', 12, 'italic'),
                                      fg='gray',
                                      bg=main_bg,
                                      pady=0)
        self.pinyin_display.pack(pady=(0, 5))

        self.progress_label = tk.Label(self.master,
                                      textvariable=self.progress_var,
                                      font=('Microsoft JhengHei', 10),
                                      bg=main_bg,
                                      pady=10)
        self.progress_label.pack()

        self.draw_button = tk.Button(self.master,
                                     text="抽取下一個名字 (Click)",
                                     command=self.draw_name,
                                     font=('Microsoft JhengHei', 14),
                                     bg="#4CAF50",
                                     fg='white',
                                     width=30,
                                     height=2)
        self.draw_button.pack(pady=10)

        batch_frame = tk.Frame(self.master, bg=main_bg)
        batch_frame.pack(pady=5)

        tk.Label(batch_frame, text="批量數量:", bg=main_bg, font=('Microsoft JhengHei', 10)).pack(side=tk.LEFT, padx=(10, 2))
        self.batch_count_var = tk.StringVar(self.master, value='50')
        self.batch_count_entry = tk.Entry(batch_frame,
                                          textvariable=self.batch_count_var,
                                          width=6,
                                          font=('Microsoft JhengHei', 10))
        self.batch_count_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.batch_draw_button = tk.Button(batch_frame,
                                           text="批量抽取並預覽",
                                           command=self.batch_draw_gui,
                                           font=('Microsoft JhengHei', 10),
                                           bg="#03A9F4",
                                           fg='white',
                                           width=18)
        self.batch_draw_button.pack(side=tk.LEFT, padx=5)

        button_frame_top = tk.Frame(self.master, bg=main_bg)
        button_frame_top.pack(pady=5)

        self.history_button = tk.Button(button_frame_top, text="檢視歷史 (l)", command=self.view_history_gui, font=('Microsoft JhengHei', 10), width=12)
        self.history_button.pack(side=tk.LEFT, padx=5)

        self.view_favorite_button = tk.Button(button_frame_top, text="檢視收藏 (v)", command=self.view_favorites_gui, font=('Microsoft JhengHei', 10), width=12)
        self.view_favorite_button.pack(side=tk.LEFT, padx=5)

        self.export_button = tk.Button(button_frame_top, text="匯出歷史 (e)", command=self.export_history_gui, font=('Microsoft JhengHei', 10), width=12)
        self.export_button.pack(side=tk.LEFT, padx=5)

        self.info_button = tk.Button(button_frame_top, text="系統資訊 (i)", command=self.display_info_gui, font=('Microsoft JhengHei', 10), width=12)
        self.info_button.pack(side=tk.LEFT, padx=5)

        button_frame_middle = tk.Frame(self.master, bg=main_bg)
        button_frame_middle.pack(pady=5)

        self.reset_button = tk.Button(button_frame_middle, text="重置數據庫 (r)", command=self.reset_database, font=('Microsoft JhengHei', 10), width=12, bg="#D32F2F", fg="white")
        self.reset_button.pack(side=tk.LEFT, padx=5)

        self.excluded_button = tk.Button(button_frame_middle,
                                         text="檢視排除列表",
                                         command=self.view_excluded_names_gui,
                                         font=('Microsoft JhengHei', 10),
                                         width=12)
        self.excluded_button.pack(side=tk.LEFT, padx=5)

        self.restore_button = tk.Button(button_frame_top,
                                        text="恢復排除組合",
                                        command=self.view_and_restore_excluded_gui,
                                        font=('Microsoft JhengHei', 10),
                                        width=12,
                                        bg="#CDDC39")
        self.restore_button.pack(side=tk.LEFT, padx=5)

        self.manage_button = tk.Button(button_frame_middle,
                                       text="字詞庫管理 (m)",
                                       command=self.manage_words_gui,
                                       font=('Arial', 10),
                                       width=12,
                                       bg="#FFE0B2")
        self.manage_button.pack(side=tk.LEFT, padx=5)

        self.filter_button = tk.Button(button_frame_middle,
                                       text="過濾設定 (g)",
                                       command=self.open_filter_settings,
                                       font=('Microsoft JhengHei', 10),
                                       width=12,
                                       bg="#B2DFDB")
        self.filter_button.pack(side=tk.LEFT, padx=5)

        self.search_button = tk.Button(button_frame_middle, text="查詢名字 (s)", command=self.search_name_gui, font=('Microsoft JhengHei', 10), width=12)
        self.search_button.pack(side=tk.LEFT, padx=5)

        self.frequency_button = tk.Button(button_frame_middle,
                                          text="字詞頻率 (w)",
                                          command=self.display_frequency_stats_gui,
                                          font=('Microsoft JhengHei', 10),
                                          width=12,
                                          bg="#B3E5FC")
        self.frequency_button.pack(side=tk.LEFT, padx=5)

        button_frame_last = tk.Frame(self.master, bg=main_bg)
        button_frame_last.pack(pady=5)

        self.exclude_button = tk.Button(button_frame_last,
                                        text="排除此組合 (x)",
                                        command=self.exclude_current_name_gui,
                                        font=('Microsoft JhengHei', 10),
                                        width=12,
                                        bg="#FFCDD2")
        self.exclude_button.pack(side=tk.LEFT, padx=10)

        self.favorite_button = tk.Button(button_frame_last, text="收藏名字 (f)", command=self.add_favorite_gui, font=('Microsoft JhengHei', 10), width=12)
        self.favorite_button.pack(side=tk.LEFT, padx=10)

        self.undo_button = tk.Button(button_frame_last,
                                     text="撤銷抽取 (u)",
                                     command=self.undo_last_draw_gui,
                                     font=('Microsoft JhengHei', 10),
                                     width=12,
                                     bg="#D1C4E9")
        self.undo_button.pack(side=tk.LEFT, padx=10)

        button_frame_last2 = tk.Frame(self.master, bg=main_bg)
        button_frame_last2.pack(pady=5)

        self.view_words_button = tk.Button(button_frame_last2,
                                           text="檢視字詞庫 (w)",
                                           command=self.view_word_list_gui,
                                           font=('Microsoft JhengHei', 10),
                                           width=15,
                                           bg="#C8E6C9")
        self.view_words_button.pack(side=tk.LEFT, padx=5)

    def open_filter_settings(self):
        FilterSettingsDialog(self.master)

    def view_and_restore_excluded_gui(self):
        rows = db_get_excluded()
        if not rows:
            messagebox.showwarning("提示", "目前沒有被排除的組合可供恢復。")
            return
        RestoreExcludedDialog(self, rows)

    def view_word_list_gui(self):
        word_window = tk.Toplevel(self.master)
        word_window.title(f"當前字詞庫（總字數: {len(MASTER_WORDS)}）")
        word_window.geometry("450x600")

        sorted_words = sorted(MASTER_WORDS)
        WORDS_PER_LINE = 10
        lines = [" | ".join(sorted_words[i:i + WORDS_PER_LINE]) for i in range(0, len(sorted_words), WORDS_PER_LINE)]

        text_area = scrolledtext.ScrolledText(word_window, wrap=tk.WORD, font=('Microsoft JhengHei', 12))
        text_area.insert(tk.END, "\n".join(lines))
        text_area.config(state=tk.DISABLED)
        text_area.pack(expand=True, fill='both')

        tk.Button(word_window, text="關閉", command=word_window.destroy).pack(pady=10)

    def view_excluded_names_gui(self):
        text, count = get_excluded_names()
        title = f"已被排除的組合列表 ({count:,} 個)"
        w = tk.Toplevel(self.master)
        w.title(title)
        w.geometry("550x400")
        text_area = scrolledtext.ScrolledText(w, wrap=tk.WORD, font=('Courier New', 11), padx=10, pady=10)
        text_area.insert(tk.END, text)
        text_area.config(state=tk.DISABLED)
        text_area.pack(expand=True, fill='both')
        tk.Button(w, text="關閉", command=w.destroy, width=10).pack(pady=5)

    def _get_remaining_count(self):
        return len(NAME_INDICES_CACHE)

    def _update_progress_display(self, name=None, remaining=None, pinyin_str=None):
        if remaining is None:
            remaining = self._get_remaining_count()
        progress_line = get_progress_bar(remaining)
        self.progress_var.set(progress_line)
        if name:
            self.name_var.set(name)
        elif remaining == 0:
            self.name_var.set("已全部抽取完畢！")
            self.draw_button.config(state=tk.DISABLED)
        if pinyin_str is not None:
            self.pinyin_var.set(pinyin_str)
        else:
            self.pinyin_var.set("")

    def draw_name(self):
        MAX_ATTEMPTS = min(POOL_SIZE if POOL_SIZE else 1000, 1000)
        for attempt in range(MAX_ATTEMPTS):
            name, remaining = get_unique_name()
            if not name:
                self.current_name = ""
                self._update_progress_display(name, remaining)
                messagebox.showinfo("提示", "所有名字已抽取完畢或無合適組合！")
                return

            self.current_name = name
            try:
                self.master.clipboard_clear()
                self.master.clipboard_append(name)
            except Exception:
                pass

            pinyin_str = ""
            if PINYIN_ENABLED:
                try:
                    pinyin_str, _ = get_pinyin_with_tone(name)
                except Exception:
                    pinyin_str = ""

            self._update_progress_display(name, remaining, pinyin_str)
            return

        self.current_name = ""
        self.pinyin_var.set("")
        remaining = self._get_remaining_count()
        self._update_progress_display(name="連續過濾失敗", remaining=remaining)
        messagebox.showwarning("抽取失敗", f"連續 {MAX_ATTEMPTS} 次抽取都遇到過濾情形，請重置或調整過濾規則。")
        self.draw_button.config(state=tk.DISABLED)

    def undo_last_draw_gui(self):
        # 1. 從 DB pop 最後一筆 history
        last = db_pop_last_history()
        if not last:
            messagebox.showwarning("無法撤銷", "歷史記錄為空或無法讀取。")
            return
        # last = (id, timestamp, name, tones) 但 db_pop_last_history 回傳 SELECT id,timestamp,name,tones
        # 在實作中我們設計為回傳該 tuple
        # 若回傳 (id, timestamp, name, tones)
        if len(last) >= 4:
            _id, ts, name, tones = last
        else:
            # fallback 保守處理
            messagebox.showwarning("撤銷警告", "歷史解析錯誤，請手動檢查。")
            return

        if not name or len(name) != 2:
            messagebox.showwarning("撤銷警告", f"名字長度異常：{name}")
            return

        idx = name_to_index(name)
        if idx is None:
            messagebox.showwarning("撤銷警告", f"字詞不在庫中：{name}")
            return

        # 加回 remaining（記憶體與 DB）
        if idx not in NAME_INDICES_CACHE:
            NAME_INDICES_CACHE.append(idx)
        try:
            db_insert_remaining_index(idx)
        except Exception:
            pass

        try:
            db_replace_remaining(NAME_INDICES_CACHE)
        except Exception:
            pass

        messagebox.showinfo("成功", f"已撤銷抽取：{name}")

    def batch_draw_gui(self):
        try:
            count = int(self.batch_count_var.get())
            if count <= 0 or count > 1000:
                messagebox.showwarning("警告", "批量抽取數量必須是 1 到 1000 之間的整數。")
                return
        except ValueError:
            messagebox.showwarning("警告", "請輸入有效的批量抽取數量。")
            return

        drawn_names = []
        draw_limit = min(count, self._get_remaining_count())
        if draw_limit == 0:
            messagebox.showinfo("提示", "剩餘待抽取名字數量為 0。")
            return

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
            self._batch_dialog.lift()
            return
        self._batch_dialog = BatchWordManagerDialog(self.master, self)
        self.master.wait_window(self._batch_dialog)

    def _display_batch_results(self, names, draw_count):
        results_window = tk.Toplevel(self.master)
        results_window.title(f"批量抽取結果 ({len(names)} 個)")
        results_window.geometry("400x550")

        header_text = f"成功抽取 {len(names)} 個名字。\n"
        header_label = tk.Label(results_window, text=header_text, font=('Microsoft JhengHei', 10, 'bold'), pady=5)
        header_label.pack()

        text_widget = scrolledtext.ScrolledText(results_window, wrap=tk.WORD, font=('Courier New', 12))
        text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=(0, 10))

        output_content = ""
        for i, name in enumerate(names):
            output_content += f"{i+1:03d}. {name}\n"

        text_widget.insert(tk.END, output_content)
        text_widget.config(state=tk.DISABLED)

        def copy_to_clipboard():
            full_text = "\n".join(names)
            results_window.clipboard_clear()
            results_window.clipboard_append(full_text)
            messagebox.showinfo("複製成功", f"共 {len(names)} 個名字已複製到剪貼簿！")

        copy_button = tk.Button(results_window,
                                text=f"複製 {len(names)} 個結果到剪貼簿",
                                command=copy_to_clipboard,
                                font=('Microsoft JhengHei', 10),
                                bg="#2196F3",
                                fg='white')
        copy_button.pack(pady=(0, 10), padx=10, fill=tk.X)

    def reset_database(self, show_message=True):
        if show_message:
            if not messagebox.askyesno("警告", "您確定要重置數據庫嗎？\n\n注意：重置將會清空當前未抽取的索引列表。"):
                return
            reset_type = messagebox.askquestion(
                "選擇重置模式",
                "您希望執行哪種重置模式？\n\n"
                "【標準重置 (Yes)】：清空所有歷史/收藏，從所有組合中重新生成索引。\n"
                "【智慧重置 (No)】：保留歷史/收藏，從所有組合中排除**已抽取的**組合，只對剩餘組合生成索引。",
                type=messagebox.YESNOCANCEL,
                default=messagebox.YES
            )
            if reset_type == messagebox.CANCEL:
                return
            is_standard_reset = (reset_type == messagebox.YES)
        else:
            is_standard_reset = True

        if is_standard_reset:
            initialize_database(reset_history=True, exclude_drawn=False)
            reset_message = "標準重置完成"
        else:
            initialize_database(reset_history=False, exclude_drawn=True)
            reset_message = "智慧重置完成"

        final_remaining_count = self._get_remaining_count()
        self.current_name = ""
        self._update_progress_display(remaining=final_remaining_count)
        self.draw_button.config(state=tk.NORMAL)

        if show_message:
            messagebox.showinfo(reset_message,
                                f"數據庫已重置。\n\n總字數: {WORD_COUNT} 個\n總組合數: {POOL_SIZE:,} 個\n剩餘待抽取數量: {final_remaining_count:,} 個")
            self.name_var.set("重置完成，請點擊抽取")

    def exclude_current_name_gui(self):
        name_to_exclude = self.current_name
        if not name_to_exclude or name_to_exclude == "已全部抽取完畢！" or len(name_to_exclude) != 2:
            messagebox.showwarning("無法排除", "請先抽取一個名字，且名字必須為兩個漢字。")
            return
        if not messagebox.askyesno("確認排除", f"您確定要將名字 '{name_to_exclude}' 從待抽取列表**永久排除**嗎？\n\n注意：這將減少總待抽取組合數。"):
            return

        try:
            idx = name_to_index(name_to_exclude)
            if idx is None:
                raise ValueError("字詞庫中不存在該字")
            # 從記憶體緩存移除（若存在）
            try:
                NAME_INDICES_CACHE.remove(idx)
            except ValueError:
                pass
            # 從 DB 中刪除
            try:
                db_delete_remaining_index(idx)
            except Exception:
                pass
            # 記錄到 excluded 表
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            db_insert_excluded(ts, name_to_exclude)
            # 同步 DB
            try:
                db_replace_remaining(NAME_INDICES_CACHE)
            except Exception:
                pass

            self.current_name = ""
            self._update_progress_display(name=f"'{name_to_exclude}' 已永久排除", remaining=len(NAME_INDICES_CACHE))
            messagebox.showinfo("排除成功", f"名字 '{name_to_exclude}' 已從待抽取組合中永久移除。")
        except ValueError:
            messagebox.showerror("錯誤", "當前字詞庫中不包含此名字的字詞，無法排除。")
        except Exception as e:
            messagebox.showerror("錯誤", f"執行排除操作時發生錯誤: {e}")

    def add_favorite_gui(self):
        if self.current_name:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                db_insert_favorite(timestamp, self.current_name)
                messagebox.showinfo("收藏成功", f"'{self.current_name}' 已加入收藏清單。")
            except Exception as e:
                messagebox.showerror("錯誤", f"無法寫入收藏: {e}")
        else:
            messagebox.showwarning("提示", "請先抽取一個名字再進行收藏。")

    def view_history_gui(self):
        rows = db_get_history()
        w = tk.Toplevel(self.master)
        w.title("抽取歷史紀錄")
        w.geometry("450x600")
        text_widget = scrolledtext.ScrolledText(w, wrap=tk.WORD, font=('Courier New', 10))
        text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        if not rows:
            text_widget.insert(tk.END, "尚無歷史記錄。")
        else:
            for ts, name, tones in rows:
                t_display = f"[{ts}] - {name}"
                if tones:
                    try:
                        t_list = json.loads(tones)
                        t_display += f" [{','.join(map(str,t_list))}]"
                    except Exception:
                        pass
                text_widget.insert(tk.END, t_display + "\n")
            text_widget.see(tk.END)
        text_widget.config(state=tk.DISABLED)

    def view_favorites_gui(self):
        rows = db_get_favorites()
        w = tk.Toplevel(self.master)
        w.title("收藏名字清單")
        w.geometry("400x500")
        text_widget = scrolledtext.ScrolledText(w, wrap=tk.WORD, font=('Courier New', 10))
        text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        if not rows:
            text_widget.insert(tk.END, "尚無收藏記錄。")
        else:
            for ts, name in rows:
                text_widget.insert(tk.END, f"[{ts}] - {name}\n")
            text_widget.see(tk.END)
        text_widget.config(state=tk.DISABLED)

    def export_history_gui(self):
        rows = db_get_history()
        if not rows:
            messagebox.showwarning("匯出失敗", "歷史記錄為空，無法匯出。")
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_filename = f"Export_History_{timestamp}.txt"
        try:
            with open(export_filename, 'w', encoding='utf-8') as f:
                for ts, name, tones in rows:
                    line = f"[{ts}] - {name}"
                    if tones:
                        try:
                            tlist = json.loads(tones)
                            line += f" [{','.join(map(str,tlist))}]"
                        except Exception:
                            pass
                    f.write(line + "\n")
            messagebox.showinfo("匯出成功", f"歷史記錄已成功匯出至:\n{export_filename}")
        except Exception as e:
            messagebox.showerror("匯出失敗", f"檔案寫入錯誤: {e}")

    def search_name_gui(self):
        name = simpledialog.askstring("名字查詢", "請輸入要查詢的兩個漢字名字:")
        if not name:
            return
        name = name.strip()
        if len(name) != 2:
            messagebox.showwarning("查詢失敗", "名字必須為兩個漢字。")
            return
        char_a = name[0]; char_b = name[1]
        is_in_pool = char_a in MASTER_WORDS and char_b in MASTER_WORDS

        drawn_status = "❌ 待抽取"
        rows = db_get_history()
        for ts, n, tones in rows:
            if n == name:
                drawn_status = "✅ 已抽取"
                break

        result_title = f"【名字】: {name} 查詢結果"
        result_message = f"【總體狀態】:\n{'✅ 存在於字詞庫組合中' if is_in_pool else '❌ 不存在於字詞庫組合中'}\n"
        if is_in_pool:
            result_message += f"\n【抽取狀態】:\n{drawn_status}"
        else:
            result_message += f" (字詞 '{char_a}' 或 '{char_b}' 不在 {WORD_COUNT} 個字庫中)"
        messagebox.showinfo(result_title, result_message)

    def add_words_to_master(self, new_words):
        global MASTER_WORDS, WORD_COUNT, POOL_SIZE
        if not new_words:
            return True
        try:
            with open(WORDS_FILE, 'a', encoding='utf-8') as f:
                for word in new_words:
                    f.write(word + '\n')
        except Exception as e:
            messagebox.showerror("檔案錯誤", f"無法寫入 {WORDS_FILE} 檔案: {e}")
            return False
        MASTER_WORDS.extend(new_words)
        WORD_COUNT = len(MASTER_WORDS)
        POOL_SIZE = WORD_COUNT * WORD_COUNT
        self.master.title(f"名字抽取器 | 總組合數: {POOL_SIZE:,}")
        self.reset_database(show_message=False)
        self._update_progress_display(name=f"新增 {len(new_words)} 字後已重置", remaining=self._get_remaining_count())
        return True

    def display_info_gui(self):
        info = f"[一、字詞庫資訊]\n"
        info += f"  - 來源檔案: {WORDS_FILE}\n"
        info += f"  - 總字詞數 (N): {WORD_COUNT:,} 個\n"
        info += f"  - 總名字組合 (N x N): {POOL_SIZE:,} 個\n"
        remaining_count = self._get_remaining_count()
        drawn_count = POOL_SIZE - remaining_count
        info += f"\n[二、抽取進度]\n"
        info += f"  - 已抽取: {drawn_count:,} 個\n"
        info += f"  - 剩餘數量: {remaining_count:,} 個\n"
        last_reset = "N/A (未重置或檔案遺失)"
        if os.path.exists(STATUS_FILE):
            try:
                with open(STATUS_FILE, 'r', encoding='utf-8') as sf:
                    status_data = json.load(sf)
                    last_reset = status_data.get("last_reset", last_reset)
            except:
                pass
        info += f"\n[三、檔案狀態]\n"
        info += f"  - DB: {'✅ 存在' if os.path.exists(DB_FILE) else '❌ 遺失'}\n"
        info += f"  - 索引狀態 ({STATE_FILE}): {'✅ 存在' if os.path.exists(STATE_FILE) else '❌ 遺失'}\n"
        info += f"  - 歷史紀錄 (DB history table): {'✅ 存在' if os.path.exists(DB_FILE) else '❌ 遺失'}\n"
        info += f"  - 收藏清單 (DB favorites table): {'✅ 存在' if os.path.exists(DB_FILE) else '❌ 遺失'}\n"
        info += f"  - 上次重置時間: {last_reset}"
        messagebox.showinfo("系統狀態與資訊 (INFO)", info)

    def display_frequency_stats_gui(self):
        stats = get_word_frequency_stats()
        if all(count == 0 for count in stats.values()):
            messagebox.showinfo("字詞抽取頻率", "尚未抽取任何名字，或歷史記錄中沒有與當前字詞庫匹配的字。")
            return
        sorted_stats = sorted(stats.items(), key=lambda item: item[1], reverse=True)
        total_draws = sum(stats.values()) // 2
        stat_output = "【字詞抽取頻率統計】\n\n"
        stat_output += f"總抽取名字數: {total_draws:,} 個 (雙字計算: {total_draws * 2:,} 次)\n\n"
        for word, count in sorted_stats:
            if count > 0:
                stat_output += f"  - 字 '{word}': 被抽取 {count} 次\n"
        stats_window = tk.Toplevel(self.master)
        stats_window.title("字詞抽取頻率統計")
        stats_window.geometry("400x500")
        text_widget = scrolledtext.ScrolledText(stats_window, wrap=tk.WORD, font=('Courier New', 10))
        text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        text_widget.insert(tk.END, stat_output)
        text_widget.config(state=tk.DISABLED)

    def open_config_editor(self):
        editor_window = tk.Toplevel(self.master)
        editor_window.title("字詞庫管理 (words_list.txt)")
        editor_window.geometry("500x600")
        text_widget = scrolledtext.ScrolledText(editor_window, wrap=tk.WORD, font=('Courier New', 12))
        text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=0)
        stats_label = tk.Label(editor_window, textvariable=self.config_stats_var, font=('Microsoft JhengHei', 10, 'italic'), fg='gray')
        stats_label.pack(pady=(5, 10))
        try:
            with open(WORDS_FILE, 'r', encoding='utf-8') as f:
                content = f.read()
            text_widget.insert(tk.END, content)
            self.update_config_stats_gui(content)
        except Exception:
            text_widget.insert(tk.END, "# 找不到現有字詞庫檔案，請輸入您的漢字列表：\n愛\n麗\n雅\n靜\n風\n雲\n月\n星")
        text_widget.bind('<KeyRelease>', lambda event: self.update_config_stats_gui(text_widget.get("1.0", tk.END)))
        save_button = tk.Button(editor_window,
                                text="儲存字詞庫並重置索引",
                                command=lambda: self.save_and_reset_words(editor_window, text_widget.get("1.0", tk.END)),
                                font=('Microsoft JhengHei', 12),
                                bg="#FFB300")
        save_button.pack(pady=10)
        editor_window.protocol("WM_DELETE_WINDOW", lambda: self.confirm_close_editor(editor_window))

    def update_config_stats_gui(self, content):
        word_count, pool_size = analyze_words_from_text(content)
        self.config_stats_var.set(f"當前輸入分析： 總字數(N): {word_count:,} 個 │ 總組合數(N x N): {pool_size:,} 個")

    def confirm_close_editor(self, window):
        if messagebox.askyesno("確認關閉", "您確定要關閉編輯器而不儲存變更嗎？"):
            window.destroy()

    def save_and_reset_words(self, window, new_content):
        try:
            atomic_write(WORDS_FILE, new_content)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法寫入字詞庫檔案: {e}")
            return
        try:
            load_master_words()
        except SystemExit:
            return
        except Exception:
            messagebox.showerror("錯誤", "重新載入字詞庫時發生未知錯誤。")
            return
        initialize_database(reset_history=True)
        self.current_name = ""
        self._update_progress_display(remaining=POOL_SIZE)
        self.draw_button.config(state=tk.NORMAL)
        messagebox.showinfo("成功", f"字詞庫已更新。\n已使用 {WORD_COUNT} 個字重新生成 {POOL_SIZE:,} 個索引。")
        window.destroy()

# ==================== 程式啟動點 ====================
def load_master_words():
    """從外部檔案加載 MASTER_WORDS，並設置全域變數"""
    global MASTER_WORDS, POOL_SIZE, WORD_COUNT, WORD_TO_INDEX

    if not os.path.exists(WORDS_FILE):
        try:
            with open(WORDS_FILE, 'w', encoding='utf-8') as f:
                f.write("愛\n麗\n雅\n靜\n")
                f.write("風\n雲\n月\n星\n")
            messagebox.showerror("錯誤：找不到字詞庫",
                                 f"找不到字詞庫檔案 '{WORDS_FILE}'。\n\n"
                                 f"程式已在資料夾中為您創建了範本檔案，請編輯該檔案後，再次運行程式。")
            sys.exit(1)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法創建字詞庫檔案: {e}")
            sys.exit(1)

    try:
        with open(WORDS_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        final_words = []
        for ch in content:
            if ch.strip() and not ch.isspace() and ch not in [',', '，', '#', '\n']:
                final_words.append(ch)
        if not final_words:
            messagebox.showerror("錯誤", f"字詞庫檔案 '{WORDS_FILE}' 內容為空。請編輯後重新運行。")
            sys.exit(1)
        MASTER_WORDS = final_words
        WORD_COUNT = len(MASTER_WORDS)
        POOL_SIZE = WORD_COUNT * WORD_COUNT
        WORD_TO_INDEX = {word: i for i, word in enumerate(MASTER_WORDS)}
    except Exception as e:
        messagebox.showerror("錯誤", f"加載字詞庫時發生錯誤: {e}")
        sys.exit(1)

if __name__ == "__main__":
    setup_data_paths()
    load_master_words()

    # 初始化 DB 並載入索引緩存
    init_db()
    if not os.path.exists(DB_FILE) or (not db_get_remaining() and POOL_SIZE > 0):
        # 如果 DB 尚無 remaining 資料（首次啟動），則生成索引
        initialize_database(reset_history=True)
    else:
        load_indices_cache()

    root = tk.Tk()
    app = NameGeneratorApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()