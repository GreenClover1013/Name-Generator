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

# --- [新增] Pinyin 庫與語音學分析 ---
try:
    import pypinyin
    PINYIN_ENABLED = True
    
    # 定義被認為不流暢的聲調組合（聲調編號：1-5，5為輕聲）
    # 範例：(3, 3) 難以連讀，(4, 4) 語氣過重，(1, 1) 缺乏變化。
    UNSMOOTH_TONES = {(3, 3), (4, 4), (1, 1), (2, 2)} 
    
    def get_pinyin_with_tone(name):
        """轉換名字為拼音（帶聲調符號）和數字聲調元組。"""
        # 使用 TONE 樣式獲取帶符號的拼音用於顯示
        pinyin_display_result = pypinyin.pinyin(name, style=pypinyin.Style.TONE)
        display_pinyin = " ".join([p[0] for p in pinyin_display_result])
        
        # 使用 TONE3 樣式獲取帶數字的拼音用於分析
        pinyin_num_result = pypinyin.pinyin(name, style=pypinyin.Style.TONE3)
        tones = []
        for p in pinyin_num_result:
            p_str = p[0]
            # 提取聲調數字 (如果最後一個字元是數字，否則默認為 5 (輕聲))
            tone_num = int(p_str[-1]) if p_str and p_str[-1].isdigit() else 5
            tones.append(tone_num)
            
        return display_pinyin, tuple(tones)

    def is_tone_combination_smooth(tones):
        """檢查 2 字聲調組合是否在黑名單中。"""
        if not isinstance(tones, tuple) or len(tones) != 2:
            return True 
        
        if tones in UNSMOOTH_TONES:
            return False
            
        # 您可以根據需要在此處添加更多過濾規則
            
        return True

except ImportError:
    PINYIN_ENABLED = False
    print("Warning: pypinyin library not found. Pinyin analysis is disabled.")
# --- [結束新增] Pinyin 庫與語音學分析 ---

# ----------------- 配置 -----------------
WORDS_FILE = 'words_list.txt'
STATE_FILE = 'name_indices.json'
HISTORY_FILE = 'drawn_history.txt'
FAVORITES_FILE = 'favorites.txt'
STATUS_FILE = 'system_status.json'
DATA_DIR = 'name_generator_data'

MASTER_WORDS = []
POOL_SIZE = 0
WORD_COUNT = 0
# **核心優化：將索引緩存在記憶體中**
NAME_INDICES_CACHE = []
WORD_TO_INDEX = {}
HISTORY_RE = re.compile(r"\] - (.+?)(?: \[|$)")
# ------------------------------------------

# ==================== 核心數據函數 ====================
# 假設這個函數已經存在，用於將名字轉換回索引
def name_to_index(name):
    """將名字 (例如 '雅靜') 轉換回其在 MASTER_WORDS 中的索引。"""
    try:
        char_a = name[0]
        char_b = name[1]
        idx_a = MASTER_WORDS.index(char_a)
        idx_b = MASTER_WORDS.index(char_b)
        return idx_a * WORD_COUNT + idx_b
    except (ValueError, IndexError):
        return None

# ====================================================================
class RestoreExcludedDialog(tk.Toplevel):
    # 這裡的 master_app 就是 NameGeneratorApp 的實例 (self)
    def __init__(self, master_app, history_lines): 
        
        # 關鍵修正點：
        # super() 必須使用 NameGeneratorApp 實例的 'master' (即 Tkinter 的根視窗)
        # 來初始化 Toplevel，而不是 NameGeneratorApp 實例本身。
        super().__init__(master_app.master) 
        
        self.title("恢復已排除組合")
        self.geometry("600x500")
        
        # 儲存 NameGeneratorApp 實例，用於調用其方法（如 save_state_to_file）
        self.master_app = master_app 
        
        self.history_lines = history_lines  
        self.selected_indices = []

        # 顯示歷史記錄 (以下程式碼保持不變)
        tk.Label(self, text="請選擇要恢復的組合 (將重新加入待抽取列表)：").pack(pady=5)
        
        # 使用 Listbox 允許選擇
        self.listbox = tk.Listbox(self, selectmode=tk.MULTIPLE, height=20, width=80, font=('Courier New', 11))
        self.listbox.pack(padx=10, pady=5, fill="both", expand=True)

        # 載入歷史記錄到 Listbox (顯示名字部分)
        self.name_entries = [] 
        
        for line in history_lines:
            try:
                name_part = line.split(' - ')[1].split(' [')[0].strip()
                self.name_entries.append(name_part)
                self.listbox.insert(tk.END, line.strip())
            except Exception:
                pass

        # 按鈕區
        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="恢復選定組合", command=self.restore_selected, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="關閉", command=self.destroy).pack(side=tk.LEFT, padx=10)

    def restore_selected(self):
        selected_items_indices = self.listbox.curselection()
        if not selected_items_indices:
            messagebox.showwarning("提示", "請至少選擇一個組合進行恢復。")
            return

        indices_to_restore = []
        names_to_restore = []
        lines_to_keep = []
        
        # 1. 計算要恢復的索引
        for i in selected_items_indices:
            name = self.name_entries[i]
            index = name_to_index(name)
            if index is not None:
                indices_to_restore.append(index)
                names_to_restore.append(name)

        if not indices_to_restore:
            messagebox.showerror("錯誤", "無法解析選定的名字，無索引可恢復。")
            self.destroy()
            return
            
        # 2. 將索引添加回緩存 (NAME_INDICES_CACHE)
        global NAME_INDICES_CACHE
        NAME_INDICES_CACHE.extend(indices_to_restore)

        # 3. 從 HISTORY_FILE 中移除已恢復的行
        current_lines_set = set(self.history_lines)
        for i in selected_items_indices:
            # 從集合中移除選定的行
            current_lines_set.discard(self.history_lines[i]) 
            
        lines_to_keep = list(current_lines_set)

        # 4. 覆蓋寫入 HISTORY_FILE (只保留未選擇的行)
        try:
            with open(HISTORY_FILE, 'w', encoding='utf-8') as hf:
                hf.writelines([line + '\n' for line in lines_to_keep])
        except PermissionError:
            messagebox.showerror("錯誤", "寫入歷史檔案權限不足，請關閉佔用程式。")
            return

        # 5. 儲存更新後的 STATE_FILE
        self.master_app.save_state_to_file(NAME_INDICES_CACHE)
        
        messagebox.showinfo("成功", f"成功恢復 {len(indices_to_restore)} 個組合: {', '.join(names_to_restore[:3])}...")
        
        # 更新主介面顯示
        self.master_app._update_progress_display(remaining=len(NAME_INDICES_CACHE))
        self.destroy()

def setup_data_paths():
    """設定數據檔案的實際路徑並確保資料夾存在。"""
    global WORDS_FILE, STATE_FILE, HISTORY_FILE, FAVORITES_FILE, STATUS_FILE
    
    # 確保數據資料夾存在
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
        
    # 重新指定所有檔案的路徑
    WORDS_FILE = os.path.join(DATA_DIR, 'words_list.txt')
    STATE_FILE = os.path.join(DATA_DIR, 'name_indices.json')
    HISTORY_FILE = os.path.join(DATA_DIR, 'drawn_history.txt')
    FAVORITES_FILE = os.path.join(DATA_DIR, 'favorites.txt')
    STATUS_FILE = os.path.join(DATA_DIR, 'system_status.json')

def load_master_words():
    """從外部檔案加載 MASTER_WORDS，並設置全域變數"""
    global MASTER_WORDS, POOL_SIZE, WORD_COUNT
    
    if not os.path.exists(WORDS_FILE):
        try:
            with open(WORDS_FILE, 'w', encoding='utf-8') as f:
                f.write("愛 麗 雅 靜\n")
                f.write("風 雲 月 星\n")
                
            messagebox.showerror("錯誤：找不到字詞庫", 
                                f"找不到字詞庫檔案 '{WORDS_FILE}'。\n\n"
                                f"程式已在當前目錄為您創建了範本檔案，請編輯該檔案後，再次運行程式。")
            sys.exit(1)
            
        except Exception as e:
             messagebox.showerror("錯誤", f"無法創建字詞庫檔案: {e}")
             sys.exit(1)

    try:
        with open(WORDS_FILE, 'r', encoding='utf-8') as f:
            content = f.read()

        final_words = []
        for char in content:
            if char.strip() and not char.isspace() and char not in [',', '，', '#']:
                final_words.append(char)

        if not final_words:
            messagebox.showerror("錯誤", f"字詞庫檔案 '{WORDS_FILE}' 內容為空。請編輯後重新運行。")
            sys.exit(1)

        MASTER_WORDS = final_words
        WORD_COUNT = len(MASTER_WORDS)
        POOL_SIZE = WORD_COUNT * WORD_COUNT

        # **【新】性能優化：建立字元到索引的快速查找字典**
        global WORD_TO_INDEX 
        WORD_TO_INDEX = {word: i for i, word in enumerate(MASTER_WORDS)}
        
    except Exception as e:
        messagebox.showerror("錯誤", f"加載字詞庫時發生錯誤: {e}")
        sys.exit(1)


def analyze_words_from_text(content):
    """(功能 5 輔助) 從文字內容中分析總字數和總組合數，不影響全域狀態。"""
    final_words = []
    for char in content:
        if char.strip() and not char.isspace() and char not in [',', '，', '#']:
            final_words.append(char)
            
    word_count = len(final_words)
    pool_size = word_count * word_count
    return word_count, pool_size


def get_drawn_indices():
    """
    【優化版本】讀取歷史檔案，將已抽取的組合成員轉換回其索引。
    使用 WORD_TO_INDEX 字典將查找速度從 O(N) 降到 O(1)。
    """
    drawn_indices = set()
    # 確保 WORD_COUNT > 0 且字典存在
    if not os.path.exists(HISTORY_FILE) or WORD_COUNT == 0 or 'WORD_TO_INDEX' not in globals():
        return drawn_indices

    try:
        with open(HISTORY_FILE, 'r', encoding='utf-8') as hf:
            for line in hf:
                parts = line.split(' - ')
                if len(parts) < 2:
                    continue
                
                # 假設 name 格式為 兩個字
                name = parts[1].strip()
                if len(name) == 2:
                    char_a = name[0]
                    char_b = name[1]
                    
                    # **【關鍵優化】使用字典進行 O(1) 查找**
                    idx_a = WORD_TO_INDEX.get(char_a)
                    idx_b = WORD_TO_INDEX.get(char_b)
                    
                    # 只有當兩個字都存在於當前字詞庫時才計算索引
                    if idx_a is not None and idx_b is not None:
                        index = idx_a * WORD_COUNT + idx_b
                        drawn_indices.add(index)
                    # 否則，忽略 (該名字使用了已移除的字)
                        
    except Exception:
        pass # 容錯處理

    return drawn_indices


def get_word_frequency_stats():
    """(功能 4 輔助) 計算每個字詞在歷史記錄中的抽取頻率。"""
    frequency = {word: 0 for word in MASTER_WORDS}
    if not os.path.exists(HISTORY_FILE) or WORD_COUNT == 0:
        return frequency

    try:
        with open(HISTORY_FILE, 'r', encoding='utf-8') as hf:
            for line in hf:
                parts = line.split(' - ')
                if len(parts) < 2:
                    continue
                
                name = parts[1].strip()
                if len(name) == 2:
                    char_a = name[0]
                    char_b = name[1]
                    
                    if char_a in frequency:
                        frequency[char_a] += 1
                    if char_b in frequency:
                        frequency[char_b] += 1
                        
    except Exception:
        pass

    return frequency


def initialize_database(reset_history=True, exclude_drawn=False):
    """初始化數據庫，可選是否排除已抽取的組合，並將索引寫入緩存和檔案。"""
    global NAME_INDICES_CACHE # 【新】宣告使用全域緩存
    
    if POOL_SIZE == 0:
        return
        
    all_indices = list(range(POOL_SIZE))
    
    if exclude_drawn:
        # 注意：您需要確保 get_drawn_indices() 函數已定義
        drawn_indices = get_drawn_indices() 
        
        initial_indices_set = set(all_indices)
        remaining_indices_set = initial_indices_set - drawn_indices
        indices = list(remaining_indices_set)
        
        excluded_count = len(drawn_indices)
        if excluded_count > 0:
            messagebox.showinfo("智慧重置資訊", f"已從 {POOL_SIZE:,} 個組合中，排除 {excluded_count:,} 個已抽取的組合。\n\n"
                                             f"剩餘 {len(indices):,} 個組合將用於新的抽取列表。")
            
        if not indices:
            messagebox.showwarning("警告", "所有組合皆已被排除！新的索引列表為空。")
        
    else:
        indices = all_indices
    
    random.shuffle(indices)
    
    # ---------------- 緩存優化相關修改 ----------------
    # 1. 更新記憶體緩存
    NAME_INDICES_CACHE = indices 
    
    # 2. 將完整的列表寫入 STATE_FILE 
    with open(STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(indices, f, ensure_ascii=False)

    # ----------------------------------------------------
        
    if reset_history:
        # 清空歷史和收藏
        if os.path.exists(HISTORY_FILE):
            os.remove(HISTORY_FILE)
        if os.path.exists(FAVORITES_FILE):
            os.remove(FAVORITES_FILE)
            
        # 記錄重置時間
        reset_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status_data = {"last_reset": reset_time}
        with open(STATUS_FILE, 'w', encoding='utf-8') as sf:
            json.dump(status_data, sf)

def load_indices_cache():
    """【新函數】程式啟動時，從檔案載入索引到記憶體緩存。"""
    global NAME_INDICES_CACHE
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r', encoding='utf-8') as f:
                NAME_INDICES_CACHE = json.load(f)
        except json.JSONDecodeError:
            # 檔案損壞時，執行重置
            messagebox.showerror("錯誤", "狀態檔案損壞，將重置數據庫。")
            initialize_database()       

def save_indices_cache():
    """【新函數】程式退出前，將記憶體緩存寫回檔案。"""
    global NAME_INDICES_CACHE
    try:
        # **【注意】這可能需要幾秒，但只在結束時執行**
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(NAME_INDICES_CACHE, f, ensure_ascii=False)
    except Exception as e:
        print(f"警告：無法保存狀態檔案: {e}")

def get_unique_name():
    """
    【優化版本】從記憶體緩存中抽取索引，並在抽取時檢查拼音流暢度。
    """
    global NAME_INDICES_CACHE
    
    if not NAME_INDICES_CACHE:
        return None, 0 # 已抽取完畢
    
    name = None
    is_valid_name = False
    
    # 循環直到找到一個有效且流暢的名字，或緩存耗盡
    while NAME_INDICES_CACHE:
        # 1. 記憶體操作：抽取索引 (極速)
        next_index = NAME_INDICES_CACHE.pop()
    
        idx_a = next_index // WORD_COUNT
        idx_b = next_index % WORD_COUNT
    
        if idx_a >= WORD_COUNT or idx_b >= WORD_COUNT:
             messagebox.showerror("數據錯誤", "索引超出範圍，請重置數據庫。")
             return None, len(NAME_INDICES_CACHE) 
    
        name = MASTER_WORDS[idx_a] + MASTER_WORDS[idx_b]
        
        # 2. **【關鍵新增】拼音和聲調檢查邏輯**
        if PINYIN_ENABLED:
            try:
                # 獲取聲調元組
                _, tones = get_pinyin_with_tone(name) 
                
                if is_tone_combination_smooth(tones):
                    is_valid_name = True
                    break # 找到流暢的名字，跳出迴圈
            except Exception:
                # 拼音轉換失敗（罕見字等），視為有效並跳出
                is_valid_name = True
                break
        else:
            # 如果沒有安裝 pypinyin，則視為有效
            is_valid_name = True
            break
            
    # 3. 檢查是否因為緩存耗盡而退出
    if not is_valid_name:
        return None, 0

    # 4. 檔案操作：只記錄歷史 (檔案小，速度快)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    history_entry = f"[{timestamp}] - {name}\n"
    with open(HISTORY_FILE, 'a', encoding='utf-8') as hf:
        hf.write(history_entry)

    return name, len(NAME_INDICES_CACHE)


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

class BatchWordManagerDialog(tk.Toplevel):
    """字詞庫內容管理對話框 (模態)"""
    def __init__(self, master, app_instance): 
        super().__init__(master)
        self.title("字詞庫內容管理")
        self.geometry("500x600")
        self.app_instance = app_instance 

        # tk.Label(self, text="編輯當前字詞 (每行一個字)：", # 如果要移除提示，請刪除此行
        #          font=('Microsoft JhengHei', 11, 'bold')).pack(pady=5)
        
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
        """保存字詞庫變更並自動重啟應用程式。"""
        raw = self.word_edit_area.get('1.0', tk.END)
        clean_words = sorted({w.strip() for w in raw.split('\n') if w.strip()})

        if len(clean_words) < 2:
            messagebox.showerror("保存失敗", "字詞庫至少需要兩個字。")
            return

        # 寫入檔案
        try:
            with open(WORDS_FILE, 'w', encoding='utf-8') as f:
                f.write("\n".join(clean_words))
        except Exception as e:
            messagebox.showerror("保存失敗", f"寫入檔案時發生錯誤:\n{e}")
            return

        messagebox.showinfo(
            "保存成功",
            f"字詞庫已更新，共 {len(clean_words)} 個字，程式將重新啟動。"
        )

        # 關閉對話框 → 停止主迴圈 → 重啟
        self.master.quit()

        python = sys.executable
        try:
            os.execl(python, python, *sys.argv)
        except Exception as e:
            messagebox.showerror(
                "重啟失敗",
                f"無法自動重啟，請手動重新啟動：\n{e}"
            ) 

# -------------------------------------------------------------------
# 將 `BatchWordManagerDialog` 放在 NameGeneratorApp 類別之外
# -------------------------------------------------------------------

def get_excluded_names():
    """從歷史紀錄讀取所有已抽取的名字。"""
    if not os.path.exists(HISTORY_FILE):
        return "尚無抽取歷史記錄。", 0

    try:
        with open(HISTORY_FILE, 'r', encoding='utf-8') as hf:
            lines = [line.strip() for line in hf if line.strip()]
        return "\n".join(reversed(lines)), len(lines)
    except Exception as e:
        return f"讀取歷史檔案時發生錯誤:\n{e}", 0

# ==================== Tkinter 介面邏輯 ====================

class NameGeneratorApp:
    def __init__(self, master):
        
        # 這是關鍵的一行！請檢查您的程式碼中是否遺失了它。
        self.master = master 
        
        # 確保 master 被正確賦值後，才能在後續方法中使用 self.master
        master.title(f"名字抽取器 | 總組合數: {POOL_SIZE:,}")
        master.geometry("750x450") # 請使用您當前的尺寸
        
        master.config(bg='#F0F0F0')
        self.name_var = tk.StringVar(master, value="準備就緒，請點擊抽取")
        self.progress_var = tk.StringVar(master)
        self.pinyin_var = tk.StringVar(master, value="") # <<< 新增 Pinyin 顯示變數
        self.config_stats_var = tk.StringVar(master, value="") 

        self._setup_ui()
        self._update_progress_display()
    
    def on_closing(self):
        """視窗關閉時執行保存操作，將緩存寫回硬碟。"""
        save_indices_cache()
        self.master.destroy()

    def _setup_ui(self):
        main_bg = '#F0F0F0' 
        name_fg = 'black' 
        
        # 1. 名字顯示區 (Entry) - 支援複製
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

        # --- [新增] Pinyin 顯示區 ---
        self.pinyin_display = tk.Label(self.master,
                                      textvariable=self.pinyin_var,
                                      font=('Microsoft JhengHei', 12, 'italic'),
                                      fg='gray',
                                      bg=main_bg,
                                      pady=0)
        self.pinyin_display.pack(pady=(0, 5))
        # --- [結束新增] Pinyin 顯示區 ---

        # 2. 進度條顯示區
        self.progress_label = tk.Label(self.master, 
                                        textvariable=self.progress_var,
                                        font=('Microsoft JhengHei', 10),
                                        bg=main_bg, 
                                        pady=10)
        self.progress_label.pack()
        
        # 3. 抽取按鈕
        self.draw_button = tk.Button(self.master, 
                                     text="抽取下一個名字 (Click)", 
                                     command=self.draw_name, 
                                     font=('Microsoft JhengHei', 14), 
                                     bg="#4CAF50", 
                                     fg='white', 
                                     width=25,
                                     height=2)
        self.draw_button.pack(pady=10)

        # --- [新增功能] 批量抽取區塊 ---
        batch_frame = tk.Frame(self.master, bg=main_bg)
        batch_frame.pack(pady=5)

        tk.Label(batch_frame, text="批量數量:", bg=main_bg, font=('Microsoft JhengHei', 10)).pack(side=tk.LEFT, padx=(10, 2))
        
        self.batch_count_var = tk.StringVar(self.master, value='50') # 預設值 50
        self.batch_count_entry = tk.Entry(batch_frame, 
                                          textvariable=self.batch_count_var, 
                                          width=5, 
                                          font=('Microsoft JhengHei', 10))
        self.batch_count_entry.pack(side=tk.LEFT, padx=(0, 10))

        self.batch_draw_button = tk.Button(batch_frame, 
                                     text="批量抽取並預覽", 
                                     command=self.batch_draw_gui, # 呼叫新方法
                                     font=('Microsoft JhengHei', 10), 
                                     bg="#03A9F4", 
                                     fg='white', 
                                     width=15)
        self.batch_draw_button.pack(side=tk.LEFT, padx=5)
        # --- [結束新增] 批量抽取區塊 ---
        
        # 4. 輔助功能按鈕群組 (上排 - 歷史與操作)
        button_frame_top = tk.Frame(self.master, bg=main_bg)
        button_frame_top.pack(pady=5)
        
        # 1. 檢視歷史 (左)
        self.history_button = tk.Button(button_frame_top, text="檢視歷史 (l)", command=self.view_history_gui, font=('Microsoft JhengHei', 10), width=12)
        self.history_button.pack(side=tk.LEFT, padx=5)

        # 2. 檢視收藏
        self.view_favorite_button = tk.Button(button_frame_top, text="檢視收藏 (v)", command=self.view_favorites_gui, font=('Microsoft JhengHei', 10), width=12)
        self.view_favorite_button.pack(side=tk.LEFT, padx=5)

        # 3. 匯出歷史
        self.export_button = tk.Button(button_frame_top, text="匯出歷史 (e)", command=self.export_history_gui, font=('Microsoft JhengHei', 10), width=12)
        self.export_button.pack(side=tk.LEFT, padx=5)

        # 4. 系統資訊 (右)
        self.info_button = tk.Button(button_frame_top, text="系統資訊 (i)", command=self.display_info_gui, font=('Microsoft JhengHei', 10), width=12)
        self.info_button.pack(side=tk.LEFT, padx=5)
        
        # 5. 輔助功能按鈕群組 (中排 - 操作與管理)
        button_frame_middle = tk.Frame(self.master, bg=main_bg)
        button_frame_middle.pack(pady=5)

        # 1. **重置數據庫 (放在最左邊)**
        self.reset_button = tk.Button(button_frame_middle, text="重置數據庫 (r)", command=self.reset_database, font=('Microsoft JhengHei', 10), width=12, bg="#D32F2F", fg="white")
        self.reset_button.pack(side=tk.LEFT, padx=5)

        # 檢視排除列表
        self.excluded_button = tk.Button(button_frame_middle,
                                text="檢視排除列表",
                                command=self.view_excluded_names_gui, # 調用新函數
                                font=('Microsoft JhengHei', 10),
                                width=12)
        self.excluded_button.pack(side=tk.LEFT, padx=5)

        self.restore_button = tk.Button(button_frame_top, 
                                text="恢復排除組合", 
                                command=self.view_and_restore_excluded_gui, 
                                font=('Microsoft JhengHei', 10), 
                                width=12,
                                bg="#CDDC39") # 淺綠色
        self.restore_button.pack(side=tk.LEFT, padx=5)

        # 2. **字詞庫管理 (原來的 manage_button)**
        self.manage_button = tk.Button(button_frame_middle,
                                        text="字詞庫管理 (m)",
                                        command=self.manage_words_gui,
                                        font=('Arial', 10),
                                        width=12,
                                        bg="#FFE0B2")
        self.manage_button.pack(side=tk.LEFT, padx=5)

        # 3. 查詢名字
        self.search_button = tk.Button(button_frame_middle, text="查詢名字 (s)", command=self.search_name_gui, font=('Microsoft JhengHei', 10), width=12)
        self.search_button.pack(side=tk.LEFT, padx=5)

        # 4. 字詞頻率 (原來的 frequency_button，因為它也與字詞管理相關)
        self.frequency_button = tk.Button(button_frame_middle,
                                            text="字詞頻率 (w)",
                                            command=self.display_frequency_stats_gui,
                                            font=('Microsoft JhengHei', 10),
                                            width=12, # 寬度統一
                                            bg="#B3E5FC") 
        self.frequency_button.pack(side=tk.LEFT, padx=5)

        # 6. 抽取操作輔助按鈕群組 (下排)
        button_frame_last = tk.Frame(self.master, bg=main_bg)
        button_frame_last.pack(pady=5)

        # 1. 排除
        self.exclude_button = tk.Button(button_frame_last,
                                        text="排除此組合 (x)",
                                        command=self.exclude_current_name_gui,
                                        font=('Microsoft JhengHei', 10),
                                        width=12,
                                        bg="#FFCDD2") 
        self.exclude_button.pack(side=tk.LEFT, padx=10) # 增大 padx

        # 2. 收藏
        self.favorite_button = tk.Button(button_frame_last, text="收藏名字 (f)", command=self.add_favorite_gui, font=('Microsoft JhengHei', 10), width=12)
        self.favorite_button.pack(side=tk.LEFT, padx=10) # 增大 padx

        # 3. 撤銷抽取
        self.undo_button = tk.Button(button_frame_last,
                                        text="撤銷抽取 (u)",
                                        command=self.undo_last_draw_gui,
                                        font=('Microsoft JhengHei', 10),
                                        width=12,
                                        bg="#D1C4E9")
        self.undo_button.pack(side=tk.LEFT, padx=10) # 增大 padx

        # 6. 字詞統計與管理按鈕群組 (下排)
        button_frame_last = tk.Frame(self.master, bg=main_bg)
        button_frame_last.pack(pady=5)

        # --- 新增：檢視字詞庫按鈕 ---
        self.view_words_button = tk.Button(button_frame_last,
                                        text="檢視字詞庫 (w)",
                                        command=self.view_word_list_gui, # 呼叫新方法
                                        font=('Microsoft JhengHei', 10),
                                        width=15,
                                        bg="#C8E6C9") # 淺綠色
        self.view_words_button.pack(side=tk.LEFT, padx=5)
        # ------------------------------
    
    def view_and_restore_excluded_gui(self):
        """啟動對話框來選擇並恢復已排除的組合。"""
        if not os.path.exists(HISTORY_FILE):
            messagebox.showwarning("提示", "目前沒有抽取歷史記錄可供恢復。")
            return
            
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as hf:
                # 讀取所有行，準備傳遞給對話框
                history_lines = [line.strip() for line in hf.readlines() if line.strip()]

            if not history_lines:
                messagebox.showwarning("提示", "歷史記錄檔案為空。")
                return

            RestoreExcludedDialog(self, history_lines)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取歷史檔案時發生錯誤: {e}")
    
    def view_word_list_gui(self):
        """顯示當前字詞庫內容。"""
        word_window = tk.Toplevel(self.master)
        word_window.title(f"當前字詞庫（總字數: {len(MASTER_WORDS)}）")
        word_window.geometry("450x600")

        sorted_words = sorted(MASTER_WORDS)
        WORDS_PER_LINE = 10
        lines = [
            " | ".join(sorted_words[i:i + WORDS_PER_LINE])
            for i in range(0, len(sorted_words), WORDS_PER_LINE)
        ]

        text_area = scrolledtext.ScrolledText(
            word_window, wrap=tk.WORD, font=('Microsoft JhengHei', 12)
        )
        text_area.insert(tk.END, "\n".join(lines))
        text_area.config(state=tk.DISABLED)
        text_area.pack(expand=True, fill='both')

        tk.Button(word_window, text="關閉", command=word_window.destroy).pack(pady=10)
        
    def view_excluded_names_gui(self):
        """顯示所有已抽取的/被排除的名字列表。"""
        # 由於已抽取的組合就是歷史記錄，我們直接調用獲取歷史記錄的函數
        history_text, count = get_excluded_names()
            
        title = f"已被排除的組合列表 ({count:,} 個)"
            
        # 創建一個新的 Toplevel 視窗
        history_window = tk.Toplevel(self.master)
        history_window.title(title)
        history_window.geometry("550x400")
            
        # 創建 ScrolledText 區塊來顯示內容
        text_area = scrolledtext.ScrolledText(history_window, 
                                            wrap=tk.WORD, 
                                            font=('Courier New', 11), 
                                            padx=10, pady=10)
        text_area.insert(tk.END, history_text)
        text_area.config(state=tk.DISABLED) # 設為只讀
        text_area.pack(expand=True, fill='both')
            
        # 新增一個關閉按鈕
        tk.Button(history_window, text="關閉", command=history_window.destroy, width=10).pack(pady=5)

    def _get_remaining_count(self):
        """【改】從記憶體緩存中獲取當前剩餘數量"""
        global NAME_INDICES_CACHE
        return len(NAME_INDICES_CACHE)

    def _update_progress_display(self, name=None, remaining=None, pinyin_str=None): # <<< 增加 pinyin_str 參數
        if remaining is None:
            remaining = self._get_remaining_count()
            
        progress_line = get_progress_bar(remaining)
        self.progress_var.set(progress_line)
        
        if name:
            self.name_var.set(name)
        elif remaining == 0:
            self.name_var.set("已全部抽取完畢！")
            self.draw_button.config(state=tk.DISABLED)
            
        # --- Pinyin 顯示更新 ---
        if pinyin_str is not None:
            self.pinyin_var.set(pinyin_str)
        else:
            self.pinyin_var.set("")
        # --- 結束 Pinyin 顯示更新 ---


    def draw_name(self):
        """點擊抽取按鈕時執行的函數，包含聲調過濾邏輯。"""
        
        # 設定一個合理的嘗試次數，防止無限循環 (例如：總組合數的 10 倍，或者最大 1000 次)
        MAX_ATTEMPTS = min(POOL_SIZE, 1000) 

        for attempt in range(MAX_ATTEMPTS):
            name, remaining = get_unique_name()

            if not name:
                self.current_name = ""
                self._update_progress_display(name, remaining)
                messagebox.showinfo("提示", "所有名字已抽取完畢！")
                return

            pinyin_str = None
            tones = None
            
            if PINYIN_ENABLED:
                try:
                    pinyin_str, tones = get_pinyin_with_tone(name)
                    is_smooth = is_tone_combination_smooth(tones)
                except Exception:
                    # 如果 pypinyin 處理漢字失敗，則視為通過
                    is_smooth = True
            else:
                # 如果 pypinyin 未安裝，則不執行過濾
                is_smooth = True

            if is_smooth:
                # --- 找到符合條件的名字：執行複製和更新 ---
                self.current_name = name
                
                try:
                    self.master.clipboard_clear()
                    self.master.clipboard_append(name)
                except Exception:
                    pass 

                self._update_progress_display(name, remaining, pinyin_str)
                return
            else:
                # --- 不符合條件：跳過此名字並繼續抽取下一個 ---
                # 注意：由於 get_unique_name() 已將此名字記錄到歷史並從索引中移除，
                # 因此它會被視為「已抽但未顯示」，這可能會導致歷史記錄中出現大量被跳過的名字。
                # 這是後過濾機制（Post-filtering）的必然結果。
                print(f"Skipped name: {name} (Tones: {tones}) due to unsmooth combination.")
                
        # 如果嘗試次數耗盡仍未找到流暢的名字
        self.current_name = ""
        self.pinyin_var.set("")
        remaining = self._get_remaining_count()
        self._update_progress_display(name="連續過濾失敗", remaining=remaining)
        messagebox.showwarning("抽取失敗", f"連續 {MAX_ATTEMPTS} 次抽取都遇到不流暢的聲調組合，請嘗試重置或調整過濾規則。")
        self.draw_button.config(state=tk.DISABLED)


    def undo_last_draw_gui(self):

        if not (os.path.exists(HISTORY_FILE) and os.path.exists(STATE_FILE)):
            messagebox.showwarning("無法撤銷", "缺少歷史或狀態檔案。")
            return

        # 讀取歷史紀錄
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                lines = f.readlines()
        except Exception as e:
            messagebox.showerror("撤銷失敗", f"讀取歷史檔案錯誤:\n{e}")
            return

        if not lines:
            messagebox.showwarning("無法撤銷", "歷史記錄為空。")
            return

        last_line = lines[-1]

        # 解析名稱
        m = HISTORY_RE.search(last_line)
        if not m:
            messagebox.showwarning("撤銷警告", f"無法從歷史解析名字：\n{last_line}")
            return

        name = m.group(1)
        if len(name) != 2:
            messagebox.showwarning("撤銷警告", f"名字長度異常：{name}")
            return

        # 回寫歷史檔案
        try:
            with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
                f.writelines(lines[:-1])
        except Exception as e:
            messagebox.showerror("撤銷失敗", f"寫入歷史檔案錯誤:\n{e}")
            return

        # 恢復索引
        try:
            idx_a = MASTER_WORDS.index(name[0])
            idx_b = MASTER_WORDS.index(name[1])
            index_to_restore = idx_a * len(MASTER_WORDS) + idx_b
        except ValueError:
            messagebox.showwarning("撤銷警告", f"字詞不在庫中：{name}")
            return

        # 修改 STATE_FILE
        try:
            with open(STATE_FILE, 'r', encoding='utf-8') as sf:
                indices = json.load(sf)
            if index_to_restore not in indices:
                indices.append(index_to_restore)
            with open(STATE_FILE, 'w', encoding='utf-8') as sf:
                json.dump(indices, sf)
        except Exception as e:
            messagebox.showerror("撤銷失敗", f"狀態檔案處理錯誤:\n{e}")
            return

        messagebox.showinfo("成功", f"已撤銷抽取：{name}")

    def batch_draw_gui(self):
        """批量抽取名字並在獨立視窗中顯示預覽。"""
        try:
            count = int(self.batch_count_var.get())
            if count <= 0 or count > 1000:
                messagebox.showwarning("警告", "批量抽取數量必須是 1 到 1000 之間的整數。")
                return
        except ValueError:
            messagebox.showwarning("警告", "請輸入有效的批量抽取數量。")
            return

        
        drawn_names = []
        original_remaining = self._get_remaining_count()
        draw_limit = min(count, original_remaining)
        
        if draw_limit == 0:
            messagebox.showinfo("提示", "剩餘待抽取名字數量為 0。")
            return

        # 重複呼叫核心抽取邏輯
        for _ in range(draw_limit):
            name, remaining = get_unique_name()
            if name:
                drawn_names.append(name)
            else:
                break
        
        # 批量抽取後，更新主界面的狀態
        final_remaining = self._get_remaining_count()
        self.current_name = drawn_names[-1] if drawn_names else ""
        self._update_progress_display(name=self.current_name, remaining=final_remaining)
        
        # 顯示批量抽取結果視窗
        self._display_batch_results(drawn_names, draw_limit)
    
    def manage_words_gui(self):
        """呼叫批量字詞庫管理對話框。"""
        if hasattr(self, '_batch_dialog') and self._batch_dialog.winfo_exists():
            self._batch_dialog.lift()
            return
            
        # 模態對話框
        self._batch_dialog = BatchWordManagerDialog(self.master, self)
        self.master.wait_window(self._batch_dialog) 

    def _display_batch_results(self, names, draw_count):
        """顯示批量抽取結果的獨立視窗。"""
        results_window = tk.Toplevel(self.master)
        results_window.title(f"批量抽取結果 ({len(names)} 個)")
        results_window.geometry("400x550")
        
        header_text = f"成功抽取 {len(names)} 個名字，總共消耗了 {len(names)} 個索引。\n"
        header_label = tk.Label(results_window, text=header_text, font=('Microsoft JhengHei', 10, 'bold'), pady=5)
        header_label.pack()

        # 1. 結果列表
        text_widget = scrolledtext.ScrolledText(results_window, wrap=tk.WORD, font=('Courier New', 12))
        text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=(0, 10))
        
        output_content = ""
        for i, name in enumerate(names):
            output_content += f"{i+1:03d}. {name}\n"
            
        text_widget.insert(tk.END, output_content)
        text_widget.config(state=tk.DISABLED)

        # 2. 複製按鈕
        def copy_to_clipboard():
            """將所有結果複製到剪貼簿"""
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

    def reset_database(self):
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
        
        messagebox.showinfo(reset_message, f"數據庫已重置。\n\n"
                                          f"總字數: {WORD_COUNT} 個\n"
                                          f"總組合數: {POOL_SIZE:,} 個\n"
                                          f"剩餘待抽取數量: {final_remaining_count:,} 個")
        self.name_var.set("重置完成，請點擊抽取")
    
    def exclude_current_name_gui(self):
        """永久排除當前顯示的名字組合，並從待抽取列表中移除索引。"""
        name_to_exclude = self.current_name
        
        if not name_to_exclude or name_to_exclude == "已全部抽取完畢！" or len(name_to_exclude) != 2:
            messagebox.showwarning("無法排除", "請先抽取一個名字，且名字必須為兩個漢字。")
            return
            
        if not messagebox.askyesno("確認排除", f"您確定要將名字 '{name_to_exclude}' 從待抽取列表**永久排除**嗎？\n\n注意：這將減少總待抽取組合數。"):
            return

        try:
            # 1. 計算要排除的名字的索引 (與抽取消組合的邏輯相反)
            char_a = name_to_exclude[0]
            char_b = name_to_exclude[1]
            idx_a = MASTER_WORDS.index(char_a)
            idx_b = MASTER_WORDS.index(char_b)
            index_to_exclude = idx_a * WORD_COUNT + idx_b

            # 2. 從 STATE_FILE 索引列表中移除它
            with open(STATE_FILE, 'r', encoding='utf-8') as sf:
                indices = json.load(sf)
            
            try:
                indices.remove(index_to_exclude)
            except ValueError:
                # 如果索引不在列表中（可能已經被抽走），則提示但仍記錄
                messagebox.showwarning("提示", f"名字 '{name_to_exclude}' 不在當前待抽取列表中，但將記錄排除動作。")
            
            with open(STATE_FILE, 'w', encoding='utf-8') as sf:
                json.dump(indices, sf, ensure_ascii=False)

            # 3. 記錄到排除檔案 (用於追蹤和確認排除記錄)
            EXCLUDED_FILE = 'excluded_names_log.txt'
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_entry = f"[{timestamp}] - {name_to_exclude}\n"
            with open(EXCLUDED_FILE, 'a', encoding='utf-8') as ef:
                ef.write(log_entry)
            
            # 4. 更新 GUI 狀態
            self.current_name = ""
            self._update_progress_display(name=f"'{name_to_exclude}' 已永久排除", remaining=len(indices))
            
            messagebox.showinfo("排除成功", f"名字 '{name_to_exclude}' 已從待抽取組合中永久移除。")
            
        except ValueError:
            messagebox.showerror("錯誤", "當前字詞庫中不包含此名字的字詞，無法排除。")
        except Exception as e:
            messagebox.showerror("錯誤", f"執行排除操作時發生錯誤: {e}")

    def add_favorite_gui(self):
        if self.current_name:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            entry = f"[{timestamp}] - {self.current_name}\n"
            try:
                with open(FAVORITES_FILE, 'a', encoding='utf-8') as ff:
                    ff.write(entry)
                messagebox.showinfo("收藏成功", f"'{self.current_name}' 已加入收藏清單 ({FAVORITES_FILE})。")
            except Exception as e:
                messagebox.showerror("錯誤", f"無法寫入收藏檔案: {e}")
        else:
            messagebox.showwarning("提示", "請先抽取一個名字再進行收藏。")

    def view_history_gui(self):
        history_window = tk.Toplevel(self.master)
        history_window.title("抽取歷史紀錄 (drawn_history.txt)")
        history_window.geometry("400x500")
        
        text_widget = scrolledtext.ScrolledText(history_window, wrap=tk.WORD, font=('Courier New', 10))
        text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        
        if not os.path.exists(HISTORY_FILE):
            text_widget.insert(tk.END, "檔案不存在，尚未抽取任何名字。\n(請先抽取名字)")
        else:
            try:
                with open(HISTORY_FILE, 'r', encoding='utf-8') as hf:
                    history_content = hf.read()
                    text_widget.insert(tk.END, history_content)
                    text_widget.see(tk.END)
            except Exception as e:
                text_widget.insert(tk.END, f"無法讀取歷史檔案: {e}")
                
        text_widget.config(state=tk.DISABLED)

    def view_favorites_gui(self):
        favorite_window = tk.Toplevel(self.master)
        favorite_window.title("收藏名字清單 (favorites.txt)")
        favorite_window.geometry("400x500")
        
        text_widget = scrolledtext.ScrolledText(favorite_window, wrap=tk.WORD, font=('Courier New', 10))
        text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        
        if not os.path.exists(FAVORITES_FILE):
            text_widget.insert(tk.END, "檔案不存在，尚未收藏任何名字。\n(請先抽取並收藏名字)")
        else:
            try:
                with open(FAVORITES_FILE, 'r', encoding='utf-8') as ff:
                    favorite_content = ff.read()
                    text_widget.insert(tk.END, favorite_content)
                    text_widget.see(tk.END)
            except Exception as e:
                text_widget.insert(tk.END, f"無法讀取收藏檔案: {e}")
                
        text_widget.config(state=tk.DISABLED)

    def export_history_gui(self):
        if not os.path.exists(HISTORY_FILE):
            messagebox.showwarning("匯出失敗", "抽取歷史檔案不存在，無法匯出。")
            return
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_filename = f"Export_History_{timestamp}.txt"
        
        try:
            shutil.copyfile(HISTORY_FILE, export_filename)
            messagebox.showinfo("匯出成功", f"歷史記錄已成功匯出至:\n{export_filename}")
        except Exception as e:
            messagebox.showerror("匯出失敗", f"檔案複製錯誤: {e}")

    def search_name_gui(self):
        name = simpledialog.askstring("名字查詢", "請輸入要查詢的兩個漢字名字:")
        
        if not name:
            return
            
        name = name.strip()
        
        if len(name) != 2:
            messagebox.showwarning("查詢失敗", "名字必須為兩個漢字。")
            return

        char_a = name[0]
        char_b = name[1]
        is_in_pool = char_a in MASTER_WORDS and char_b in MASTER_WORDS

        drawn_status = "❌ 待抽取"
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, 'r', encoding='utf-8') as hf:
                    history_content = hf.read()
                
                if f"- {name}\n" in history_content or f"- {name}\r\n" in history_content:
                    drawn_status = "✅ 已抽取"
            except:
                drawn_status = "❓ 歷史檔案讀取錯誤"
        
        result_title = f"【名字】: {name} 查詢結果"
        result_message = f"【總體狀態】:\n{'✅ 存在於字詞庫組合中' if is_in_pool else '❌ 不存在於字詞庫組合中'}\n"
        
        if is_in_pool:
            result_message += f"\n【抽取狀態】:\n{drawn_status}"
        else:
             result_message += f" (字詞 '{char_a}' 或 '{char_b}' 不在 {WORD_COUNT} 個字庫中)"

        messagebox.showinfo(result_title, result_message)
    
    def add_words_to_master(self, new_words):
        """
        [NEW] 將多個新字添加到 words_list.txt 檔案中，並更新 MASTER_WORDS。
        返回 True 表示成功，False 表示失敗。
        """
        global MASTER_WORDS, WORD_COUNT, POOL_SIZE

        if not new_words:
            return True

        # 1. 寫入檔案
        try:
            with open(WORDS_FILE, 'a', encoding='utf-8') as f:
                for word in new_words:
                    f.write(word + '\n')
        except Exception as e:
            messagebox.showerror("檔案錯誤", f"無法寫入 {WORDS_FILE} 檔案: {e}")
            return False

        # 2. 更新全局變數
        MASTER_WORDS.extend(new_words)
        WORD_COUNT = len(MASTER_WORDS)
        POOL_SIZE = WORD_COUNT * WORD_COUNT
        
        # 3. 更新主視窗標題
        self.master.title(f"名字抽取器 | 總組合數: {POOL_SIZE:,}")
        
        # 4. 初始化狀態：當字詞庫改變時，重置整個抽取狀態
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
        info += f"  - 索引狀態 ({STATE_FILE}): {'✅ 存在' if os.path.exists(STATE_FILE) else '❌ 遺失'}\n"
        info += f"  - 歷史紀錄 ({HISTORY_FILE}): {'✅ 存在' if os.path.exists(HISTORY_FILE) else '❌ 遺失'}\n"
        info += f"  - 收藏清單 ({FAVORITES_FILE}): {'✅ 存在' if os.path.exists(FAVORITES_FILE) else '❌ 遺失'}\n"
        info += f"  - 上次重置時間: {last_reset}"

        messagebox.showinfo("系統狀態與資訊 (INFO)", info)

    def manage_words_gui(self):
        """[NEW] 呼叫批量字詞庫管理對話框。"""
        if hasattr(self, '_batch_dialog') and self._batch_dialog.winfo_exists():
            # 如果對話框已存在，則將其帶到前景
            self._batch_dialog.lift()
            return
            
        self._batch_dialog = BatchWordManagerDialog(self.master, self)
        self.master.wait_window(self._batch_dialog) # 等待對話框關閉

    def display_frequency_stats_gui(self):
        """功能 4: 顯示字詞抽取頻率統計。"""
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

        # 編輯區
        text_widget = scrolledtext.ScrolledText(editor_window, wrap=tk.WORD, font=('Courier New', 12))
        text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=0)
        
        # 功能 5: 統計標籤
        stats_label = tk.Label(editor_window, textvariable=self.config_stats_var, font=('Microsoft JhengHei', 10, 'italic'), fg='gray')
        stats_label.pack(pady=(5, 10))
        
        # 讀取現有內容
        try:
            with open(WORDS_FILE, 'r', encoding='utf-8') as f:
                content = f.read()
            text_widget.insert(tk.END, content)
            # 初始統計
            self.update_config_stats_gui(content) 
        except Exception:
            text_widget.insert(tk.END, "# 找不到現有字詞庫檔案，請輸入您的漢字列表：\n愛 麗 雅 靜\n風 雲 月 星")
            
        # 綁定按鍵釋放事件，實現即時統計更新
        text_widget.bind('<KeyRelease>', lambda event: self.update_config_stats_gui(text_widget.get("1.0", tk.END)))

        # 儲存按鈕
        save_button = tk.Button(editor_window, 
                                text="儲存字詞庫並重置索引",
                                command=lambda: self.save_and_reset_words(editor_window, text_widget.get("1.0", tk.END)),
                                font=('Microsoft JhengHei', 12),
                                bg="#FFB300")
        save_button.pack(pady=10)
        
        editor_window.protocol("WM_DELETE_WINDOW", lambda: self.confirm_close_editor(editor_window))

    
    def update_config_stats_gui(self, content):
        """功能 5: 更新字詞庫編輯器中的即時統計數據。"""
        word_count, pool_size = analyze_words_from_text(content)
        self.config_stats_var.set(f"當前輸入分析： 總字數(N): {word_count:,} 個 │ 總組合數(N x N): {pool_size:,} 個")

    def confirm_close_editor(self, window):
        if messagebox.askyesno("確認關閉", "您確定要關閉編輯器而不儲存變更嗎？"):
            window.destroy()

    def save_and_reset_words(self, window, new_content):
        try:
            with open(WORDS_FILE, 'w', encoding='utf-8') as f:
                f.write(new_content)
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
if __name__ == "__main__":

    setup_data_paths() # **【新增】在載入數據前設置路徑**
    
    load_master_words() # 步驟 1: 載入字詞庫
    
    # 步驟 2: 檢查索引狀態
    if not os.path.exists(STATE_FILE) and POOL_SIZE > 0:
        initialize_database(reset_history=True) # 首次啟動，生成索引
    else:
        # **【新】載入緩存，速度快！**
        load_indices_cache() 
        
    # 步驟 3: 啟動 Tkinter 介面
    root = tk.Tk()
    app = NameGeneratorApp(root)
    root.mainloop()