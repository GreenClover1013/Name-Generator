# tts.py - pyttsx3 TTS manager（可中斷先前播放，優先播放最新請求）
# 使用： from tts import speak_text, stop_worker
# speak_text(text, interrupt=True)  -> 會中斷當前播放並立刻播放 text
# speak_text(text)                 -> 會排隊播放（不中斷）
# 請在 NameGeneratorApp.on_closing 中呼叫 stop_worker()

import threading
import queue
import time
import platform
import os

try:
    import pyttsx3
    _PYTTSX3_AVAILABLE = True
except Exception:
    pyttsx3 = None
    _PYTTSX3_AVAILABLE = False

# Windows COM helpers (若可用)
_pythoncom = None
_comtypes = None
if platform.system() == "Windows":
    try:
        import pythoncom
        _pythoncom = pythoncom
    except Exception:
        try:
            import comtypes
            _comtypes = comtypes
        except Exception:
            _pythoncom = None
            _comtypes = None

_TTS_QUEUE = queue.Queue()
_TTS_WORKER = None
_TTS_STOP = threading.Event()
_TTS_LOCK = threading.Lock()

# Engine lock & reference to current engine (for interrupt)
_ENGINE_LOCK = threading.Lock()
_CURRENT_ENGINE = None

_DEBUG = bool(os.environ.get("TTS_DEBUG", ""))

def _dprint(*args):
    if _DEBUG:
        print("[TTS]", *args)

def _select_chinese_voice(engine):
    try:
        voices = engine.getProperty("voices")
        for v in voices:
            vid = getattr(v, "id", "") or ""
            name = getattr(v, "name", "") or ""
            if any(k in vid.lower() for k in ("zh", "chinese")) or any(k in name.lower() for k in ("zh", "chinese", "國語", "普通話", "中文")):
                return v.id
    except Exception:
        pass
    return None

def _speak_once_internal(text, rate=160, volume=1.0):
    """在 worker 執行緒內建立 local engine，並將 engine 設為 _CURRENT_ENGINE，完成後清理。"""
    global _CURRENT_ENGINE
    if not _PYTTSX3_AVAILABLE:
        _dprint("pyttsx3 not available")
        return

    coinit_kind = None
    try:
        # Windows: 初始化 COM 在目前執行緒
        if platform.system() == "Windows":
            try:
                if _pythoncom is not None:
                    _pythoncom.CoInitialize()
                    coinit_kind = "pythoncom"
                elif _comtypes is not None:
                    _comtypes.CoInitialize()
                    coinit_kind = "comtypes"
                _dprint("CoInitialize:", coinit_kind)
            except Exception as e:
                _dprint("CoInitialize failed:", e)
                coinit_kind = None

        # 建立 engine
        try:
            engine = pyttsx3.init()
        except Exception as e:
            _dprint("pyttsx3.init failed:", e)
            engine = None

        if engine:
            try:
                engine.setProperty("rate", rate)
            except Exception:
                pass
            try:
                engine.setProperty("volume", volume)
            except Exception:
                pass
            try:
                zh = _select_chinese_voice(engine)
                if zh:
                    engine.setProperty("voice", zh)
            except Exception:
                pass

            # set current engine so main thread can call stop() to interrupt
            with _ENGINE_LOCK:
                _CURRENT_ENGINE = engine

            try:
                engine.say(text)
                engine.runAndWait()
            except Exception as e:
                _dprint("engine say/runAndWait error:", e)
            finally:
                try:
                    engine.stop()
                except Exception:
                    pass

    finally:
        # clear current engine
        with _ENGINE_LOCK:
            _CURRENT_ENGINE = None

        # Windows: uninit COM
        if coinit_kind == "pythoncom":
            try:
                _pythoncom.CoUninitialize()
                _dprint("CoUninitialize pythoncom")
            except Exception as e:
                _dprint("CoUninitialize failed:", e)
        elif coinit_kind == "comtypes":
            try:
                _comtypes.CoUninitialize()
                _dprint("CoUninitialize comtypes")
            except Exception as e:
                _dprint("CoUninitialize failed:", e)

def _worker_loop(rate=160, volume=1.0):
    _dprint("TTS worker starting")
    while not _TTS_STOP.is_set():
        try:
            item = _TTS_QUEUE.get(timeout=0.5)
        except queue.Empty:
            continue

        if item is None:
            try:
                _TTS_QUEUE.task_done()
            except Exception:
                pass
            break

        try:
            if isinstance(item, tuple) and len(item) == 3:
                text, r, v = item
            else:
                text, r, v = item, rate, volume
            _dprint("Worker speaking:", text)
            _speak_once_internal(text, r, v)
        except Exception as e:
            _dprint("Worker exception:", e)
        finally:
            try:
                _TTS_QUEUE.task_done()
            except Exception:
                pass

    _dprint("TTS worker exiting")

def _ensure_worker(rate=160, volume=1.0):
    global _TTS_WORKER
    with _TTS_LOCK:
        if _TTS_WORKER is None or not _TTS_WORKER.is_alive():
            _TTS_STOP.clear()
            _TTS_WORKER = threading.Thread(target=_worker_loop, args=(rate, volume), daemon=True)
            _TTS_WORKER.start()
            _dprint("TTS worker launched")

def _clear_queue():
    """清空佇列中尚未處理的項目（會對每個已入隊的 put 對應呼叫 task_done）。"""
    try:
        while True:
            item = _TTS_QUEUE.get_nowait()
            try:
                _TTS_QUEUE.task_done()
            except Exception:
                pass
    except queue.Empty:
        pass

def _interrupt_current_playback():
    """嘗試中斷當前播放（呼叫 engine.stop()），此函式 thread-safe。"""
    with _ENGINE_LOCK:
        eng = _CURRENT_ENGINE
    if eng:
        try:
            eng.stop()
        except Exception as e:
            _dprint("interrupt stop() failed:", e)

def speak_text(text, rate=160, volume=1.0, interrupt=False):
    """
    非阻塞：將發音請求放入佇列。
    如果 interrupt=True：清空隊列並中斷目前播放，保證新請求立刻成為下一個被播放的項目。
    """
    if not text:
        return
    try:
        _ensure_worker(rate, volume)
        if interrupt:
            # 優先中斷正在播放的 engine 並清空佇列的舊請求
            try:
                _interrupt_current_playback()
            except Exception:
                pass
            try:
                _clear_queue()
            except Exception:
                pass
        _TTS_QUEUE.put((str(text), rate, volume))
    except Exception as e:
        _dprint("speak_text enqueue failed:", e)

def stop_worker(timeout=1.0):
    """請在程式結束時呼叫以嘗試優雅停止 worker；也會嘗試中斷當前播放。"""
    try:
        _TTS_STOP.set()
        # 嘗試中斷當前播放
        try:
            _interrupt_current_playback()
        except Exception:
            pass
        _TTS_QUEUE.put(None)
        global _TTS_WORKER
        if _TTS_WORKER is not None:
            _TTS_WORKER.join(timeout=timeout)
    except Exception:
        pass