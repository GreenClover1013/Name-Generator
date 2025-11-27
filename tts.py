# tts.py - Windows-friendly pyttsx3 TTS manager (queue + single worker)
# 使用方法：
# from tts import speak_text, stop_worker
# 在 NameGeneratorApp.on_closing 呼叫 stop_worker() 再關閉視窗

import threading
import queue
import time
import os
import platform

try:
    import pyttsx3
    _PYTTSX3_AVAILABLE = True
except Exception:
    pyttsx3 = None
    _PYTTSX3_AVAILABLE = False

# 嘗試載入 comtypes 或 pythoncom 以支援 Windows COM 初始化
_COM_MODULE = None
if platform.system() == "Windows":
    try:
        import comtypes
        _COM_MODULE = "comtypes"
    except Exception:
        try:
            import pythoncom
            _COM_MODULE = "pythoncom"
        except Exception:
            _COM_MODULE = None

_TTS_QUEUE = queue.Queue()
_TTS_WORKER_THREAD = None
_TTS_STOP_EVENT = threading.Event()
_TTS_LOCK = threading.Lock()

# debug：設定環境變數 TTS_DEBUG=1 可看到簡單的輸出
_DEBUG = bool(os.environ.get("TTS_DEBUG", ""))

def _dprint(*args):
    if _DEBUG:
        print("[TTS]", *args)

def _select_chinese_voice(engine):
    try:
        voices = engine.getProperty('voices')
        for v in voices:
            vid = getattr(v, 'id', '') or ''
            name = getattr(v, 'name', '') or ''
            if any(k in vid.lower() for k in ('zh', 'chinese')) or any(k in name.lower() for k in ('zh', 'chinese', '國語', '普通話', '中文')):
                return v.id
    except Exception:
        pass
    return None

def _init_engine(rate, volume):
    """嘗試建立並設定 pyttsx3 engine，並在 Windows 執行緒上做 COM 初始化。"""
    if not _PYTTSX3_AVAILABLE:
        _dprint("pyttsx3 not available")
        return None, None

    coinit_used = None
    # Windows: 初始化 COM apartment 在本執行緒
    if platform.system() == "Windows" and _COM_MODULE:
        try:
            if _COM_MODULE == "comtypes":
                import comtypes
                comtypes.CoInitialize()
                coinit_used = "comtypes"
            else:
                import pythoncom
                pythoncom.CoInitialize()
                coinit_used = "pythoncom"
            _dprint("CoInitialize done via", coinit_used)
        except Exception as e:
            _dprint("CoInitialize failed:", e)
            coinit_used = None

    try:
        engine = pyttsx3.init()
        try:
            engine.setProperty('rate', rate)
        except Exception:
            pass
        try:
            engine.setProperty('volume', volume)
        except Exception:
            pass
        try:
            zh = _select_chinese_voice(engine)
            if zh:
                engine.setProperty('voice', zh)
        except Exception:
            pass
        return engine, coinit_used
    except Exception as e:
        _dprint("engine init failed:", e)
        # 若 engine init 失敗，嘗試在遇到異常時再 CoUninitialize
        return None, coinit_used

def _uninit_com_module(coinit_used):
    """在 Windows 上對應做 CoUninitialize（安全呼叫）。"""
    if not coinit_used:
        return
    try:
        if coinit_used == "comtypes":
            import comtypes
            comtypes.CoUninitialize()
        elif coinit_used == "pythoncom":
            import pythoncom
            pythoncom.CoUninitialize()
        _dprint("CoUninitialize done via", coinit_used)
    except Exception as e:
        _dprint("CoUninitialize failed:", e)

def _tts_worker(rate, volume):
    """worker：序列化處理佇列，遇錯嘗試重建 engine；確保 Windows COM 有被初始化/反初始化。"""
    engine, coinit_used = _init_engine(rate, volume)
    _dprint("worker started, engine:", bool(engine), "coinit:", coinit_used)
    while not _TTS_STOP_EVENT.is_set():
        try:
            txt = _TTS_QUEUE.get(timeout=0.5)
        except queue.Empty:
            continue

        if txt is None:
            try:
                _TTS_QUEUE.task_done()
            except Exception:
                pass
            break

        success = False
        for attempt in range(2):
            if engine is None:
                # 重新 init 引擎（同時嘗試再初始化 COM）
                engine, coinit_used = _init_engine(rate, volume)
                if engine is None:
                    time.sleep(0.05)
            if engine:
                try:
                    _dprint("saying:", txt)
                    engine.say(txt)
                    engine.runAndWait()
                    success = True
                    break
                except Exception as e:
                    _dprint("engine error during say/run:", e)
                    try:
                        engine.stop()
                    except Exception:
                        pass
                    # 丟棄 engine 並讓下一次嘗試重建
                    engine = None
                    # 短暫等待
                    time.sleep(0.05)
            else:
                time.sleep(0.05)

        try:
            _TTS_QUEUE.task_done()
        except Exception:
            pass

    # worker 結束前確保釋放 COM
    try:
        if engine:
            try:
                engine.stop()
            except Exception:
                pass
    except Exception:
        pass
    _uninit_com_module(coinit_used)
    _dprint("worker exiting")

def _ensure_worker(rate=160, volume=1.0):
    global _TTS_WORKER_THREAD
    with _TTS_LOCK:
        if _TTS_WORKER_THREAD is None or not _TTS_WORKER_THREAD.is_alive():
            _TTS_STOP_EVENT.clear()
            _TTS_WORKER_THREAD = threading.Thread(target=_tts_worker, args=(rate, volume), daemon=True)
            _TTS_WORKER_THREAD.start()
            _dprint("worker thread started")

def speak_text(text, rate=160, volume=1.0):
    if not text:
        return
    try:
        _ensure_worker(rate, volume)
        _TTS_QUEUE.put(str(text))
    except Exception as e:
        _dprint("speak_text enqueue failed:", e)

def stop_worker(timeout=1.0):
    try:
        _TTS_STOP_EVENT.set()
        _TTS_QUEUE.put(None)
        global _TTS_WORKER_THREAD
        if _TTS_WORKER_THREAD is not None:
            _TTS_WORKER_THREAD.join(timeout=timeout)
    except Exception:
        pass