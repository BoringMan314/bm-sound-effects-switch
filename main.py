from __future__ import annotations

import atexit
import copy
import hashlib
import json
import os
import subprocess
import sys
import threading
import time
import webbrowser
import winsound
import ctypes
from ctypes import wintypes

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

import bm_single_instance

try:
    import keyboard                
except Exception:
    keyboard = None

                                          
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
user32 = ctypes.WinDLL("user32", use_last_error=True)

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000
WM_HOTKEY = 0x0312
PM_REMOVE = 0x0001

SINGLE_APP_ID = "bm-sound-effects-switch"
_app_mutex_handle = None


def _release_app_mutex() -> None:
    global _app_mutex_handle
    bm_single_instance.release_mutex(_app_mutex_handle)
    _app_mutex_handle = None


ABOUT_URL = "http://exnormal.com:81/"

                                                 
WINDOW_TITLE_PREFIX = "[B.M]"
WINDOW_TITLE_VERSION = "V1.0"
                      
WINDOW_TITLE_AUTHOR_SUFFIX = " By. [B.M] 圓周率 3.14"
DEFAULT_TK_FONT_SIZE = 10

BUILTIN_UI_LANGUAGES = ("zh_TW", "zh_CN", "ja_JP", "en_US")


def apply_default_font_size(root: tk.Tk) -> None:
    for name in (
        "TkDefaultFont",
        "TkTextFont",
        "TkFixedFont",
        "TkMenuFont",
        "TkHeadingFont",
        "TkCaptionFont",
        "TkSmallCaptionFont",
        "TkIconFont",
        "TkTooltipFont",
    ):
        try:
            tkfont.nametofont(name).configure(size=DEFAULT_TK_FONT_SIZE)
        except tk.TclError:
            pass


                                                                                           
_DEFAULT_ZH_TW = {
    "language_name": "繁體中文",
    "settings": "設定",
    "project_name": "音效切換器",
    "combo_empty": "（未選擇）",
    "offline_prefix": "[離線] ",
    "playback_devices": "播放裝置",
    "recording_devices": "錄製裝置",
    "refresh": "重新整理",
    "set_default_playback": "設為系統預設播放",
    "set_default_recording": "設為系統預設錄製",
    "group_settings_frame": "組別設定",
    "col_group": "組",
    "col_playback": "播放裝置",
    "col_recording": "錄製裝置",
    "col_hotkey": "快捷鍵",
    "col_apply": "套用組合",
    "col_save": "儲存組合",
    "reset_group": "重設組合",
    "reset_group_button": "重設",
    "apply_this_group": "套用",
    "save_group": "儲存",
    "open_system_sound": "開啟系統音效",
    "open_sound_console": "開聲音控制台",
    "lang_cycle": "語言",
    "err_title": "錯誤",
    "err_save": "儲存失敗",
    "err_mutex": "無法取得單一實例鎖定，請稍後再試。",
    "err_enumerate_playback": "無法列舉播放裝置：{e}",
    "err_enumerate_record": "無法列舉錄製裝置：{e}",
    "tip_title": "提示",
    "done_title": "完成",
    "fail_title": "失敗",
    "tip_select_playback": "請先選取一個播放裝置。",
    "tip_select_record": "請先選取一個錄製裝置。",
    "tip_offline_playback": "此裝置目前離線，無法設為預設。",
    "tip_offline_record": "此裝置目前離線，無法設為預設。",
    "done_playback": "已切換系統預設播放裝置。",
    "done_record": "已切換系統預設錄製裝置。",
    "tip_group_empty": "此組尚未儲存或皆為未選擇。",
    "done_apply_group": "已套用組 {n}。",
    "saved_title": "已儲存",
    "msg_saved_group": "組 {n} 已寫入設定檔。",
    "tip_max_groups": "最多 {n} 組。",
    "tip_min_groups": "至少保留 {n} 組。",
    "hotkey_capture_title": "設定快捷鍵",
    "hotkey_listen_inline": "按下組合鍵…",
    "hotkey_press_prompt": "請按下要使用的組合鍵（Esc 取消）",
    "hotkey_click_hint": "左鍵設定，右鍵清除",
    "tip_hotkey_duplicate": "此組合與第 {n} 組相同。",
    "tip_hotkey_register_failed": "無法註冊快捷鍵（可能與其他程式衝突）。",
    "tip_hotkey_need_modifier": "請至少按住 Ctrl、Alt、Shift 或 Win 之一再按一鍵。",
    "about": "關於",
    "exit": "離開",
}
_DEFAULT_ZH_CN = {
    "language_name": "简体中文",
    "settings": "设置",
    "project_name": "音效切换器",
    "combo_empty": "（未选择）",
    "offline_prefix": "[离线] ",
    "playback_devices": "播放设备",
    "recording_devices": "录制设备",
    "refresh": "刷新",
    "set_default_playback": "设为系统默认播放",
    "set_default_recording": "设为系统默认录制",
    "group_settings_frame": "组设置",
    "col_group": "组",
    "col_playback": "播放设备",
    "col_recording": "录制设备",
    "col_hotkey": "快捷键",
    "col_apply": "应用组合",
    "col_save": "保存组合",
    "reset_group": "重置组合",
    "reset_group_button": "重置",
    "apply_this_group": "应用",
    "save_group": "保存",
    "open_system_sound": "打开系统声音",
    "open_sound_console": "打开声音控制台",
    "lang_cycle": "语言",
    "err_title": "错误",
    "err_save": "保存失败",
    "err_mutex": "无法取得单实例锁定，请稍后再试。",
    "err_enumerate_playback": "无法枚举播放设备：{e}",
    "err_enumerate_record": "无法枚举录制设备：{e}",
    "tip_title": "提示",
    "done_title": "完成",
    "fail_title": "失败",
    "tip_select_playback": "请先选择一个播放设备。",
    "tip_select_record": "请先选择一个录制设备。",
    "tip_offline_playback": "此设备当前离线，无法设为默认。",
    "tip_offline_record": "此设备当前离线，无法设为默认。",
    "done_playback": "已切换系统默认播放设备。",
    "done_record": "已切换系统默认录制设备。",
    "tip_group_empty": "此组尚未保存或均为未选择。",
    "done_apply_group": "已应用组 {n}。",
    "saved_title": "已保存",
    "msg_saved_group": "组 {n} 已写入配置文件。",
    "tip_max_groups": "最多 {n} 组。",
    "tip_min_groups": "至少保留 {n} 组。",
    "hotkey_capture_title": "设置快捷键",
    "hotkey_listen_inline": "按下组合键…",
    "hotkey_press_prompt": "请按下组合键（Esc 取消）",
    "hotkey_click_hint": "左键设置，右键清除",
    "tip_hotkey_duplicate": "此组合与第 {n} 组相同。",
    "tip_hotkey_register_failed": "无法注册快捷键（可能与其他程序冲突）。",
    "tip_hotkey_need_modifier": "请至少按住 Ctrl、Alt、Shift 或 Win 之一再按键。",
    "about": "关于",
    "exit": "离开",
}
_DEFAULT_JA = {
    "language_name": "日本語",
    "settings": "設定",
    "project_name": "サウンド切替",
    "combo_empty": "（未選択）",
    "offline_prefix": "[オフライン] ",
    "playback_devices": "再生デバイス",
    "recording_devices": "録音デバイス",
    "refresh": "更新",
    "set_default_playback": "既定の再生に設定",
    "set_default_recording": "既定の録音に設定",
    "group_settings_frame": "プリセット設定",
    "col_group": "番号",
    "col_playback": "再生",
    "col_recording": "録音",
    "col_hotkey": "ショートカット",
    "col_apply": "組み合わせを適用",
    "col_save": "組み合わせを保存",
    "reset_group": "組み合わせをリセット",
    "reset_group_button": "リセット",
    "apply_this_group": "適用",
    "save_group": "保存",
    "open_system_sound": "サウンド設定を開く",
    "open_sound_console": "サウンド（再生デバイス）",
    "lang_cycle": "言語",
    "err_title": "エラー",
    "err_save": "保存に失敗しました",
    "err_mutex": "単一インスタンスのロックを取得できません。",
    "err_enumerate_playback": "再生デバイスを列挙できません：{e}",
    "err_enumerate_record": "録音デバイスを列挙できません：{e}",
    "tip_title": "ヒント",
    "done_title": "完了",
    "fail_title": "失敗",
    "tip_select_playback": "再生デバイスを選択してください。",
    "tip_select_record": "録音デバイスを選択してください。",
    "tip_offline_playback": "オフラインのため既定にできません。",
    "tip_offline_record": "オフラインのため既定にできません。",
    "done_playback": "既定の再生デバイスを変更しました。",
    "done_record": "既定の録音デバイスを変更しました。",
    "tip_group_empty": "未保存か、どちらも未選択です。",
    "done_apply_group": "プリセット {n} を適用しました。",
    "saved_title": "保存済み",
    "msg_saved_group": "プリセット {n} を保存しました。",
    "tip_max_groups": "最大 {n} 組です。",
    "tip_min_groups": "最低 {n} 組は必要です。",
    "hotkey_capture_title": "ショートカット設定",
    "hotkey_listen_inline": "キーを押してください…",
    "hotkey_press_prompt": "キーの組み合わせを押してください（Esc で取消）",
    "hotkey_click_hint": "左クリックで設定・右クリックで解除",
    "tip_hotkey_duplicate": "プリセット {n} と同じ組み合わせです。",
    "tip_hotkey_register_failed": "ショートカットを登録できません（他アプリと競合している可能性があります）。",
    "tip_hotkey_need_modifier": "Ctrl / Alt / Shift / Win のいずれかを押しながらキーを押してください。",
    "about": "バージョン情報",
    "exit": "終了",
}
_DEFAULT_EN = {
    "language_name": "English",
    "settings": "Settings",
    "project_name": "Sound Switcher",
    "combo_empty": "(None)",
    "offline_prefix": "[Offline] ",
    "playback_devices": "Playback devices",
    "recording_devices": "Recording devices",
    "refresh": "Refresh",
    "set_default_playback": "Set as default playback",
    "set_default_recording": "Set as default recording",
    "group_settings_frame": "Preset settings",
    "col_group": "#",
    "col_playback": "Playback",
    "col_recording": "Recording",
    "col_hotkey": "Hotkey",
    "col_apply": "Apply combination",
    "col_save": "Save combination",
    "reset_group": "Reset combination",
    "reset_group_button": "Reset",
    "apply_this_group": "Apply",
    "save_group": "Save",
    "open_system_sound": "System sound (Settings)",
    "open_sound_console": "Sound control panel",
    "lang_cycle": "Language",
    "err_title": "Error",
    "err_save": "Save failed",
    "err_mutex": "Could not acquire single-instance lock.",
    "err_enumerate_playback": "Cannot list playback devices: {e}",
    "err_enumerate_record": "Cannot list recording devices: {e}",
    "tip_title": "Notice",
    "done_title": "Done",
    "fail_title": "Failed",
    "tip_select_playback": "Select a playback device first.",
    "tip_select_record": "Select a recording device first.",
    "tip_offline_playback": "Device is offline; cannot set as default.",
    "tip_offline_record": "Device is offline; cannot set as default.",
    "done_playback": "Default playback device changed.",
    "done_record": "Default recording device changed.",
    "tip_group_empty": "Preset not saved or both are unset.",
    "done_apply_group": "Applied preset {n}.",
    "saved_title": "Saved",
    "msg_saved_group": "Preset {n} saved to config file.",
    "tip_max_groups": "At most {n} presets.",
    "tip_min_groups": "Keep at least {n} presets.",
    "hotkey_capture_title": "Set hotkey",
    "hotkey_listen_inline": "Press keys…",
    "hotkey_press_prompt": "Press the key combination (Esc to cancel)",
    "hotkey_click_hint": "Left-click to set, right-click to clear",
    "tip_hotkey_duplicate": "Same combination as preset {n}.",
    "tip_hotkey_register_failed": "Could not register hotkey (may conflict with another app).",
    "tip_hotkey_need_modifier": "Hold Ctrl, Alt, Shift, or Win, then press a key.",
    "about": "About",
    "exit": "Exit",
}

DEFAULT_TRANSLATIONS: dict[str, dict[str, str]] = {
    "zh_TW": dict(_DEFAULT_ZH_TW),
    "zh_CN": dict(_DEFAULT_ZH_CN),
    "ja_JP": dict(_DEFAULT_JA),
    "en_US": dict(_DEFAULT_EN),
}

SOUND_I18N_REF_KEYS = frozenset(_DEFAULT_ZH_TW.keys())

_singleton_window_title: str = ""

MAX_GROUPS = 10
MIN_GROUPS = 2
DEFAULT_NUM_GROUPS = 2
               
DEVICE_LIST_VISIBLE_ROWS = 5
                                        
GROUPS_VIEWPORT_HEIGHT = 150
                                                        
CONFIG_BASENAME_MAIN = "bm-sound-effects-switch.json"

WINDOW_WIDTH = 800
WINDOW_HEIGHT = 510
WINDOW_POS_X = 100
WINDOW_POS_Y = 100


def _primary_screen_origin() -> tuple[int, int]:
    if sys.platform != "win32":
        return 0, 0
    try:
        MONITOR_DEFAULTTOPRIMARY = 1

        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long),
            ]

        class MONITORINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.c_ulong),
                ("rcMonitor", RECT),
                ("rcWork", RECT),
                ("dwFlags", ctypes.c_ulong),
            ]

        m = user32.MonitorFromPoint(wintypes.POINT(0, 0), MONITOR_DEFAULTTOPRIMARY)
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(MONITORINFO)
        if m and user32.GetMonitorInfoW(m, ctypes.byref(mi)):
            return int(mi.rcMonitor.left), int(mi.rcMonitor.top)
    except Exception:
        pass
    return 0, 0


def resource_path(*parts: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def app_bundle_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def config_basename() -> str:
    return CONFIG_BASENAME_MAIN


def config_path() -> str:
    return os.path.join(app_bundle_dir(), config_basename())


def merged_lang_table(cfg: dict) -> dict[str, dict[str, str]]:
    base = copy.deepcopy(DEFAULT_TRANSLATIONS)
    raw_t = cfg.get("languages")
    if isinstance(raw_t, dict):
        for lang, kv in raw_t.items():
            if isinstance(lang, str) and isinstance(kv, dict):
                if lang not in base:
                    base[lang] = {}
                base[lang].update(kv)
    return base


def get_current_lang(cfg: dict) -> str:
    settings = cfg.get("settings")
    if isinstance(settings, dict):
        language = settings.get("languages")
        if isinstance(language, str) and language.strip():
            return language.strip()
    return "zh_TW"


def set_current_lang(cfg: dict, lang: str) -> None:
    settings = cfg.get("settings")
    if not isinstance(settings, dict):
        settings = {}
    settings["languages"] = lang
    cfg["settings"] = settings


def available_ui_languages(cfg: dict) -> list[str]:
    mt = merged_lang_table(cfg)
    ordered: list[str] = [x for x in BUILTIN_UI_LANGUAGES if x in mt]
    extras = [k for k in mt.keys() if k not in BUILTIN_UI_LANGUAGES]
    ordered.extend(extras)
    if not ordered:
        return ["zh_TW"]
    return ordered


def lang_button_label(cfg: dict, lang: str) -> str:
    mt = merged_lang_table(cfg)
    name = mt.get(lang, {}).get("language_name")
    if isinstance(name, str) and name.strip():
        return name.strip()
    return lang


def tr_cfg(cfg: dict, key: str, **kwargs: object) -> str:
    langs = available_ui_languages(cfg)
    lang = get_current_lang(cfg)
    if lang not in langs:
        lang = "zh_TW"
    mt = merged_lang_table(cfg)
    order = [lang] + [x for x in langs if x != lang]
    if "zh_TW" not in order:
        order.append("zh_TW")
    s = None
    for L in order:
        v = mt.get(L, {}).get(key)
        if isinstance(v, str) and v:
            s = v
            break
    if s is None:
        s = DEFAULT_TRANSLATIONS["zh_TW"].get(key, key)
    if kwargs:
        try:
            return s.format(**kwargs)
        except (KeyError, ValueError):
            return s
    return s


def window_title_product_only(s: str) -> str:
    return (s or "").rstrip()


def window_title_full(cfg: dict) -> str:
    product = window_title_product_only(tr_cfg(cfg, "project_name"))
    return (
        f"{WINDOW_TITLE_PREFIX} {product} {WINDOW_TITLE_VERSION}"
        + WINDOW_TITLE_AUTHOR_SUFFIX
    )


def _empty_group() -> dict:
    return {
        "playback_id": None,
        "recording_id": None,
        "playback_name": "",
        "recording_name": "",
                                                              
        "hotkey_mods": None,
        "hotkey_vk": None,
    }


def effective_group_hotkey(group: dict, index: int) -> tuple[int, int] | None:
    hm, hv = group.get("hotkey_mods"), group.get("hotkey_vk")
    if hm is not None and hv is not None:
        mi, vi = int(hm), int(hv)
        if mi == 0 and vi == 0:
            return None
        return mi, vi
    if index < 9:
        return MOD_CONTROL | MOD_ALT, ord("1") + index
    if index == 9:
        return MOD_CONTROL | MOD_ALT, ord("0")
    return None


def vk_display_label(vk: int) -> str:
    if 0x30 <= vk <= 0x39:
        return chr(vk)
    if 0x41 <= vk <= 0x5A:
        return chr(vk)
    if 0x60 <= vk <= 0x69:
        return "Num" + str(vk - 0x60)
    if 0x70 <= vk <= 0x87:
        return "F" + str(vk - 0x6F)
    return f"VK{vk:02X}"


def vk_to_keyboard_key(vk: int) -> str | None:
                                                    
    if 0x30 <= vk <= 0x39:       
        return chr(vk).lower()
    if 0x41 <= vk <= 0x5A:       
        return chr(vk).lower()
    if 0x70 <= vk <= 0x87:          
        return f"f{vk - 0x6F}"
    if 0x60 <= vk <= 0x69:              
        return f"num {vk - 0x60}"
    return None


def mods_vk_to_keyboard_combos(mods: int, vk: int) -> list[str]:
    key = vk_to_keyboard_key(vk)
    if not key:
        return []
    parts: list[str] = []
    if mods & MOD_CONTROL:
        parts.append("ctrl")
    if mods & MOD_SHIFT:
        parts.append("shift")
    if mods & MOD_ALT:
        parts.append("alt")
    if not (parts or (mods & MOD_WIN)):
        return []
    base_parts = list(parts)
    combos: list[str] = []
    if mods & MOD_WIN:
        for win_key in ("windows", "left windows", "right windows", "win"):
            combos.append("+".join(base_parts + [win_key, key]))
    else:
        combos.append("+".join(base_parts + [key]))
    return combos


def hotkey_display_text(mods: int, vk: int) -> str:
    parts: list[str] = []
    if mods & MOD_CONTROL:
        parts.append("Ctrl")
    if mods & MOD_SHIFT:
        parts.append("Shift")
    if mods & MOD_ALT:
        parts.append("Alt")
    if mods & MOD_WIN:
        parts.append("Win")
    parts.append(vk_display_label(vk))
    return "+".join(parts)


def tk_keyevent_to_mods_vk(event: tk.Event) -> tuple[int, int] | None:
    ks = event.keysym
    if ks in (
        "Shift_L",
        "Shift_R",
        "Control_L",
        "Control_R",
        "Alt_L",
        "Alt_R",
        "Win_L",
        "Win_R",
        "Caps_Lock",
    ):
        return None
    if ks == "Escape":
        return None
    state = int(event.state)
    mods = 0
    if state & 0x0004:
        mods |= MOD_CONTROL
    if state & 0x0001:
        mods |= MOD_SHIFT
    if state & 0x20000:
        mods |= MOD_ALT
    if state & 0x40000:
        mods |= MOD_WIN
                                                                                    
    try:
        if user32.GetAsyncKeyState(0xA2) & 0x8000 or user32.GetAsyncKeyState(0xA3) & 0x8000:
            mods |= MOD_CONTROL
        if user32.GetAsyncKeyState(0xA0) & 0x8000 or user32.GetAsyncKeyState(0xA1) & 0x8000:
            mods |= MOD_SHIFT
        if user32.GetAsyncKeyState(0xA4) & 0x8000 or user32.GetAsyncKeyState(0xA5) & 0x8000:
            mods |= MOD_ALT
        if user32.GetAsyncKeyState(0x5B) & 0x8000 or user32.GetAsyncKeyState(0x5C) & 0x8000:
            mods |= MOD_WIN
    except Exception:
        pass
    try:
        vk = int(event.keycode)
    except (TypeError, ValueError):
        return None
    if vk < 0 or vk > 255:
        return None
    if vk in (16, 17, 18, 91, 92, 93, 144, 145):
        return None
    if mods == 0:
        return None
    return mods, vk


def default_config() -> dict:
    return {
        "settings": {"languages": "zh_TW"},
        "languages": copy.deepcopy(DEFAULT_TRANSLATIONS),
        "num_groups": DEFAULT_NUM_GROUPS,
        "groups": [_empty_group() for _ in range(MAX_GROUPS)],
    }


def _is_valid_sound_config(raw: dict) -> bool:
    if not isinstance(raw, dict):
        return False
    st = raw.get("settings")
    if not isinstance(st, dict):
        return False
    code = st.get("languages")
    if not isinstance(code, str) or not code.strip():
        return False
    code = code.strip()
    lg = raw.get("languages")
    if not isinstance(lg, dict) or not lg or code not in lg:
        return False
    for m in lg.values():
        if not isinstance(m, dict):
            return False
        if frozenset(m.keys()) != SOUND_I18N_REF_KEYS:
            return False
        for rk in SOUND_I18N_REF_KEYS:
            v = m.get(rk)
            if not isinstance(v, str):
                return False
            if rk in ("language_name", "project_name", "settings") and not v.strip():
                return False
    num = raw.get("num_groups")
    if not isinstance(num, int) or not (MIN_GROUPS <= num <= MAX_GROUPS):
        return False
    gr = raw.get("groups")
    if not isinstance(gr, list) or len(gr) != MAX_GROUPS:
        return False
    return True


def load_config() -> dict:
    p = config_path()
    data = default_config()
    if not os.path.isfile(p):
        return data
    try:
        with open(p, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except (OSError, json.JSONDecodeError):
        try:
            os.remove(p)
        except OSError:
            pass
        save_config(data)
        return data
    if not _is_valid_sound_config(raw):
        try:
            os.remove(p)
        except OSError:
            pass
        save_config(data)
        return data
    old_groups = raw.get("groups")
    if isinstance(old_groups, list):
        for i in range(MAX_GROUPS):
            if i < len(old_groups) and isinstance(old_groups[i], dict):
                g = _empty_group()
                g.update(old_groups[i])
                data["groups"][i] = g
    if (
        "num_groups" in raw
        and isinstance(raw["num_groups"], int)
        and MIN_GROUPS <= raw["num_groups"] <= MAX_GROUPS
    ):
        data["num_groups"] = raw["num_groups"]
    else:
                                    
        data["num_groups"] = DEFAULT_NUM_GROUPS
    if data["num_groups"] < MIN_GROUPS:
        data["num_groups"] = MIN_GROUPS
    data["languages"] = merged_lang_table(raw)
    raw_settings = raw.get("settings")
    ul = raw_settings.get("languages") if isinstance(raw_settings, dict) else None
    langs = available_ui_languages(data)
    if isinstance(ul, str) and ul in langs:
        set_current_lang(data, ul)
    else:
        set_current_lang(data, "zh_TW" if "zh_TW" in langs else langs[0])
    return data


def save_config(data: dict) -> None:
    try:
        set_current_lang(data, get_current_lang(data))
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def play_success_sound() -> None:
    sound_path = os.path.abspath(os.path.normpath(resource_path("wav", "switch.wav")))
    if not os.path.isfile(sound_path):
        return
    try:
        winsound.PlaySound(
            sound_path,
            winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_NODEFAULT,
        )
    except Exception:
        pass


                  
from pycaw.constants import DEVICE_STATE, EDataFlow, ERole
from pycaw.pycaw import AudioUtilities

ALL_ROLES = [ERole.eConsole, ERole.eMultimedia, ERole.eCommunications]


def _coerce_int(value: object) -> int | None:
    try:
        return int(value)                          
    except Exception:
        return None


def _device_state_raw(dev: object) -> int | None:
    st = getattr(dev, "state", None)
    if st is None:
        for attr in ("State", "DeviceState", "device_state"):
            st = getattr(dev, attr, None)
            if st is not None:
                break
    if st is None:
        return None
    if hasattr(st, "value"):
        return _coerce_int(st.value)
    return _coerce_int(st)


def _device_state_ok_for_sound_panel(dev: object) -> bool:
    s = _device_state_raw(dev)
    if s is None:
        return True
    if s & DEVICE_STATE.DISABLED.value:
        return False
    if s & DEVICE_STATE.NOTPRESENT.value:
        return False
    if s & DEVICE_STATE.UNPLUGGED.value:
        return False
    return bool(s & DEVICE_STATE.ACTIVE.value)


def _device_ui_hidden_from_sound_panel(dev: object) -> bool:
    props = getattr(dev, "properties", None)
    if not isinstance(props, dict):
        return False
    for k, v in props.items():
        ks = str(k).lower()
        if "f3e75c5c-11e8-4e8a-9ce2-9d46438a6176" not in ks:
            continue
        try:
            return bool(v)
        except Exception:
            return False
    return False


def _endpoint_data_flow(dev: object) -> int | None:
    dev_id = getattr(dev, "id", None)
    if dev_id:
        try:
            return int(AudioUtilities.GetEndpointDataFlow(dev_id, 1))
        except Exception:
            pass
    raw = getattr(dev, "_dev", None)
    if raw is not None:
        try:
            from pycaw.api.mmdeviceapi import IMMEndpoint

            return int(raw.QueryInterface(IMMEndpoint).GetDataFlow())
        except Exception:
            pass
    for attr in ("DataFlow", "dataFlow", "data_flow", "Flow"):
        v = _coerce_int(getattr(dev, attr, None))
        if v is not None:
            return v
    return None


def _device_matches_flow(dev: object, flow_value: int, enum_includes_all_flows: bool) -> bool:
    v = _endpoint_data_flow(dev)
    if v is not None:
        return v == flow_value
    return not enum_includes_all_flows


def _normalize_audio_device(dev: object) -> object | None:
    if dev is None:
        return None
    if hasattr(dev, "id"):
        return dev
    try:
        return AudioUtilities.CreateDevice(dev)
    except Exception:
        return None


def _normalize_device_list(devices: list) -> list:
    out: list = []
    for d in devices:
        nd = _normalize_audio_device(d)
        if nd is not None:
            out.append(nd)
    return out


def _device_ok_for_sound_list(
    dev: object, flow_value: int, enum_includes_all_flows: bool
) -> bool:
    if _device_ui_hidden_from_sound_panel(dev):
        return False
    if not _device_state_ok_for_sound_panel(dev):
        return False
    return _device_matches_flow(dev, flow_value, enum_includes_all_flows)


def _get_all_devices_compat(flow_value: int, active_state: int) -> list:
    prefer_zero_arg = False
    if sys.platform == "win32":
        try:
            prefer_zero_arg = sys.getwindowsversion().major < 10
        except Exception:
            prefer_zero_arg = False

    def _filter(raw_list: list, enum_includes_all_flows: bool) -> list:
        normalized = _normalize_device_list(raw_list)
        return [
            d
            for d in normalized
            if _device_ok_for_sound_list(d, flow_value, enum_includes_all_flows)
        ]

    raw: list = []
    enum_mixed = False
    if prefer_zero_arg:
        try:
            raw = AudioUtilities.GetAllDevices()
            enum_mixed = True
        except Exception:
            raw = []
        out = _filter(raw, enum_mixed)
        if not out:
            try:
                raw = AudioUtilities.GetAllDevices(flow_value, active_state)
            except Exception:
                raw = []
            out = _filter(raw, False)
        return out

    try:
        raw = AudioUtilities.GetAllDevices(flow_value, active_state)
    except Exception:
        try:
            raw = AudioUtilities.GetAllDevices()
            enum_mixed = True
        except Exception:
            raw = []
    out = _filter(raw, enum_mixed)
    if not out and not enum_mixed:
        try:
            raw = AudioUtilities.GetAllDevices()
            out = _filter(raw, True)
        except Exception:
            pass
    return out


def list_playback_devices():
    return _get_all_devices_compat(EDataFlow.eRender.value, DEVICE_STATE.ACTIVE.value)


def list_recording_devices():
    return _get_all_devices_compat(EDataFlow.eCapture.value, DEVICE_STATE.ACTIVE.value)


def default_playback_device():
    try:
        return _normalize_audio_device(AudioUtilities.GetSpeakers())
    except Exception:
        return None


def default_recording_device():
    try:
        mic = AudioUtilities.GetMicrophone()
        if mic is None:
            return None
        return _normalize_audio_device(mic)
    except Exception:
        return None


def _mmdevice_ids_equal(a: str | None, b: str | None) -> bool:
    if not a or not b:
        return False
    return a.strip().lower() == b.strip().lower()


def _default_endpoint_id_for_flow_and_role(flow_int: int, role) -> str | None:
    try:
        de = AudioUtilities.GetDeviceEnumerator()
        ep = de.GetDefaultAudioEndpoint(flow_int, role.value)
        if ep is None:
            return None
        return ep.GetId()
    except Exception:
        return None


def _default_ids_for_playback() -> set[str]:
    out: set[str] = set()
    for role in (ERole.eMultimedia, ERole.eConsole):
        did = _default_endpoint_id_for_flow_and_role(EDataFlow.eRender.value, role)
        if did:
            out.add(did)
    return out


def _default_ids_for_recording() -> set[str]:
    out: set[str] = set()
    for role in (ERole.eMultimedia, ERole.eConsole):
        did = _default_endpoint_id_for_flow_and_role(EDataFlow.eCapture.value, role)
        if did:
            out.add(did)
    return out


def _device_row_is_default(did: str, default_ids: set[str]) -> bool:
    if not default_ids:
        return False
    for x in default_ids:
        if _mmdevice_ids_equal(did, x):
            return True
    return False


def set_default_playback(dev_id: str) -> None:
    AudioUtilities.SetDefaultDevice(dev_id, ALL_ROLES)


def set_default_recording(dev_id: str) -> None:
    AudioUtilities.SetDefaultDevice(dev_id, ALL_ROLES)


def _safe_tree_iid(dev_id: str) -> str:
    return "d" + hashlib.md5(dev_id.encode("utf-8", errors="replace")).hexdigest()[:20]


class SoundSwitcherApp:
    def tr(self, key: str, **kwargs: object) -> str:
        return tr_cfg(self.cfg, key, **kwargs)

    def _win_title(self) -> str:
        return window_title_full(self.cfg)

    def _move_to_default_position(self) -> None:
        ox, oy = _primary_screen_origin()
        self.root.geometry(
            f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{ox + WINDOW_POS_X}+{oy + WINDOW_POS_Y}"
        )

    def __init__(self) -> None:
        self.root = tk.Tk()
        apply_default_font_size(self.root)
        self.cfg = load_config()
                                                              
        try:
            if not os.path.isfile(config_path()):
                save_config(self.cfg)
        except Exception:
            pass
        _ng = int(self.cfg["num_groups"])
        self.num_groups = max(MIN_GROUPS, _ng)
        if self.num_groups != _ng:
            self.cfg["num_groups"] = self.num_groups
            save_config(self.cfg)
        global _singleton_window_title
        self.root.title(self._win_title())
        _singleton_window_title = self._win_title()
        self._set_app_user_model_id()
        self._set_app_user_model_id()
        self.root.resizable(False, False)
        self.root.minsize(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.root.maxsize(WINDOW_WIDTH, WINDOW_HEIGHT)
        self._move_to_default_position()
        self.root.grid_columnconfigure(0, weight=1)
        self._hotkey_hwnd = None
        self._hotkey_ids = list(range(1, MAX_GROUPS + 1))
        self._hotkey_capture_index: int | None = None
        self._hotkey_clickout_bind_id: str | None = None
        self.group_hotkey_entries: list[ttk.Entry] = []

        self._device_name_to_id_play: dict[str, str] = {}
        self._device_name_to_id_rec: dict[str, str] = {}
        self._pb_iid_meta: dict[str, tuple[str, bool]] = {}
        self._rec_iid_meta: dict[str, tuple[str, bool]] = {}

        self._tray_icon = None
        self._hiding_to_tray = False
        self._apply_window_icon()

        self._build_ui()
        self._disable_button_focus_rect(self.root)
        self._install_ttk_focus_ring_mitigation()
        self.refresh_playback_list()
        self.refresh_recording_list()
        self._rebuild_groups_ui()
        self._load_group_fields_from_config()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Unmap>", self._on_window_unmap, add="+")
        self.root.bind_all("<KeyPress-space>", self._guard_space_key, add="+")
        self.root.bind_all("<KeyRelease-space>", self._guard_space_key, add="+")
        self.root.bind_class("TButton", "<KeyPress-space>", lambda e: "break", add="+")
        self.root.bind_class("TButton", "<KeyRelease-space>", lambda e: "break", add="+")
        self.root.bind_class("Button", "<KeyPress-space>", lambda e: "break", add="+")
        self.root.bind_class("Button", "<KeyRelease-space>", lambda e: "break", add="+")
        self.root.bind_class("TCheckbutton", "<KeyPress-space>", lambda e: "break", add="+")
        self.root.bind_class("TCheckbutton", "<KeyRelease-space>", lambda e: "break", add="+")
        self.root.bind_class("Checkbutton", "<KeyPress-space>", lambda e: "break", add="+")
        self.root.bind_class("Checkbutton", "<KeyRelease-space>", lambda e: "break", add="+")
        self._bind_space_block_recursive(self.root)
        atexit.register(self._cleanup_hotkeys)
        self.root.after(80, self._setup_hotkeys)
        self._start_tray()

    def _apply_window_icon(self) -> None:
        ico = os.path.abspath(os.path.normpath(resource_path("icons", "icon.ico")))
        if not os.path.isfile(ico):
            return
        try:
                                                 
            self.root.iconbitmap(default=ico)
        except tk.TclError:
            try:
                self.root.iconbitmap(ico)
            except tk.TclError:
                pass

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=8)
        main.grid(row=0, column=0, sticky=tk.NW)
        main.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Treeview", rowheight=22)
        style.configure("RedRO.TCombobox", fieldbackground="#ffcccc")
        style.map(
            "RedRO.TCombobox",
            fieldbackground=[("readonly", "#ffcccc")],
        )

        r = 0
                    
        self._lbl_pb_title = ttk.Label(main, text=self.tr("playback_devices"))
        self._lbl_pb_title.grid(row=r, column=0, sticky=tk.W)
        r += 1
        self._pb_frame = ttk.Frame(main)
        self._pb_frame.grid(row=r, column=0, sticky=tk.NSEW, pady=(0, 4))
        self._pb_frame.grid_columnconfigure(0, weight=1)
        self.pb_tree = ttk.Treeview(
            self._pb_frame,
            show="tree",
            height=DEVICE_LIST_VISIBLE_ROWS,
            selectmode="browse",
        )
        self.pb_tree.column("#0", width=WINDOW_WIDTH - 20, stretch=True)
        self.pb_sb = ttk.Scrollbar(
            self._pb_frame, orient=tk.VERTICAL, command=self.pb_tree.yview
        )
        self.pb_tree.configure(yscrollcommand=self._pb_sb_set)
        self.pb_tree.grid(row=0, column=0, sticky=tk.NSEW)

        bf = ttk.Frame(main)
        bf.grid(row=r + 1, column=0, sticky=tk.EW, pady=4)
        self._btn_pb_refresh = ttk.Button(
            bf, text=self.tr("refresh"), command=self.refresh_playback_list
        )
        self._btn_pb_refresh.pack(side=tk.LEFT, padx=2)
        self._btn_pb_default = ttk.Button(
            bf, text=self.tr("set_default_playback"), command=self.apply_selected_playback
        )
        self._btn_pb_default.pack(side=tk.LEFT, padx=2)
        r += 2

                    
        self._lbl_rec_title = ttk.Label(
            main,
            text=self.tr("recording_devices"),
        )
        self._lbl_rec_title.grid(row=r, column=0, sticky=tk.W, pady=(12, 0))
        r += 1
        self._rec_frame = ttk.Frame(main)
        self._rec_frame.grid(row=r, column=0, sticky=tk.NSEW, pady=(0, 4))
        self._rec_frame.grid_columnconfigure(0, weight=1)
        self.rec_tree = ttk.Treeview(
            self._rec_frame,
            show="tree",
            height=DEVICE_LIST_VISIBLE_ROWS,
            selectmode="browse",
        )
        self.rec_tree.column("#0", width=WINDOW_WIDTH - 20, stretch=True)
        self.rec_sb = ttk.Scrollbar(
            self._rec_frame, orient=tk.VERTICAL, command=self.rec_tree.yview
        )
        self.rec_tree.configure(yscrollcommand=self._rec_sb_set)
        self.rec_tree.grid(row=0, column=0, sticky=tk.NSEW)
        self._bind_tree_mousewheel(self.pb_tree)
        self._bind_tree_mousewheel(self.rec_tree)

        self.pb_tree.tag_configure("offline", background="#ffcccc")
        self.pb_tree.tag_configure("default", background="#fff59d", foreground="#1a1a1a")
        self.rec_tree.tag_configure("offline", background="#ffcccc")
        self.rec_tree.tag_configure("default", background="#fff59d", foreground="#1a1a1a")

        rf = ttk.Frame(main)
        rf.grid(row=r + 1, column=0, sticky=tk.EW, pady=(4, 12))
        self._btn_rec_refresh = ttk.Button(
            rf, text=self.tr("refresh"), command=self.refresh_recording_list
        )
        self._btn_rec_refresh.pack(side=tk.LEFT, padx=2)
        self._btn_rec_default = ttk.Button(
            rf, text=self.tr("set_default_recording"), command=self.apply_selected_recording
        )
        self._btn_rec_default.pack(side=tk.LEFT, padx=2)
        ttk.Button(rf, text="－", width=3, command=self._remove_group).pack(
            side=tk.RIGHT, padx=2
        )
        ttk.Button(rf, text="＋", width=3, command=self._add_group).pack(
            side=tk.RIGHT, padx=2
        )
        self._btn_lang = ttk.Button(
            rf,
            text=lang_button_label(self.cfg, get_current_lang(self.cfg)),
            command=self._cycle_language,
        )
        self._btn_lang.pack(side=tk.RIGHT, padx=6)
        self._btn_sound_console = ttk.Button(
            rf,
            text=self.tr("open_sound_console"),
            command=self._open_sound_console,
        )
        self._btn_sound_console.pack(side=tk.RIGHT, padx=6)
        self._btn_sys_sound = ttk.Button(
            rf,
            text=self.tr("open_system_sound"),
            command=self._open_system_sound,
        )
        self._btn_sys_sound.pack(side=tk.RIGHT, padx=6)
        r += 2

        main.grid_rowconfigure(r, weight=0)
        self._groups_canvas_row = r
        _canvas_bg = style.lookup("TFrame", "background") or self.root.cget("background")
        self._groups_canvas = tk.Canvas(
            main,
            highlightthickness=0,
            bd=0,
            background=_canvas_bg,
            height=GROUPS_VIEWPORT_HEIGHT,
        )
        self._groups_sb = ttk.Scrollbar(
            main, orient=tk.VERTICAL, command=self._groups_canvas.yview
        )
        self._groups_canvas.configure(yscrollcommand=self._groups_yscroll_set)
        self._groups_inner = ttk.Frame(self._groups_canvas)
        self._groups_canvas_window_id = self._groups_canvas.create_window(
            (0, 0), window=self._groups_inner, anchor=tk.NW
        )
        self._groups_inner.bind("<Configure>", self._on_groups_inner_configure)
        self._groups_canvas.bind("<Configure>", self._on_groups_canvas_resize)
        self._groups_canvas.grid(row=r, column=0, sticky=tk.EW)
        self._groups_sb.grid(row=r, column=1, sticky=tk.NS)

        self.groups_outer = ttk.LabelFrame(
            self._groups_inner, text=self.tr("group_settings_frame"), padding=6
        )
        self.groups_outer.pack(fill=tk.X, expand=False)
        self.groups_table = ttk.Frame(self.groups_outer)
        self.groups_table.pack(fill=tk.X, expand=False)

        self._bind_wheel_subtree(self._groups_inner)

        self.group_play_vars: list[tk.StringVar] = []
        self.group_rec_vars: list[tk.StringVar] = []
        self.group_play_combo: list[ttk.Combobox] = []
        self.group_rec_combo: list[ttk.Combobox] = []

    def _launch_mmsys(self, tab: int) -> None:
        if sys.platform != "win32":
            return
        arg = f"mmsys.cpl,,{tab}"
        cmds: list[list[str]] = [
            ["rundll32.exe", "shell32.dll,Control_RunDLL", arg],
            ["control.exe", arg],
        ]
        if tab == 0:
            cmds.append(["rundll32.exe", "shell32.dll,Control_RunDLL", "mmsys.cpl"])
        last_err: OSError | None = None
        for argv in cmds:
            try:
                subprocess.Popen(argv, close_fds=True)
                return
            except OSError as e:
                last_err = e
        if last_err is not None:
            return

    def _open_sound_console(self) -> None:
        self._launch_mmsys(0)

    def _open_system_sound(self) -> None:
        if sys.platform != "win32":
            return
        if sys.getwindowsversion().major >= 10:
            try:
                os.startfile("ms-settings:sound")
                return
            except OSError:
                pass
        self._launch_mmsys(2)

    def _cycle_language(self) -> None:
        global _singleton_window_title
        langs = available_ui_languages(self.cfg)
        if not langs:
            langs = ["zh_TW"]
        try:
            i = langs.index(get_current_lang(self.cfg))
        except ValueError:
            set_current_lang(self.cfg, "zh_TW" if "zh_TW" in langs else langs[0])
            i = 0
        set_current_lang(self.cfg, langs[(i + 1) % len(langs)])
        save_config(self.cfg)
        _singleton_window_title = self._win_title()
        self._refresh_i18n()

    def _refresh_i18n(self) -> None:
        global _singleton_window_title
        self.root.title(self._win_title())
        _singleton_window_title = self._win_title()
        self._set_app_user_model_id()
        self._set_app_user_model_id()
        self._lbl_pb_title.configure(text=self.tr("playback_devices"))
        self._btn_pb_refresh.configure(text=self.tr("refresh"))
        self._btn_pb_default.configure(text=self.tr("set_default_playback"))
        self._lbl_rec_title.configure(text=self.tr("recording_devices"))
        self._btn_rec_refresh.configure(text=self.tr("refresh"))
        self._btn_rec_default.configure(text=self.tr("set_default_recording"))
        self.groups_outer.configure(text=self.tr("group_settings_frame"))
        self._btn_lang.configure(text=lang_button_label(self.cfg, get_current_lang(self.cfg)))
        self._btn_sound_console.configure(text=self.tr("open_sound_console"))
        self._btn_sys_sound.configure(text=self.tr("open_system_sound"))
        self.refresh_playback_list()
        self.refresh_recording_list()
        self._rebuild_groups_ui()
        self._load_group_fields_from_config()
        self._apply_tray_language()

    def _tray_build_menu(self):
        import pystray

        return pystray.Menu(
            pystray.MenuItem(
                " ",
                self._tray_restore_at_100,
                default=True,
                visible=False,
            ),
            pystray.MenuItem(self.tr("about"), self._tray_about),
            pystray.MenuItem(self.tr("exit"), self._tray_exit),
        )

    def _apply_tray_language(self) -> None:
        if self._tray_icon is None:
            return
        try:
            self._tray_icon.menu = self._tray_build_menu()
            if hasattr(self._tray_icon, "update_menu"):
                self._tray_icon.update_menu()
            self._tray_icon.title = self._win_title()
        except Exception:
            pass

    def _groups_yscroll_set(self, first: str, last: str) -> None:
        self._groups_sb.set(first, last)
        self.root.after_idle(self._update_groups_sb_visibility)

    def _on_groups_inner_configure(self, event) -> None:
        self._groups_canvas.configure(scrollregion=self._groups_canvas.bbox("all"))
        self.root.after_idle(self._update_groups_sb_visibility)

    def _on_groups_canvas_resize(self, event) -> None:
        try:
            self._groups_canvas.itemconfigure(
                self._groups_canvas_window_id, width=event.width
            )
        except tk.TclError:
            pass
        self.root.after_idle(self._update_groups_sb_visibility)

    def _update_groups_sb_visibility(self) -> None:
        try:
            self._groups_canvas.update_idletasks()
            cvh = self._groups_canvas.winfo_height()
            ch = self._groups_inner.winfo_reqheight()
            if cvh > 1 and ch > cvh + 2:
                self._groups_sb.grid(
                    row=self._groups_canvas_row, column=1, sticky=tk.NS
                )
            else:
                self._groups_sb.grid_remove()
        except tk.TclError:
            pass

    def _bind_tree_mousewheel(self, tree: ttk.Treeview) -> None:
        def on_wheel(event: tk.Event) -> str | None:
            tree.yview_scroll(int(-event.delta / 120), "units")
            return "break"

        tree.bind("<MouseWheel>", on_wheel)

    def _bind_wheel_subtree(self, root_w: tk.Misc) -> None:
        def on_wheel(event: tk.Event) -> str | None:
            self._groups_canvas.yview_scroll(int(-event.delta / 120), "units")
            return "break"

        def walk(w: tk.Misc) -> None:
            w.bind("<MouseWheel>", on_wheel)
            for c in w.winfo_children():
                walk(c)

        walk(root_w)

    def _pb_sb_set(self, a, b) -> None:
        self.pb_sb.set(a, b)
        self._update_sb_visibility(self.pb_tree, self.pb_sb, self._pb_frame)

    def _rec_sb_set(self, a, b) -> None:
        self.rec_sb.set(a, b)
        self._update_sb_visibility(self.rec_tree, self.rec_sb, self._rec_frame)

    def _update_sb_visibility(self, tree: ttk.Treeview, sb: ttk.Scrollbar, parent: ttk.Frame) -> None:
        n = len(tree.get_children())
        need = n > DEVICE_LIST_VISIBLE_ROWS
        if need:
            sb.grid(row=0, column=1, sticky=tk.NS)
            parent.grid_columnconfigure(0, weight=1)
        else:
            sb.grid_remove()

    def _collect_offline_from_cfg(self, flow: str) -> list[tuple[str, str]]:
        out: list[tuple[str, str]] = []
        seen: set[str] = set()
        key_id = "playback_id" if flow == "play" else "recording_id"
        key_name = "playback_name" if flow == "play" else "recording_name"
        for g in self.cfg["groups"]:
            did = g.get(key_id)
            if not did or did in seen:
                continue
            seen.add(did)
            nm = (g.get(key_name) or "").strip() or did[:20] + "…"
            out.append((f"{self.tr('offline_prefix')}{nm}", did))
        return out

    def _merge_device_tree(
        self,
        devices: list,
        offline_entries: list[tuple[str, str]],
        online_ids: set[str],
    ) -> list[tuple[str, str, bool]]:
        rows: list[tuple[str, str, bool]] = []
        seen: set[str] = set()
        for d in devices:
            name = d.FriendlyName or d.id
            base = name
            if name in seen:
                name = f"{base} ({d.id[:8]}…)"
            seen.add(name)
            rows.append((name, d.id, False))
        for label, did in offline_entries:
            if did in online_ids:
                continue
            if did in seen:
                continue
            seen.add(did)
            rows.append((label, did, True))
        return rows

    def refresh_playback_list(self) -> None:
        self.pb_tree.delete(*self.pb_tree.get_children())
        self._pb_iid_meta.clear()
        self._device_name_to_id_play.clear()
        try:
            devices = list_playback_devices()
        except Exception:
            return
        online_ids = {d.id for d in devices}
        offline = self._collect_offline_from_cfg("play")
        rows = self._merge_device_tree(devices, offline, online_ids)
        default_ids = _default_ids_for_playback()
        for disp, did, off in rows:
            self._device_name_to_id_play[disp] = did
            iid = _safe_tree_iid(did)
            self._pb_iid_meta[iid] = (did, off)
            if _device_row_is_default(did, default_ids):
                tags = ("default",)
            elif off:
                tags = ("offline",)
            else:
                tags = ()
            self.pb_tree.insert("", tk.END, iid=iid, text=disp, tags=tags)
        self._sync_combobox_choices()
        self._update_sb_visibility(self.pb_tree, self.pb_sb, self._pb_frame)
        if self.group_play_combo:
            self._paint_group_combos_offline()

    def refresh_recording_list(self) -> None:
        self.rec_tree.delete(*self.rec_tree.get_children())
        self._rec_iid_meta.clear()
        self._device_name_to_id_rec.clear()
        try:
            devices = list_recording_devices()
        except Exception:
            return
        online_ids = {d.id for d in devices}
        offline = self._collect_offline_from_cfg("rec")
        rows = self._merge_device_tree(devices, offline, online_ids)
        default_ids = _default_ids_for_recording()
        for disp, did, off in rows:
            self._device_name_to_id_rec[disp] = did
            iid = _safe_tree_iid(did)
            self._rec_iid_meta[iid] = (did, off)
            if _device_row_is_default(did, default_ids):
                tags = ("default",)
            elif off:
                tags = ("offline",)
            else:
                tags = ()
            self.rec_tree.insert("", tk.END, iid=iid, text=disp, tags=tags)
        self._sync_combobox_choices()
        self._update_sb_visibility(self.rec_tree, self.rec_sb, self._rec_frame)
        if self.group_play_combo:
            self._paint_group_combos_offline()

    def _sync_combobox_choices(self) -> None:
        ce = self.tr("combo_empty")
        play_names = [ce] + list(self._device_name_to_id_play.keys())
        rec_names = [ce] + list(self._device_name_to_id_rec.keys())
        for i in range(len(self.group_play_combo)):
            self.group_play_combo[i]["values"] = play_names
            self.group_rec_combo[i]["values"] = rec_names

    def _rebuild_groups_ui(self) -> None:
        for w in self.groups_table.winfo_children():
            w.destroy()
        self.group_play_vars.clear()
        self.group_rec_vars.clear()
        self.group_play_combo.clear()
        self.group_rec_combo.clear()
        self.group_hotkey_entries.clear()

        ttk.Label(self.groups_table, text=self.tr("col_group")).grid(
            row=0, column=0, padx=4, pady=2
        )
        ttk.Label(self.groups_table, text=self.tr("col_playback")).grid(
            row=0, column=1, padx=4, pady=2
        )
        ttk.Label(self.groups_table, text=self.tr("col_recording")).grid(
            row=0, column=2, padx=4, pady=2
        )
        ttk.Label(self.groups_table, text=self.tr("col_hotkey")).grid(
            row=0, column=3, padx=4, pady=2
        )
        ttk.Label(self.groups_table, text=self.tr("col_apply")).grid(
            row=0, column=4, padx=4, pady=2
        )
        ttk.Label(self.groups_table, text=self.tr("col_save")).grid(
            row=0, column=5, padx=4, pady=2
        )
        ttk.Label(self.groups_table, text=self.tr("reset_group")).grid(
            row=0, column=6, padx=4, pady=2
        )
        self.groups_table.columnconfigure(1, weight=1)
        self.groups_table.columnconfigure(2, weight=1)

        for i in range(self.num_groups):
            r = i + 1
            ttk.Label(self.groups_table, text=str(i + 1)).grid(
                row=r, column=0, padx=4, pady=3, sticky=tk.W
            )
            pv = tk.StringVar(value=self.tr("combo_empty"))
            rv = tk.StringVar(value=self.tr("combo_empty"))
            pc = ttk.Combobox(
                self.groups_table, textvariable=pv, width=28, state="readonly"
            )
            rc = ttk.Combobox(
                self.groups_table, textvariable=rv, width=28, state="readonly"
            )
            pc.grid(row=r, column=1, padx=4, pady=3, sticky=tk.EW)
            rc.grid(row=r, column=2, padx=4, pady=3, sticky=tk.EW)
            he = ttk.Entry(self.groups_table, width=18, state="readonly", cursor="hand2")
            he.grid(row=r, column=3, padx=4, pady=3, sticky=tk.W)
            he.bind("<Button-1>", lambda e, idx=i: self._begin_hotkey_capture(idx))
            he.bind("<Button-3>", lambda e, idx=i: self._clear_group_hotkey(idx))
            ttk.Button(
                self.groups_table,
                text=self.tr("apply_this_group"),
                command=lambda idx=i: self.apply_group(idx),
            ).grid(row=r, column=4, padx=4, pady=3)
            ttk.Button(
                self.groups_table,
                text=self.tr("save_group"),
                command=lambda idx=i: self.save_group(idx),
            ).grid(row=r, column=5, padx=4, pady=3)
            ttk.Button(
                self.groups_table,
                text=self.tr("reset_group_button"),
                command=lambda idx=i: self.reset_group(idx),
            ).grid(row=r, column=6, padx=4, pady=3)
            self.group_play_vars.append(pv)
            self.group_rec_vars.append(rv)
            self.group_play_combo.append(pc)
            self.group_rec_combo.append(rc)
            self.group_hotkey_entries.append(he)
        self._sync_combobox_choices()
        self._paint_group_combos_offline()
        for pv, rv in zip(self.group_play_vars, self.group_rec_vars):
            pv.trace_add("write", lambda *_: self._schedule_paint_group_combos())
            rv.trace_add("write", lambda *_: self._schedule_paint_group_combos())
        self._bind_wheel_subtree(self.groups_table)
        self._bind_space_block_recursive(self.groups_table)
        self._disable_button_focus_rect(self.groups_table)
        self.root.after_idle(self._update_groups_sb_visibility)

    def _schedule_paint_group_combos(self) -> None:
        self.root.after_idle(self._paint_group_combos_offline)

    def _paint_group_combos_offline(self) -> None:
        for pv, pc in zip(self.group_play_vars, self.group_play_combo):
            try:
                pc.configure(
                    style="RedRO.TCombobox"
                    if pv.get().startswith(self.tr("offline_prefix"))
                    else "TCombobox"
                )
            except tk.TclError:
                pass
        for rv, rc in zip(self.group_rec_vars, self.group_rec_combo):
            try:
                rc.configure(
                    style="RedRO.TCombobox"
                    if rv.get().startswith(self.tr("offline_prefix"))
                    else "TCombobox"
                )
            except tk.TclError:
                pass

    def _add_group(self) -> None:
        if self.num_groups >= MAX_GROUPS:
            return
        self.num_groups += 1
        self.cfg["num_groups"] = self.num_groups
        save_config(self.cfg)
        self._rebuild_groups_ui()
        self._load_group_fields_from_config()
        self._try_apply_hotkeys_from_cfg()

    def _remove_group(self) -> None:
        if self.num_groups <= MIN_GROUPS:
            return
        self.num_groups -= 1
        self.cfg["num_groups"] = self.num_groups
        save_config(self.cfg)
        self._rebuild_groups_ui()
        self._load_group_fields_from_config()
        self._try_apply_hotkeys_from_cfg()

    def _display_for_saved_id(
        self, name_map: dict[str, str], dev_id: str | None, saved_name: str
    ) -> str:
        if not dev_id:
            return self.tr("combo_empty")
        for label, did in name_map.items():
            if did == dev_id:
                return label
        pfx = self.tr("offline_prefix")
        if saved_name:
            return f"{pfx}{saved_name}"
        return f"{pfx}{dev_id[:24]}…"

    def _load_group_fields_from_config(self) -> None:
        for i in range(len(self.group_play_vars)):
            g = self.cfg["groups"][i]
            pid = g.get("playback_id")
            rid = g.get("recording_id")
            pn = (g.get("playback_name") or "").strip()
            rn = (g.get("recording_name") or "").strip()
            self.group_play_vars[i].set(
                self._display_for_saved_id(self._device_name_to_id_play, pid, pn)
            )
            self.group_rec_vars[i].set(
                self._display_for_saved_id(self._device_name_to_id_rec, rid, rn)
            )
        for hi in range(len(self.group_hotkey_entries)):
            self._set_hotkey_entry_text(hi)
        self._paint_group_combos_offline()

    def _resolve_combo_id_with_offline(
        self, name_map: dict[str, str], var: tk.StringVar
    ) -> tuple[str | None, str]:
        label = var.get().strip()
        ce = self.tr("combo_empty")
        if not label or label == ce:
            return None, ""
        did = name_map.get(label)
        if not did:
            return None, ""
        pfx = self.tr("offline_prefix")
        if label.startswith(pfx):
            return did, label.replace(pfx, "", 1).strip()
        return did, label

    def save_group(self, index: int) -> None:
        pid, pn = self._resolve_combo_id_with_offline(
            self._device_name_to_id_play, self.group_play_vars[index]
        )
        rid, rn = self._resolve_combo_id_with_offline(
            self._device_name_to_id_rec, self.group_rec_vars[index]
        )
        prev = self.cfg["groups"][index]
        self.cfg["groups"][index] = {
            "playback_id": pid,
            "recording_id": rid,
            "playback_name": pn,
            "recording_name": rn,
            "hotkey_mods": prev.get("hotkey_mods"),
            "hotkey_vk": prev.get("hotkey_vk"),
        }
        save_config(self.cfg)
        self.refresh_playback_list()
        self.refresh_recording_list()
        self._load_group_fields_from_config()
        play_success_sound()

    def reset_group(self, index: int) -> None:
        if index < 0 or index >= MAX_GROUPS:
            return
        self.cfg["groups"][index] = _empty_group()
        save_config(self.cfg)
        self.refresh_playback_list()
        self.refresh_recording_list()
        self._load_group_fields_from_config()
        self._try_apply_hotkeys_from_cfg()
        play_success_sound()

    def play_success_sound_after_device_change(self) -> None:
        self.root.after(150, play_success_sound)

    def apply_selected_playback(self) -> None:
        sel = self.pb_tree.selection()
        if not sel:
            return
        iid = sel[0]
        meta = self._pb_iid_meta.get(iid)
        if not meta:
            return
        did, off = meta
        if off:
            return
        try:
            set_default_playback(did)
            self.refresh_playback_list()
            self.play_success_sound_after_device_change()
        except Exception:
            return

    def apply_selected_recording(self) -> None:
        sel = self.rec_tree.selection()
        if not sel:
            return
        iid = sel[0]
        meta = self._rec_iid_meta.get(iid)
        if not meta:
            return
        did, off = meta
        if off:
            return
        try:
            set_default_recording(did)
            self.refresh_recording_list()
            self.play_success_sound_after_device_change()
        except Exception:
            return

    def apply_group(self, index: int) -> None:
        g = self.cfg["groups"][index]
        pid = g.get("playback_id")
        rid = g.get("recording_id")
        play_success_sound()
        if not pid and not rid:
            return
        try:
            if pid:
                set_default_playback(pid)
            if rid:
                set_default_recording(rid)
            self.refresh_playback_list()
            self.refresh_recording_list()
            play_success_sound()
        except Exception:
            return

    def apply_group_from_hotkey(self, index: int) -> None:
        if index >= self.num_groups:
            return
        g = self.cfg["groups"][index]
        pid = g.get("playback_id")
        rid = g.get("recording_id")
        if not pid and not rid:
            return
        try:
            if pid:
                set_default_playback(pid)
            if rid:
                set_default_recording(rid)
            self.refresh_playback_list()
            self.refresh_recording_list()
            play_success_sound()
        except Exception:
            pass

    def _set_app_user_model_id(self) -> None:
        if sys.platform != "win32":
            return
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                self._win_title()
            )
        except Exception:
            pass

    def _set_app_user_model_id(self) -> None:
        if sys.platform != "win32":
            return
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                self._win_title()
            )
        except Exception:
            pass

    def _set_hotkey_entry_text(self, index: int) -> None:
        if index >= len(self.group_hotkey_entries):
            return
        e = self.group_hotkey_entries[index]
        pair = effective_group_hotkey(self.cfg["groups"][index], index)
        text = hotkey_display_text(pair[0], pair[1]) if pair else ""
        e.configure(state="normal")
        e.delete(0, tk.END)
        if text:
            e.insert(0, text)
        e.configure(state="readonly")

    def _find_hotkey_conflict(
        self, skip_index: int, mods: int, vk: int
    ) -> int | None:
        for j in range(self.num_groups):
            if j == skip_index:
                continue
            p = effective_group_hotkey(self.cfg["groups"][j], j)
            if p and p[0] == mods and p[1] == vk:
                return j
        return None

    def _cleanup_hotkeys(self) -> None:
        try:
            if keyboard is not None:
                keyboard.clear_all_hotkeys()
        except Exception:
            pass
        self._hotkey_hwnd = None

    def _register_all_hotkeys(self) -> bool:
        if keyboard is None:
            return False
        try:
            keyboard.clear_all_hotkeys()
        except Exception:
            pass

        any_registered = False
        for i in range(self.num_groups):
            pair = effective_group_hotkey(self.cfg["groups"][i], i)
            if not pair:
                continue
            mods, vk = pair
            combos = mods_vk_to_keyboard_combos(mods, vk)
            if not combos:
                continue
            try:
                                                                          
                def _cb(idx=i) -> None:
                    try:
                        self.root.after(0, lambda: self.apply_group_from_hotkey(idx))
                    except Exception:
                        pass
                for combo in combos:
                    try:
                        keyboard.add_hotkey(combo, _cb, suppress=False)
                        any_registered = True
                    except Exception:
                        pass
            except Exception:
                                                               
                pass
        return any_registered

    def _try_apply_hotkeys_from_cfg(self) -> bool:
        self._cleanup_hotkeys()
        return self._register_all_hotkeys()

    def _setup_hotkeys(self) -> None:
        self._try_apply_hotkeys_from_cfg()

    def _commit_group_hotkey(
        self, index: int, mods: int, vk: int, parent: tk.Misc | None
    ) -> bool:
        cof = self._find_hotkey_conflict(index, mods, vk)
        if cof is not None:
            return False
        g = self.cfg["groups"][index]
        prev_m, prev_v = g.get("hotkey_mods"), g.get("hotkey_vk")
        g["hotkey_mods"] = mods
        g["hotkey_vk"] = vk
        save_config(self.cfg)
        if not self._try_apply_hotkeys_from_cfg():
            g["hotkey_mods"], g["hotkey_vk"] = prev_m, prev_v
            save_config(self.cfg)
            self._try_apply_hotkeys_from_cfg()
            return False
        self._set_hotkey_entry_text(index)
        return True

    def _detach_hotkey_capture_silent(self) -> None:
        idx = self._hotkey_capture_index
        if idx is not None:
            try:
                self.group_hotkey_entries[idx].unbind("<KeyPress>")
            except (tk.TclError, IndexError):
                pass
        if self._hotkey_clickout_bind_id is not None:
            try:
                self.root.unbind("<Button-1>", self._hotkey_clickout_bind_id)
            except tk.TclError:
                pass
            self._hotkey_clickout_bind_id = None
        self._hotkey_capture_index = None

    def _cancel_hotkey_capture(self) -> None:
        if self._hotkey_capture_index is None:
            return
        idx = self._hotkey_capture_index
        self._detach_hotkey_capture_silent()
        self._set_hotkey_entry_text(idx)

    def _hotkey_root_click(self, event: tk.Event) -> None:
        if self._hotkey_capture_index is None:
            return
        w = event.widget
        for he in self.group_hotkey_entries:
            if w == he:
                return
        self._cancel_hotkey_capture()

    def _guard_space_key(self, event: tk.Event) -> str | None:
        w = self.root.focus_get()
        if w is None:
            return None
        cls = str(getattr(w, "winfo_class", lambda: "")())
        if self._hotkey_capture_index is not None:
            return "break"
        if w in self.group_hotkey_entries:
            return "break"
        if w in (self._btn_lang, self._btn_sound_console, self._btn_sys_sound):
            return "break"
                                                                                                      
        if cls in ("TButton", "Button"):
            return "break"
        if cls in ("TCheckbutton", "Checkbutton"):
            return "break"
        if str(getattr(w, "winfo_parent", lambda: "")()) == str(self.groups_table):
            if cls in ("TButton", "TCombobox", "TEntry"):
                return "break"
        return None

    def _bind_space_block_recursive(self, root_widget) -> None:
        cls = str(getattr(root_widget, "winfo_class", lambda: "")())
        if cls in ("TButton", "Button", "TCheckbutton", "Checkbutton"):
            root_widget.bind("<KeyPress-space>", lambda _e: "break", add="+")
            root_widget.bind("<KeyRelease-space>", lambda _e: "break", add="+")
        for ch in getattr(root_widget, "winfo_children", lambda: [])():
            self._bind_space_block_recursive(ch)

    def _disable_button_focus_rect(self, root_widget) -> None:
        cls = str(getattr(root_widget, "winfo_class", lambda: "")())
        if cls in ("TButton", "Button", "TCheckbutton", "Checkbutton"):
            try:
                root_widget.configure(takefocus=False)
            except tk.TclError:
                pass
        for ch in getattr(root_widget, "winfo_children", lambda: [])():
            self._disable_button_focus_rect(ch)

    def _defocus_toplevel_of(self, w) -> None:
        try:
            t = w.winfo_toplevel()
            if t and t.winfo_exists():
                t.focus_set()
        except Exception:
            pass

    def _on_ttk_buttonlike_release(self, event) -> None:
        self.root.after_idle(lambda: self._defocus_toplevel_of(event.widget))

    def _install_ttk_focus_ring_mitigation(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.configure("TButton", focuscolor=style.lookup("TButton", "background"))
        except tk.TclError:
            pass
        try:
            bg = style.lookup("TCheckbutton", "background")
            style.configure("TCheckbutton", focuscolor=bg)
            style.map("TCheckbutton", focuscolor=[("focus", bg), ("!focus", bg)])
        except tk.TclError:
            pass
        if getattr(self, "_ttk_focus_release_bind_done", False):
            return
        self._ttk_focus_release_bind_done = True
        self.root.bind_class("TButton", "<ButtonRelease-1>", self._on_ttk_buttonlike_release, add="+")
        self.root.bind_class("TCheckbutton", "<ButtonRelease-1>", self._on_ttk_buttonlike_release, add="+")

    def _begin_hotkey_capture(self, index: int) -> None:
        if self._hotkey_capture_index == index:
            return
        if self._hotkey_capture_index is not None:
            self._cancel_hotkey_capture()
        self._hotkey_capture_index = index
        e = self.group_hotkey_entries[index]
        e.configure(state="normal")
        e.delete(0, tk.END)
        e.insert(0, self.tr("hotkey_listen_inline"))
        e.focus_set()
        e.icursor(tk.END)

        def on_key(event: tk.Event) -> str:
            return self._on_hotkey_entry_key(index, event)

        e.bind("<KeyPress>", on_key)
        self._hotkey_clickout_bind_id = self.root.bind(
            "<Button-1>", self._hotkey_root_click, add="+"
        )

    def _on_hotkey_entry_key(self, index: int, event: tk.Event) -> str:
        if self._hotkey_capture_index != index:
            return ""
        if event.keysym == "Escape":
            self._cancel_hotkey_capture()
            return "break"
        parsed = tk_keyevent_to_mods_vk(event)
        if parsed is None:
            if event.keysym not in (
                "Shift_L",
                "Shift_R",
                "Control_L",
                "Control_R",
                "Alt_L",
                "Alt_R",
                "Win_L",
                "Win_R",
            ):
                return "break"
            return "break"
        mods, vk = parsed
        if self._commit_group_hotkey(index, mods, vk, parent=self.root):
            self._detach_hotkey_capture_silent()
        return "break"

    def _clear_group_hotkey(self, index: int) -> None:
        if self._hotkey_capture_index is not None:
            if self._hotkey_capture_index == index:
                self._detach_hotkey_capture_silent()
            else:
                self._cancel_hotkey_capture()
        self.cfg["groups"][index]["hotkey_mods"] = 0
        self.cfg["groups"][index]["hotkey_vk"] = 0
        save_config(self.cfg)
        self._set_hotkey_entry_text(index)
        self._try_apply_hotkeys_from_cfg()

    def _tray_restore_at_100(self, icon=None, item=None) -> None:
        def go() -> None:
            if not self.root.winfo_exists():
                return
            self._hiding_to_tray = False
            self.root.state("normal")
            self.root.deiconify()
            self._move_to_default_position()
            self.root.lift()
            try:
                self.root.focus_force()
            except tk.TclError:
                pass

        self.root.after(0, go)

    def _tray_about(self, icon=None, item=None) -> None:
        try:
            webbrowser.open(ABOUT_URL)
        except Exception:
            pass

    def _tray_exit(self, icon=None, item=None) -> None:
        self.root.after(0, self._on_close)

    def _start_tray(self) -> None:
        try:
            import pystray
            from PIL import Image
        except ImportError:
            return
        ico_path = resource_path("icons", "icon.ico")
        if not os.path.isfile(ico_path):
            return
        try:
            img = Image.open(ico_path).convert("RGBA")
        except OSError:
            return
        try:
            self._tray_icon = pystray.Icon(
                "bm_sound_effects_switch",
                img,
                self._win_title(),
                self._tray_build_menu(),
            )
        except Exception:
            return

        def run_tray() -> None:
            try:
                self._tray_icon.run()
            except Exception:
                pass

        threading.Thread(target=run_tray, daemon=True).start()

    def _on_window_unmap(self, event=None) -> None:
        if self._tray_icon is None or not self.root.winfo_exists():
            return
        if self._hiding_to_tray:
            return
        try:
            is_minimized = self.root.state() == "iconic"
        except tk.TclError:
            return
        if not is_minimized:
            return
        self._hiding_to_tray = True

        def hide_to_tray() -> None:
            if not self.root.winfo_exists():
                return
            try:
                self.root.withdraw()
            except tk.TclError:
                pass
            self._hiding_to_tray = False

        self.root.after(0, hide_to_tray)

    def _on_close(self) -> None:
        if self._hotkey_capture_index is not None:
            self._detach_hotkey_capture_silent()
        self._cleanup_hotkeys()
        if self._tray_icon is not None:
            try:
                self._tray_icon.stop()
            except Exception:
                pass
            self._tray_icon = None
        _release_app_mutex()
        try:
            self.root.destroy()
        except tk.TclError:
            pass


def main() -> None:
    global _singleton_window_title
    if sys.platform != "win32":
        print("僅支援 Windows。")
        sys.exit(1)
    _pre_cfg = load_config()
                                                                            
    try:
        if not os.path.isfile(config_path()):
            save_config(_pre_cfg)
    except Exception:
        pass
    _singleton_window_title = window_title_full(_pre_cfg)
    global _app_mutex_handle
    _app_mutex_handle = bm_single_instance.acquire_or_handshake(SINGLE_APP_ID)
    if not _app_mutex_handle:
        sys.exit(1)
    atexit.register(_release_app_mutex)
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = SoundSwitcherApp()
    bm_single_instance.start_pipe_server(SINGLE_APP_ID, lambda: app.root.after(0, app._on_close))
    app.root.mainloop()


if __name__ == "__main__":
    main()
