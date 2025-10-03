import os
import re
import sys
import time
import configparser
from datetime import datetime
import ctypes
from ctypes import wintypes

# Constants
VERSION = "0.5"
VALID_RIPS = ["RIP1", "RIP2", "RIP3"]
DEFAULT_POLLING_INTERVAL = 5.0
RETRY_MAX_ATTEMPTS = 3
RETRY_DELAY_SECONDS = 1
LOG_DATETIME_FORMAT = "%Y%m%d_%H%M%S"
DETAILED_DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"

try:
    import win32file
    import win32con
except Exception:
    win32file = None
    win32con = None

def is_valid_tiff(filename):
    # bip<0-5>-output-1bpp-<繝壹・繧ｸ逡ｪ蜿ｷ>.tif 縺ｫ繝槭ャ繝√☆繧九°
    pattern = r"^bip([0-5])-output-1bpp-([1-9][0-9]*)\.tif$"
    return re.match(pattern, filename, re.IGNORECASE)

def is_file_locked(filepath):
    try:
        handle = win32file.CreateFile(
            filepath,
            win32con.GENERIC_READ,
            0,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None
        )
        handle.close()
        return False
    except win32file.error as e:
        print(f"File access error: {e}")
        return True
    except Exception as e:
        print(f"Unexpected error: {e}")
        return True

def is_file_complete(filepath):
    """Check if file is completely written"""
    try:
        # Check file size
        if os.path.getsize(filepath) == 0:
            return False
        
        # Check if file is readable
        with open(filepath, 'rb') as f:
            f.read(1)
        return True
    except OSError as e:
        print(f"File access error: {e}")
        return False
    except Exception as e:
        print(f"Unexpected error: {e}")
        return False

def ensure_log_directory(log_dir):
    """Ensure log directory exists; exit if it cannot be created or is not configured."""
    if not log_dir:
        print("Log directory not configured; logging is required. Exiting.")
        sys.exit(1)
    try:
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        return True
    except Exception as e:
        print(f"Failed to create/access log directory '{log_dir}': {e}")
        sys.exit(1)

def is_file_ready_for_deletion(filepath):
    """Check if file is ready for deletion with minimal I/O"""
    try:
        # 1蝗槭・繝輔ぃ繧､繝ｫ繧ｪ繝ｼ繝励Φ縺ｧ蠢・ｦ√↑諠・ｱ繧貞叙蠕・        stats = os.stat(filepath)
        
        # 繧ｵ繧､繧ｺ繝√ぉ繝・け
        if stats.st_size == 0:
            return False
            
        # 譛邨よ峩譁ｰ譎ょ綾繝√ぉ繝・け・・0遘剃ｻ･荳顔ｵ碁℃縺励◆繝輔ぃ繧､繝ｫ縺ｮ縺ｿ蟇ｾ雎｡・・        if time.time() - stats.st_mtime < 30:
            return False
            
        return True
    except Exception as e:
        print(f"File check error: {e}")
        return False

def delete_matching_files(rip_name, path, log_dir):
    if not ensure_log_directory(log_dir):
        print(f"[{rip_name}] Log directory error. Skipping operation.")
        return

    now = datetime.now().strftime(LOG_DATETIME_FORMAT)
    deleted_files = []
    skipped_files = []

    for filename in os.listdir(path):
        if is_valid_tiff(filename):
            full_path = os.path.join(path, filename)
            
            try:
                # 繝輔ぃ繧､繝ｫ縺ｮ蟄伜惠縺ｨ繧ｵ繧､繧ｺ繧剃ｸ蠎ｦ縺縺代メ繧ｧ繝・け
                if os.path.getsize(full_path) == 0:
                    print(f"[{rip_name}] Skipped (Empty file): {filename}")
                    skipped_files.append((filename, "Empty file"))
                    continue

                # 蜑企勁繧定ｩｦ陦・                if delete_with_retry(full_path, RETRY_MAX_ATTEMPTS, RETRY_DELAY_SECONDS):
                    deleted_files.append(filename)
                    print(f"[{rip_name}] Deleted: {filename}")
                else:
                    print(f"[{rip_name}] Delete failed (after retry): {filename}")
                    skipped_files.append((filename, "Delete failed"))

            except PermissionError:
                print(f"[{rip_name}] Skipped (In use): {filename}")
                skipped_files.append((filename, "In use"))
            except Exception as e:
                print(f"[{rip_name}] Error: {filename} 竊・{e}")
                skipped_files.append((filename, f"Error: {e}"))

    if deleted_files or skipped_files:  # 蜑企勁縺ｾ縺溘・繧ｹ繧ｭ繝・・縺励◆繝輔ぃ繧､繝ｫ縺後≠繧句ｴ蜷・        log_filename = f"{now}_{rip_name}.log"
        log_path = os.path.join(log_dir, log_filename)
        write_detailed_log(log_path, deleted_files, skipped_files)
    else:
        print(f"[{rip_name}] No files to delete.")

def write_detailed_log(log_path, deleted_files, skipped_files):
    """Write detailed log; exit if writing fails because logs are required."""
    try:
        with open(log_path, "w", encoding="utf-8") as log_file:
            log_file.write(f"Execution time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            log_file.write("\n=== Deleted Files ===\n")
            for name in deleted_files:
                log_file.write(f"{name}\n")
            log_file.write("\n=== Skipped Files ===\n")
            for name, reason in skipped_files:
                log_file.write(f"{name} (Reason: {reason})\n")
    except Exception as e:
        print(f"Failed to write log '{log_path}': {e}")
        sys.exit(1)

def get_config_path():
    """Get the config.ini path relative to the executable"""
    if getattr(sys, 'frozen', False):
        # PyInstaller縺ｧ繝薙Ν繝峨＆繧後◆蝣ｴ蜷・        config_path = os.path.join(os.path.dirname(sys.executable), "config.ini")
    else:
        # 騾壼ｸｸ縺ｮPython螳溯｡後・蝣ｴ蜷・        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
    
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuration file not found: {config_path}")
    return config_path

def load_config():
    config = configparser.ConfigParser()
    config_file = get_config_path()
    config.read(config_file, encoding="utf-8")
    validate_config(config)
    return config

def run_for_rip(config, rip_name):
    if rip_name not in config:
        print(f"[{rip_name}] Configuration not found.")
        return

    section = config[rip_name]
    if section.getboolean("enabled", fallback=False):
        path = section.get("path", "")
        if not os.path.isdir(path):
            print(f"[{rip_name}] Path does not exist: {path}")
            return

        log_dir = config["General"].get("log_dir", "")
        if log_dir:
            # 繝ｭ繧ｰ繝・ぅ繝ｬ繧ｯ繝医Μ縺悟ｭ伜惠縺吶ｋ蝣ｴ蜷医∝商縺・Ο繧ｰ繧呈ｸ・祉
            cleanup_old_logs(log_dir)

        delete_matching_files(rip_name, path, log_dir)
    else:
        print(f"[{rip_name}] Disabled.")

def run_polling_mode(config):
    interval = config["General"].getfloat("polling_interval", fallback=DEFAULT_POLLING_INTERVAL)
    print(f"Started in polling mode. Running every {interval} minutes.")
    try:
        while True:
            for rip in VALID_RIPS:
                run_for_rip(config, rip)
            time.sleep(interval * 60)
    except KeyboardInterrupt:
        print("Polling interrupted.")

def run_kick_mode(config, target):
    if target.upper() == "ALL":
        for rip in VALID_RIPS:
            run_for_rip(config, rip)
    else:
        run_for_rip(config, target)

def delete_with_retry(file_path, max_retries=3, retry_delay=1):
    """繝ｪ繝医Λ繧､讖溯・莉倥″繝輔ぃ繧､繝ｫ蜑企勁"""
    for attempt in range(max_retries):
        try:
            os.remove(file_path)
            return True
        except PermissionError:
            if attempt < max_retries - 1:
                time.sleep(retry_delay)
                continue
            # 譛邨ょ､ｱ謨玲凾縺ｯ萓句､悶ｒ蜀阪せ繝ｭ繝ｼ縺帙★ False 繧定ｿ斐☆・亥他縺ｳ蜃ｺ縺怜・縺ｧ繧ｹ繧ｭ繝・・蜃ｦ逅・☆繧具ｼ・            return False
        except FileNotFoundError:
            # 譌｢縺ｫ蜑企勁縺輔ｌ縺ｦ縺・ｋ蝣ｴ蜷医・謌仙粥謇ｱ縺・            return True
    return False

def validate_config(config):
    """Validate configuration"""
    if "General" not in config:
        raise ValueError("General section is required")
    
    required_general = ["log_dir", "polling_interval"]
    for key in required_general:
        if key not in config["General"]:
            raise ValueError(f"'{key}' is required in General section")
    
    interval = config["General"].getfloat("polling_interval")
    if interval <= 0:
        raise ValueError("polling_interval must be a positive value")
    
    for rip in VALID_RIPS:
        if rip in config and config[rip].getboolean("enabled", False):
            if "path" not in config[rip]:
                raise ValueError(f"'path' is required in {rip}")

def cleanup_old_logs(log_dir, days_to_keep=30):
    """蜿､縺・Ο繧ｰ繝輔ぃ繧､繝ｫ繧貞炎髯､"""
    if not log_dir:
        return
    if not os.path.isdir(log_dir):
        # 繝ｭ繧ｰ繝・ぅ繝ｬ繧ｯ繝医Μ縺後↑縺代ｌ縺ｰ菴輔ｂ縺励↑縺・ｼ亥､夜Κ繝ｭ繧ｰ隗｣譫舌ヤ繝ｼ繝ｫ縺ｨ縺ｮ謨ｴ蜷域ｧ繧堤ｶｭ謖・ｼ・        return
    current_time = datetime.now()
    try:
        entries = os.listdir(log_dir)
    except Exception as e:
        print(f"Failed to list log directory: {e}")
        return

    for filename in entries:
        if filename.endswith('.log'):
            file_path = os.path.join(log_dir, filename)
            try:
                file_time = datetime.fromtimestamp(os.path.getctime(file_path))
            except Exception:
                # 繝輔ぃ繧､繝ｫ繝｡繧ｿ諠・ｱ蜿門ｾ励↓螟ｱ謨励＠縺溘ｉ繧ｹ繧ｭ繝・・
                continue
            if (current_time - file_time).days > days_to_keep:
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"Failed to delete old log: {filename} 竊・{e}")

def disable_quick_edit():
    """Disable QuickEdit mode so console selection doesn't pause the process."""
    try:
        kernel32 = ctypes.windll.kernel32
        STD_INPUT_HANDLE = -10
        ENABLE_QUICK_EDIT_MODE = 0x0040
        ENABLE_EXTENDED_FLAGS = 0x0080

        hStdin = kernel32.GetStdHandle(STD_INPUT_HANDLE)
        mode = wintypes.DWORD()
        if kernel32.GetConsoleMode(hStdin, ctypes.byref(mode)):
            new_mode = mode.value
            new_mode &= ~ENABLE_QUICK_EDIT_MODE       # clear QuickEdit
            new_mode |= ENABLE_EXTENDED_FLAGS         # required to apply change reliably
            kernel32.SetConsoleMode(hStdin, new_mode)
    except Exception:
        # 髱杆indows迺ｰ蠅・ｄ螟ｱ謨玲凾縺ｯ辟｡隕厄ｼ亥ｮ牙・蛛ｴ・・        pass

def main():
    disable_quick_edit()
    if len(sys.argv) >= 2 and sys.argv[1] == "--version":
        print(f"ripCleaner version {VERSION}")
        return

    print(f"ripCleaner version {VERSION} started.")
    
    config = load_config()
    
    # 繧ｭ繝・け繝｢繝ｼ繝峨・蜃ｦ逅・    if len(sys.argv) >= 3 and sys.argv[1] == "--kick":
        run_kick_mode(config, sys.argv[2])
    # 繝昴・繝ｪ繝ｳ繧ｰ繝｢繝ｼ繝会ｼ医ョ繝輔か繝ｫ繝茨ｼ・    else:
        run_polling_mode(config)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Unexpected error occurred: {e}")
        sys.exit(1)


