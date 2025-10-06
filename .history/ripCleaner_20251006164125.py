import os
import re
import sys
import time
import configparser
from datetime import datetime
import ctypes
from ctypes import wintypes
from concurrent.futures import ThreadPoolExecutor, as_completed

# Constants
APP_NAME = "ripCleaner"          # <-- set your new app name here
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
    # bip<0-5>-output-1bpp-<ページ番号>.tif にマッチするか
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
        # 1回のファイルオープンで必要な情報を取得
        stats = os.stat(filepath)
        
        # サイズチェック
        if stats.st_size == 0:
            return False
            
        # 最終更新時刻チェック（30秒以上経過したファイルのみ対象）
        if time.time() - stats.st_mtime < 30:
            return False
            
        return True
    except Exception as e:
        print(f"File check error: {e}")
        return False

def delete_matching_files(rip_name, path, log_dir, summary_console=False):
    if not ensure_log_directory(log_dir):
        print(f"[{rip_name}] Log directory error. Skipping operation.")
        return

    now = datetime.now().strftime(LOG_DATETIME_FORMAT)
    deleted_files = []
    skipped_files = []
    console_msgs = []
    processed_count = 0

    # Protect os.listdir against access/network errors
    try:
        entries = os.listdir(path)
    except Exception as e:
        console_msgs.append(f"[{rip_name}] Failed to access path '{path}': {e}")
        skipped_files.append(("<ACCESS_ERROR>", f"Cannot access path '{path}': {e}"))
        log_filename = f"{now}_{rip_name}.log"
        log_path = os.path.join(log_dir, log_filename)
        write_detailed_log(log_path, deleted_files, skipped_files)
        # print according to mode
        if summary_console:
            print(f"[{rip_name}] Processed={processed_count} Deleted={len(deleted_files)} Skipped={len(skipped_files)}")
        else:
            for m in console_msgs:
                print(m)
        return

    for filename in entries:
        if is_valid_tiff(filename):
            full_path = os.path.join(path, filename)
            processed_count += 1

            try:
                # ファイルの存在とサイズを一度だけチェック
                try:
                    size = os.path.getsize(full_path)
                except FileNotFoundError:
                    # 列挙→取得の間に消えた場合は無視
                    continue
                except Exception as e:
                    console_msgs.append(f"[{rip_name}] Skipped (stat error): {filename} -> {e}")
                    skipped_files.append((filename, f"Stat error: {e}"))
                    continue

                if size == 0:
                    skipped_files.append((filename, "Empty file"))
                    console_msgs.append(f"[{rip_name}] Skipped (Empty file): {filename}")
                    continue

                # 削除を試行
                if delete_with_retry(full_path, RETRY_MAX_ATTEMPTS, RETRY_DELAY_SECONDS):
                    deleted_files.append(filename)
                    console_msgs.append(f"[{rip_name}] Deleted: {filename}")
                else:
                    skipped_files.append((filename, "Delete failed"))
                    console_msgs.append(f"[{rip_name}] Delete failed (after retry): {filename}")

            except PermissionError:
                skipped_files.append((filename, "In use"))
                console_msgs.append(f"[{rip_name}] Skipped (In use): {filename}")
            except Exception as e:
                skipped_files.append((filename, f"Error: {e}"))
                console_msgs.append(f"[{rip_name}] Error: {filename} → {e}")

    # コンソール出力はモードに応じて集計 or 詳細を出す
    if summary_console:
        print(f"[{rip_name}] Processed={processed_count} Deleted={len(deleted_files)} Skipped={len(skipped_files)}")
    else:
        for m in console_msgs:
            print(m)

    if deleted_files or skipped_files:  # 削除またはスキップしたファイルがある場合
        log_filename = f"{now}_{rip_name}.log"
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
        # PyInstallerでビルドされた場合
        config_path = os.path.join(os.path.dirname(sys.executable), "config.ini")
    else:
        # 通常のPython実行の場合
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
    
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
            # ログディレクトリが存在する場合、古いログを清掃
            cleanup_old_logs(log_dir)

        delete_matching_files(rip_name, path, log_dir)
    else:
        print(f"[{rip_name}] Disabled.")

def run_polling_mode(config):
    interval = config["General"].getfloat("polling_interval", fallback=DEFAULT_POLLING_INTERVAL)
    parallel_enabled = config["General"].getboolean("parallel", fallback=False)
    summary_console = config["General"].getboolean("summary_console", fallback=False)

    if parallel_enabled:
        workers = len(VALID_RIPS)
        print(f"Started in polling mode (parallel). Running every {interval} minutes. workers={workers} summary_console={summary_console}")
        try:
            while True:
                with ThreadPoolExecutor(max_workers=workers) as ex:
                    futures = {ex.submit(run_for_rip, config, rip): rip for rip in VALID_RIPS}
                    for fut in as_completed(futures):
                        rip = futures[fut]
                        try:
                            fut.result()
                        except Exception as e:
                            print(f"[{rip}] Thread error: {e}")
                time.sleep(interval * 60)
        except KeyboardInterrupt:
            print("Polling interrupted.")
    else:
        print(f"Started in polling mode (sequential). Running every {interval} minutes. summary_console={summary_console}")
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
    """リトライ機能付きファイル削除"""
    for attempt in range(max_retries):
        try:
            os.remove(file_path)
            return True
        except PermissionError:
            if attempt < max_retries - 1:
                time.sleep(retry_delay)
                continue
            # 最終失敗時は例外を再スローせず False を返す（呼び出し側でスキップ処理する）
            return False
        except FileNotFoundError:
            # 既に削除されている場合は成功扱い
            return True
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
    """古いログファイルを削除"""
    if not log_dir:
        return
    if not os.path.isdir(log_dir):
        # ログディレクトリがなければ何もしない（外部ログ解析ツールとの整合性を維持）
        return
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
                # ファイルメタ情報取得に失敗したらスキップ
                continue
            if (current_time - file_time).days > days_to_keep:
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"Failed to delete old log: {filename} → {e}")

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
        # 非Windows環境や失敗時は無視（安全側）
        pass

def main():
    disable_quick_edit()
    if len(sys.argv) >= 2 and sys.argv[1] == "--version":
        print(f"{APP_NAME} version {VERSION}")
        return

    print(f"{APP_NAME} version {VERSION} started.")
    
    config = load_config()
    
    # キックモードの処理
    if len(sys.argv) >= 3 and sys.argv[1] == "--kick":
        run_kick_mode(config, sys.argv[2])
    # ポーリングモード（デフォルト）
    else:
        run_polling_mode(config)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Unexpected error occurred: {e}")
        sys.exit(1)



