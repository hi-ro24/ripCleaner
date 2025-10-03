import os
import re
import sys
import time
import configparser
from datetime import datetime
import win32file
import win32con
import msvcrt

def is_valid_tiff(filename):
    # bip<0-5>-output-1bpp-<繝壹・繧ｸ逡ｪ蜿ｷ>.tif 縺ｫ繝槭ャ繝√☆繧九°
    pattern = r"^bip([0-5])-output-1bpp-([1-9][0-9]*)\.tif$"
    return re.match(pattern, filename, re.IGNORECASE)

def is_file_locked(filepath):
    """繝輔ぃ繧､繝ｫ縺後Ο繝・け縺輔ｌ縺ｦ縺・ｋ縺九メ繧ｧ繝・け"""
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
        print(f"繝輔ぃ繧､繝ｫ繧｢繧ｯ繧ｻ繧ｹ繧ｨ繝ｩ繝ｼ: {e}")
        return True
    except Exception as e:
        print(f"莠域悄縺帙〓繧ｨ繝ｩ繝ｼ: {e}")
        return True

def is_file_complete(filepath):
    """繝輔ぃ繧､繝ｫ縺悟ｮ悟・縺ｫ譖ｸ縺榊・縺輔ｌ縺ｦ縺・ｋ縺九メ繧ｧ繝・け"""
    try:
        # 繝輔ぃ繧､繝ｫ繧ｵ繧､繧ｺ縺・縺ｧ縺ｪ縺・％縺ｨ繧堤｢ｺ隱・        if os.path.getsize(filepath) == 0:
            return False
        
        # 繝輔ぃ繧､繝ｫ縺瑚ｪｭ縺ｿ蜿悶ｊ蜿ｯ閭ｽ縺九メ繧ｧ繝・け
        with open(filepath, 'rb') as f:
            # 譛蛻昴・謨ｰ繝舌う繝医ｒ隱ｭ繧薙〒縺ｿ繧・            f.read(1)
        return True
    except:
        return False

def delete_matching_files(rip_name, path, log_dir):
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    deleted_files = []
    skipped_files = []  # 繧ｹ繧ｭ繝・・縺励◆繝輔ぃ繧､繝ｫ縺ｮ繝ｪ繧ｹ繝・
    for filename in os.listdir(path):
        if is_valid_tiff(filename):
            full_path = os.path.join(path, filename)
            
            # 繝輔ぃ繧､繝ｫ縺ｮ迥ｶ諷九メ繧ｧ繝・け
            if is_file_locked(full_path):
                print(f"[{rip_name}] 繧ｹ繧ｭ繝・・・医Ο繝・け荳ｭ・・ {filename}")
                skipped_files.append((filename, "繝ｭ繝・け荳ｭ"))
                continue
                
            if not is_file_complete(full_path):
                print(f"[{rip_name}] 繧ｹ繧ｭ繝・・・域悴螳御ｺ・ｼ・ {filename}")
                skipped_files.append((filename, "譛ｪ螳御ｺ・))
                continue

            try:
                os.remove(full_path)
                deleted_files.append(filename)
                print(f"[{rip_name}] 蜑企勁: {filename}")
            except Exception as e:
                print(f"[{rip_name}] 蜑企勁螟ｱ謨・ {filename} 竊・{e}")

    if deleted_files or skipped_files:  # 蜑企勁縺ｾ縺溘・繧ｹ繧ｭ繝・・縺励◆繝輔ぃ繧､繝ｫ縺後≠繧句ｴ蜷・        log_filename = f"{now}_{rip_name}.log"
        log_path = os.path.join(log_dir, log_filename)
        write_detailed_log(log_path, deleted_files, skipped_files)
    else:
        print(f"[{rip_name}] 蜑企勁蟇ｾ雎｡縺ｪ縺励・)

def write_detailed_log(log_path, deleted_files, skipped_files):
    """隧ｳ邏ｰ縺ｪ繝ｭ繧ｰ繧定ｨ倬鹸"""
    with open(log_path, "w", encoding="utf-8") as log_file:
        log_file.write(f"螳溯｡梧律譎・ {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        log_file.write("\n=== 蜑企勁縺励◆繝輔ぃ繧､繝ｫ ===\n")
        for name in deleted_files:
            log_file.write(f"{name}\n")
        log_file.write("\n=== 繧ｹ繧ｭ繝・・縺励◆繝輔ぃ繧､繝ｫ ===\n")
        for name, reason in skipped_files:
            log_file.write(f"{name} (逅・罰: {reason})\n")

def load_config():
    config = configparser.ConfigParser()
    config.read("config.ini", encoding="utf-8")
    validate_config(config)  # 險ｭ螳壼､縺ｮ讀懆ｨｼ繧定ｿｽ蜉
    return config

def run_for_rip(config, rip_name):
    if rip_name not in config:
        print(f"[{rip_name}] 螳夂ｾｩ縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・)
        return

    section = config[rip_name]
    if section.getboolean("enabled", fallback=False):
        path = section.get("path", "")
        if not os.path.isdir(path):
            print(f"[{rip_name}] 繝代せ縺悟ｭ伜惠縺励∪縺帙ｓ: {path}")
            return

        log_dir = config["General"].get("log_dir", "") if "General" in config else ""
        if not log_dir:
            print(f"[{rip_name}] 繝ｭ繧ｰ蜃ｺ蜉帛・縺梧悴險ｭ螳壹〒縺吶・)
            return

        try:
            os.makedirs(log_dir, exist_ok=True)
        except Exception as e:
            print(f"[{rip_name}] 繝ｭ繧ｰ蜃ｺ蜉帛・繧剃ｽ懈・縺ｧ縺阪∪縺帙ｓ縺ｧ縺励◆: {log_dir} 竊・{e}")
            return

        delete_matching_files(rip_name, path, log_dir)
    else:
        print(f"[{rip_name}] 辟｡蜉ｹ蛹悶＆繧後※縺・∪縺吶・)


def run_kick_mode(config, target):
    if target.upper() == "ALL":
        for rip in ["RIP1", "RIP2", "RIP3"]:
            run_for_rip(config, rip)
    else:
        run_for_rip(config, target)

def run_polling_mode(config):
    # getint() 繧・getfloat() 縺ｫ螟画峩
    interval = config["General"].getfloat("polling_interval", fallback=5)
    print(f"繝昴・繝ｪ繝ｳ繧ｰ繝｢繝ｼ繝峨〒襍ｷ蜍輔・interval}蛻・＃縺ｨ縺ｫ螳溯｡後＠縺ｾ縺吶・)
    try:
        while True:
            for rip in ["RIP1", "RIP2", "RIP3"]:
                run_for_rip(config, rip)
            time.sleep(interval * 60)  # 蟆乗焚轤ｹ莉･荳九ｂ豁｣縺励￥險育ｮ励＆繧後ｋ
    except KeyboardInterrupt:
        print("繝昴・繝ｪ繝ｳ繧ｰ繧剃ｸｭ譁ｭ縺励∪縺励◆縲・)

VERSION = "0.1"

def main():
    if len(sys.argv) >= 2 and sys.argv[1] == "--version":
        print(f"count_valid_pixels version {VERSION}")
        return

    print(f"count_valid_pixels version {VERSION} 襍ｷ蜍輔＠縺ｾ縺励◆縲・)

    config = load_config()

    if len(sys.argv) >= 3 and sys.argv[1] == "--kick":
        rip_arg = sys.argv[2]
        run_kick_mode(config, rip_arg)
    else:
        run_polling_mode(config)

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
            raise
    return False

def validate_config(config):
    """險ｭ螳壼､縺ｮ螯･蠖捺ｧ繧呈､懆ｨｼ"""
    if "General" not in config:
        raise ValueError("General 繧ｻ繧ｯ繧ｷ繝ｧ繝ｳ縺悟ｿ・ｦ√〒縺・)
    
    required_general = ["log_dir", "polling_interval"]
    for key in required_general:
        if key not in config["General"]:
            raise ValueError(f"General 繧ｻ繧ｯ繧ｷ繝ｧ繝ｳ縺ｫ {key} 縺悟ｿ・ｦ√〒縺・)
    
    interval = config["General"].getfloat("polling_interval")
    if interval <= 0:
        raise ValueError("polling_interval 縺ｯ豁｣縺ｮ蛟､縺ｧ縺ゅｋ蠢・ｦ√′縺ゅｊ縺ｾ縺・)
    
    for rip in ["RIP1", "RIP2", "RIP3"]:
        if rip in config and config[rip].getboolean("enabled", False):
            if "path" not in config[rip]:
                raise ValueError(f"{rip} 縺ｫ path 縺悟ｿ・ｦ√〒縺・)

if __name__ == "__main__":
    main()


