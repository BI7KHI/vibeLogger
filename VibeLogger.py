import datetime
import os
import json
from openpyxl import Workbook, load_workbook

# 配置文件路径
CONFIG_FILE = "log_config.json"
EXCEL_FILE = "台网点名日志.xlsx"

# 初始默认配置
DEFAULT_CONFIG = {
    "QTH": ["广州", "深圳", "珠海", "中山", "惠州", "东莞"],
    "Rig": ["UV-K5", "UV-K6", "森海克斯8800", "八重洲FT-65R"],
    "Power": ["5W", "10W", "25W", "50W", "100W"],
    "Antenna": ["原装天线", "老鹰775拉杆天线", "IOO天线", "车载吸盘天线"]
}

def load_config():
    """读取配置文件，不存在则创建"""
    if not os.path.exists(CONFIG_FILE):
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return DEFAULT_CONFIG

def save_config(config):
    """将配置保存到 JSON 文件"""
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

def smart_input(prompt, config_key, config_data, default_val=None):
    """
    智能输入助手（具备自学习功能）：
    - 输入序号：返回对应内容
    - 输入新内容：自动存入联想库
    - 直接回车：使用默认值
    """
    options = config_data[config_key]
    print(f"\n>>> 选择/输入 {prompt}")
    for i, opt in enumerate(options, 1):
        print(f"  {i}. {opt}")
    
    hint = f"请输入序号或新内容 (默认: {default_val}): " if default_val else "请输入序号或新内容: "
    user_val = input(hint).strip()
    
    # 1. 处理回车默认值
    if not user_val and default_val:
        return default_val
    
    # 2. 处理数字序号选择
    if user_val.isdigit():
        idx = int(user_val) - 1
        if 0 <= idx < len(options):
            return options[idx]
    
    # 3. 处理新内容（自学习逻辑）
    if user_val and user_val not in options:
        # 如果是新内容，加入列表并更新配置文件
        options.append(user_val)
        config_data[config_key] = options
        save_config(config_data)
        print(f"  ✨ 已将新词汇 '{user_val}' 加入联想库")
    
    return user_val if user_val else "N/A"

def create_log():
    config = load_config()
    
    # 初始化 Excel
    if os.path.exists(EXCEL_FILE):
        try:
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
        except:
            print("\n❌ 错误：Excel 被占用，请关闭后重新运行程序。")
            return
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "点名日志"
        ws.append(["序号", "时间", "呼号", "QTH", "信号报告", "设备", "功率", "天馈", "留言"])

    print("="*50)
    print("      业余无线电台网日志助手 (智能联想输入版)      ")
    print("  输入库外新内容将自动保存，下次输入只需按序号  ")
    print("="*50)

    while True:
        next_seq = ws.max_row
        current_time = datetime.datetime.now().strftime("%H:%M")
        print(f"\n【No.{next_seq} | {current_time}】")

        # 1. 呼号 (强制大写)
        call_raw = input("请输入呼号 (Callsign): ")
        callsign = call_raw.upper()
        if not callsign: continue # 呼号不能为空

        # 2. QTH (自学习)
        qth = smart_input("QTH (所在地)", "QTH", config)

        # 3. 信号报告
        rst = input("请输入 RST [默认 59]: ").strip() or "59"

        # 4. 设备 (自学习)
        rig = smart_input("设备 (Rig)", "Rig", config)

        # 5. 功率 (自学习)
        pwr = smart_input("功率 (Power)", "Power", config, default_val="5W")

        # 6. 天馈 (自学习)
        ant = smart_input("天馈 (Antenna)", "Antenna", config)

        # 7. 留言
        msg = input("讨论话题及留言: ").strip() or "73"

        # 写入并实时保存
        ws.append([next_seq, current_time, callsign, qth, rst, rig, pwr, ant, msg])
        wb.save(EXCEL_FILE)
        print(f"\n✅ 记录成功！")

        if input("\n[回车] 下一位，[n] 退出: ").lower() == 'n':
            break

if __name__ == "__main__":
    create_log()
