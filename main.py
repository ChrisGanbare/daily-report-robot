import pandas as pd
import json
import requests
import datetime
import time
import schedule
import os
import urllib.parse
from sqlalchemy import create_engine, text

# ================= 配置区域 (请修改此处) =================
# 1. 数据库配置
DB_CONFIG = {
    'host': '8.139.83.130',  # 数据库IP
    'port': 3306,  # 端口
    'user': 'query_zr',  # 用户名
    'password': 'ZRYLPass220609!',  # 密码 (请修改)
    'db': 'oil',  # 数据库名
    'charset': 'utf8mb4'
}

# 2. API 配置
API_URL = "https://youku.zr228.com/oil-admin/api/v1/device/oil/queryDeviceOilStock"

# 3. 机器人 Webhook (企业微信/钉钉)
WEBHOOK_URL = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=928f052d-7b3a-4137-bb54-8f1528da84e0"

# 4. 配置文件设置
# 方案A：在线表格配置 (填入腾讯文档/飞书表格的导出/下载链接)
# 示例：CONFIG_EXCEL_URL = "https://docs.qq.com/d/export/YOUR_DOC_ID?format=xlsx"
CONFIG_EXCEL_URL = "https://www.kdocs.cn/l/cvpSEwIAZGtE"

# 方案B：本地文件配置 (当在线表格未配置或下载失败时使用)
CONFIG_LOCAL_FILE = 'device_config.xlsx'

# 5. 其他配置
EXCEL_PASSWORD = "AdminPassword2026"


# =======================================================

def get_db_data():
    """从数据库获取基础设备和库存信息 (使用 SQLAlchemy)"""
    print("正在从数据库读取数据...")

    safe_password = urllib.parse.quote_plus(DB_CONFIG['password'])
    db_url = (
        f"mysql+pymysql://{DB_CONFIG['user']}:{safe_password}"
        f"@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['db']}?charset={DB_CONFIG['charset']}"
    )

    try:
        engine = create_engine(db_url)

        sql = """
        SELECT
            d.device_code,
            d.create_time as install_time,
            c.customer_name,
            CONCAT(IFNULL(d.province_name,''), IFNULL(d.city_name,''), IFNULL(d.district_name,'')) as location,
            ot.oil_model,
            o.avai_oil
        FROM
            t_device d
            LEFT JOIN t_customer c ON d.customer_id = c.id
            LEFT JOIN (
                SELECT ta.* FROM t_oil_type ta
                INNER JOIN ( SELECT device_id, max(id) AS id FROM t_oil_type WHERE status=1 GROUP BY device_id ) tb ON ta.id = tb.id 
            ) ot ON d.id = ot.device_id
            LEFT JOIN t_device_oil o ON ot.id = o.oil_type_id
        WHERE
            d.del_status = 1
        """
        df = pd.read_sql(sql, engine)
        return df

    except Exception as e:
        print(f"数据库读取失败: {e}")
        return pd.DataFrame()


def get_api_data_map():
    """调用API获取设备同步时间"""
    print("正在调用API获取同步状态...")
    try:
        resp = requests.get(API_URL, timeout=15)
        if resp.status_code == 200:
            res_json = resp.json()
            if res_json.get('retState') == 'SUCCESS':
                info_list = res_json.get('apiResult', {}).get('infoList', [])
                api_map = {}
                for item in info_list:
                    code = item.get('deviceCode')
                    if code:
                        api_map[code] = {
                            'modifyTime': item.get('modifyTime'),
                            'syncTime': item.get('syncTime')
                        }
                return api_map
    except Exception as e:
        print(f"API 调用异常: {e}")
    return {}


def load_config():
    """加载配置：优先在线表格，失败则读取本地Excel"""
    exclude_list = ["中润", "内部测试"]
    device_config = {}
    
    df_devices = pd.DataFrame()
    df_settings = pd.DataFrame()
    
    loaded_source = None

    # 1. 尝试从在线表格读取
    if CONFIG_EXCEL_URL and "http" in CONFIG_EXCEL_URL:
        try:
            print(f"正在尝试从在线表格加载配置...")
            # 读取所有 sheet
            xls = pd.read_excel(CONFIG_EXCEL_URL, sheet_name=None)
            df_devices = xls.get('设备配置', pd.DataFrame())
            df_settings = xls.get('全局设置', pd.DataFrame())
            loaded_source = "在线表格"
            print("在线配置加载成功")
        except Exception as e:
            print(f"在线表格加载失败: {e}")
            
    # 2. 如果在线加载失败（df为空），尝试读取本地文件
    if loaded_source is None and os.path.exists(CONFIG_LOCAL_FILE):
        try:
            print(f"正在读取本地配置文件: {CONFIG_LOCAL_FILE}")
            xls = pd.read_excel(CONFIG_LOCAL_FILE, sheet_name=None)
            df_devices = xls.get('设备配置', pd.DataFrame())
            df_settings = xls.get('全局设置', pd.DataFrame())
            loaded_source = "本地文件"
        except Exception as e:
            print(f"本地配置文件读取失败: {e}")

    # 3. 解析数据
    if not df_devices.empty:
        # 假设列名：设备编号, 桶数, 设备归属
        # 兼容列名可能的空格
        df_devices.columns = df_devices.columns.str.strip()
        
        for _, row in df_devices.iterrows():
            code = str(row.get('设备编号', '')).strip()
            if code and code != 'nan':
                device_config[code] = {
                    'barrels': row.get('桶数', 1),
                    'owner': row.get('设备归属', '中润')
                }
    
    if not df_settings.empty:
        df_settings.columns = df_settings.columns.str.strip()
        if '排除客户' in df_settings.columns:
            excludes = df_settings['排除客户'].dropna().astype(str).tolist()
            if excludes:
                exclude_list = excludes

    if loaded_source:
        print(f"配置加载完成 (来源: {loaded_source})")
    else:
        print("未加载到有效配置，将使用默认设置")

    return exclude_list, device_config


def process_data(df, api_map):
    """核心逻辑处理"""
    if df.empty:
        return df

    # 1. 加载配置 (修改为调用 load_config)
    exclude_list, device_config = load_config()

    # 2. 排除客户
    if 'customer_name' in df.columns:
        df = df[~df['customer_name'].isin(exclude_list)].copy()

    # 3. === 排序逻辑 ===
    df['install_time'] = pd.to_datetime(df['install_time'])
    df = df.sort_values(by='install_time', ascending=True)
    df['install_time'] = df['install_time'].dt.strftime('%Y.%m.%d').fillna('')

    # 4. 补全配置字段
    df['桶数'] = df['device_code'].apply(lambda x: device_config.get(x, {}).get('barrels', 1))
    df['设备归属'] = df['device_code'].apply(lambda x: device_config.get(x, {}).get('owner', '中润'))

    # 5. 同步状态判断
    now = datetime.datetime.now()

    def check_sync(row):
        code = row['device_code']
        info = api_map.get(code)
        if not info: return "正常"
        m_str, s_str = info.get('modifyTime'), info.get('syncTime')
        if m_str != s_str:
            if m_str:
                try:
                    m_time = datetime.datetime.strptime(str(m_str), "%Y-%m-%d %H:%M:%S")
                    if (now - m_time).total_seconds() > 24 * 3600:
                        return "异常"
                except:
                    pass
        return "正常"

    df['设备数据同步'] = df.apply(check_sync, axis=1)

    # 6. 数据清洗
    df['avai_oil'] = df['avai_oil'].fillna('无数据')
    df['oil_model'] = df['oil_model'].fillna('未设置')

    # 7. 生成序号
    df.reset_index(drop=True, inplace=True)
    df.insert(0, '序号', range(1, 1 + len(df)))

    # 8. 列重命名与筛选
    df = df.rename(columns={
        'customer_name': '客户名称',
        'device_code': '设备编号',
        'oil_model': '油品型号',
        'avai_oil': '库存(%)',
        'install_time': '安装时间',
        'location': '安装地点'
    })

    target_cols = ['序号', '客户名称', '设备编号', '油品型号', '库存(%)', '桶数', '设备归属', '网络同步',
                   '安装时间', '安装地点']
    return df[[c for c in target_cols if c in df.columns]]


def generate_excel_with_format(df, filename):
    """生成Excel：支持筛选、冻结、复制、红绿灯"""
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    sheet_name = '安卓设备日统计表'
    df.to_excel(writer, index=False, sheet_name=sheet_name)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # === 样式 ===
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
    green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    yellow_fmt = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'})
    red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    orange_fmt = workbook.add_format({'bg_color': '#FFCC99', 'font_color': '#333333'})

    # 1. 设置表头
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)

    # 2. 设置列宽
    worksheet.set_column('B:D', 20)
    worksheet.set_column('H:H', 15)
    worksheet.set_column('J:J', 20)

    # === 冻结首行 ===
    worksheet.freeze_panes(1, 0)

    # === 自动筛选 ===
    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

    # 3. 条件格式 (库存列)
    try:
        inv_col_idx = df.columns.get_loc("库存(%)")
    except:
        inv_col_idx = 4

    start_row = 1
    end_row = len(df) + 1

    # 规则应用
    worksheet.conditional_format(start_row, inv_col_idx, end_row, inv_col_idx,
                                 {'type': 'text', 'criteria': 'containing', 'value': '无数据', 'format': orange_fmt})
    worksheet.conditional_format(start_row, inv_col_idx, end_row, inv_col_idx,
                                 {'type': 'cell', 'criteria': '>', 'value': 30, 'format': green_fmt})
    worksheet.conditional_format(start_row, inv_col_idx, end_row, inv_col_idx,
                                 {'type': 'cell', 'criteria': 'between', 'minimum': 5.000001, 'maximum': 30,
                                  'format': yellow_fmt})
    worksheet.conditional_format(start_row, inv_col_idx, end_row, inv_col_idx,
                                 {'type': 'cell', 'criteria': '<=', 'value': 5, 'format': red_fmt})

    # === 保护选项 ===
    worksheet.protect(EXCEL_PASSWORD, options={
        'format_cells': False,
        'format_columns': False,
        'insert_rows': False,
        'delete_rows': False,
        'sort': True,
        'autofilter': True,
        'select_locked_cells': True,
        'select_unlocked_cells': True
    })

    writer.close()
    return filename


def send_to_robot(filename):
    """发送文件"""
    if not os.path.exists(filename) or "YOUR_WEBHOOK" in WEBHOOK_URL:
        print("未配置Webhook或文件不存在，跳过发送。")
        return

    try:
        key = WEBHOOK_URL.split("key=")[1]
        upload_url = f"https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={key}&type=file"

        with open(filename, 'rb') as f:
            files = {'file': (os.path.basename(filename), f,
                              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            resp = requests.post(upload_url, files=files)
            media_id = resp.json().get('media_id')

        if media_id:
            msg = {"msgtype": "file", "file": {"media_id": media_id}}
            requests.post(WEBHOOK_URL, json=msg)
            print("报表发送成功")
        else:
            print(f"上传失败: {resp.text}")
    except Exception as e:
        print(f"发送异常: {e}")


def clean_old_files():
    """清理旧的Excel报表文件"""
    print("正在清理旧报表文件...")
    try:
        files = os.listdir('.')
        for f in files:
            if f.endswith('.xlsx') and "智能油库油量统计表" in f:
                try:
                    os.remove(f)
                    print(f"已删除旧文件: {f}")
                except Exception as e:
                    print(f"删除文件 {f} 失败: {e}")
    except Exception as e:
        print(f"清理文件过程出错: {e}")


def daily_task():
    """主任务"""
    print(f"[{datetime.datetime.now()}] 开始执行...")
    
    clean_old_files()

    df_db = get_db_data()
    api_map = get_api_data_map()

    if df_db.empty:
        print("无数据库数据")
        return

    df_final = process_data(df_db, api_map)

    today_str = datetime.datetime.now().strftime('%Y年%m月%d日')
    filename = f"{today_str}智能油库油量统计表.xlsx"

    generate_excel_with_format(df_final, filename)
    send_to_robot(filename)
    print(f"处理完成: {filename}")


if __name__ == "__main__":
    print("=== 机器人运行中 ===")

    # 测试运行一次
    #daily_task()

    schedule.every().day.at("08:00").do(daily_task)

    while True:
        schedule.run_pending()
        time.sleep(60)
