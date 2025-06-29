import matplotlib
from ai_advice import ai_advice_and_fix
from data_prepare import generate_user_data, generate_transaction_data, save_to_sqlite
from profiling_report import profiling_and_report

matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 显示中文
matplotlib.rcParams['axes.unicode_minus'] = False    # 正常显示负号

# 配置
DB_PATH = 'sqlite:///ds_profiling.db'
REPORT_NAME="ds_profiling_report.docx"
IGNORE_TABLE_LIST = ['transactions']
IGNORE_COLUMN_DICT = {'user_info': ['region']}


if __name__ == '__main__':
    # 1. 生成数据
    print("\n正在创建数据库")
    df_users = generate_user_data(100)
    df_transactions = generate_transaction_data(1000, 100)
    save_to_sqlite(df_users, 'user_info', DB_PATH)
    save_to_sqlite(df_transactions, 'transactions', DB_PATH)

    # 2. profiling分析并生成报告（自动分析所有表）
    print("\n正在profiling分析并生成报告...")
    report = profiling_and_report(db_path=DB_PATH, report_name=REPORT_NAME, ignore_table_list=IGNORE_TABLE_LIST, ignore_column_dict=IGNORE_COLUMN_DICT)

    # 3. AI建议与修复
    print("\n正在基于AI给出修复建议...")
    ai_advice_and_fix(report, report_name=REPORT_NAME)