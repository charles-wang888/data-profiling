import pandas as pd
import numpy as np
import random
from faker import Faker
from sqlalchemy import create_engine

fake = Faker('zh_CN')
regions = ['北京', '上海', '深圳', '广州']


def generate_user_data(n=100):
    users = []
    for i in range(n):
        user = {
            'user_id': i+1,
            'name': fake.name()[:15],
            'email': fake.email(),
            'address': fake.address().replace('\n', '')[:30],
            'region': random.choice(regions),
            'phone': fake.phone_number(),
            'created_date': fake.date_time_this_decade().strftime('%Y-%m-%d %H:%M:%S')
        }
        users.append(user)
    df = pd.DataFrame(users)

    # 制造5%-8%正则不合理数据
    n_invalid = random.randint(int(n*0.05), int(n*0.08))
    invalid_idx = np.random.choice(df.index, n_invalid, replace=False)
    for idx in invalid_idx:
        if random.random() < 0.5:
            df.at[idx, 'phone'] = '123456'  # 明显不合规
        else:
            df.at[idx, 'email'] = 'invalid_email'

    # 制造3%缺失值
    n_missing = max(1, int(n*0.03))
    for col in ['email', 'address', 'phone', 'created_date']:
        missing_idx = np.random.choice(df.index, n_missing, replace=False)
        df.loc[missing_idx, col] = None

    # phone字段5-6条重复
    dup_phones = np.random.choice(df['phone'], 2, replace=False)
    for phone in dup_phones:
        dup_idx = np.random.choice(df.index, 3, replace=False)
        df.loc[dup_idx, 'phone'] = phone

    # email字段3-4条重复
    dup_emails = np.random.choice(df['email'].dropna(), 1, replace=False)
    for email in dup_emails:
        dup_idx = np.random.choice(df.index, 3, replace=False)
        df.loc[dup_idx, 'email'] = email

    return df

def generate_transaction_data(n=1000, user_count=100):
    transactions = []
    for i in range(n):
        transaction = {
            'transaction_id': i+1,
            'user_id': random.randint(1, user_count),
            'product_id': random.randint(1, 50),
            'transaction_time': fake.date_time_this_year().strftime('%Y-%m-%d %H:%M:%S')
        }
        transactions.append(transaction)
    df = pd.DataFrame(transactions)

    # 9%不合理数据
    n_invalid = int(n*0.09)
    invalid_idx = np.random.choice(df.index, n_invalid, replace=False)
    for idx in invalid_idx:
        if random.random() < 0.5:
            df.at[idx, 'user_id'] = 999999  # 不存在的用户
        else:
            df.at[idx, 'transaction_time'] = None  # 缺失交易时间
    return df

def save_to_sqlite(df, table_name, db_path):
    engine = create_engine(db_path)
    df.to_sql(table_name, engine, if_exists='replace', index=False)