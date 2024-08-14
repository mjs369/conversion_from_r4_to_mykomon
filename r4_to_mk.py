import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def process_data(df):
    columns_mk = [
        "NO", "従業員コード", "氏名", "フリガナ", "性別", "生年月日", "自宅郵便番号",
        "自宅住所１", "自宅住所２", "自宅住所３", "自宅電話番号", "所属", "所属コード",
        "従業員区分", "役職", "入社日", "退職日", "最終給与", "健康保険料", "介護保険料",
        "厚生年金保険料", "雇用保険料", "支給形態", "時給1", "時給2", "時給3", "時給4",
        "時給5", "時給6", "通勤手当", "賞与", "所得税区分", "控除対象扶養親族等の数",
        "住民税の設定", "住民税5月分まで", "住民税6月分", "住民税7月分以降", "時間の転記先1",
        "時間の転記先2", "時間の転記先3", "時間の転記先4", "時間の転記先5", "時間の転記先6",
        "金額の転記先1", "金額の転記先2", "金額の転記先3", "金額の転記先4", "金額の転記先5",
        "金額の転記先6", "日給1", "日給2", "日給3"
    ]

    # 正規表現パターンを定義
    pattern = (
        r"^(北海道|青森県|岩手県|宮城県|秋田県|山形県|福島県|茨城県|栃木県|群馬県|"
        r"埼玉県|千葉県|東京都|神奈川県|新潟県|富山県|石川県|福井県|山梨県|長野県|"
        r"岐阜県|静岡県|愛知県|三重県|滋賀県|京都府|大阪府|兵庫県|奈良県|和歌山県|"
        r"鳥取県|島根県|岡山県|広島県|山口県|徳島県|香川県|愛媛県|高知県|福岡県|"
        r"佐賀県|長崎県|熊本県|大分県|宮崎県|鹿児島県|沖縄県)"
        r"(.+?)([0-9０-９].*)$"
    )
    # 置き換えのルールを定義
    salary_dict = {
        "日給月給": "月給制",
        "月給": "月給制",
        "日給": "日給制",
        "時給": "時間給制"
    }

    income_tax_dict = {
        "甲欄": "甲欄",
        "乙欄": "乙欄",
        "入力": "その他",
    }

    post_pattern = r"(\d+):\(.+?\) (.+)"

    df_mk = pd.DataFrame(columns = columns_mk, index = df.index)

    # 日付の修正
    df['生年月日'] = pd.to_datetime(df['生年月日'], format='%Y-%m-%d').dt.strftime('%Y/%m/%d')
    df['入社年月日'] = pd.to_datetime(df['入社年月日'], format='%Y-%m-%d').dt.strftime('%Y/%m/%d')
    df['退職年月日'] = pd.to_datetime(df['退職年月日'], format='%Y-%m-%d').dt.strftime('%Y/%m/%d')

    # データの加工
    df[['都道府県', '市町村', '番地']] = df['住所'].str.extract(pattern)
    df['給与区分'] = df['給与区分'].replace(salary_dict)
    df['税表区分'] = df['税表区分'].replace(income_tax_dict)
    df['性別'] = df['性別'].str.replace('性', '')
    df[['従業員区分', '役職名']] = df['役職'].str.extract(post_pattern)
    df['源泉count'] = df['配偶者区分'].apply(lambda x: 1 if x == '源泉控除対象' else 0)

    mapping = {
        '従業員コード': '従業員コード',
        '従業員名 ※': '氏名',
        '従業員名カナ': 'フリガナ',
        '性別': '性別',
        '生年月日': '生年月日',
        '郵便番号': '自宅郵便番号',
        '都道府県': '自宅住所１',
        '市町村': '自宅住所２',
        '番地': '自宅住所３',
        '部門': '所属',
        '部門コード': '所属コード',
        '従業員区分': '従業員区分',
        '役職名': '役職',
        '入社年月日': '入社日',
        '退職年月日': '退職日',
        '給与区分':'支給形態',
        '税表区分': '所得税区分',
        '雇用保険区分':'雇用保険料',
        '給与所得種別':'賞与',
        '住民税の設定方法':'住民税の設定'
    }

    # マッピング辞書に基づいて値を転記
    for src_col, dest_col in mapping.items():
        df_mk[dest_col] = df[src_col]

    si_dict = {
        "あり": "控除する",
        "なし": "控除しない"
    }

    bonus_dict = {
        '給料・賞与':'する',
        '賞与':'する',
        '給与':'しない'
    }

    r_tax_dict = {
        '通常':'金額設定',
        '月別':'前月コピー'
    }

    df_mk['雇用保険料'] = df_mk['雇用保険料'].replace(si_dict)
    df_mk['賞与'] = df_mk['賞与'].replace(bonus_dict)
    df_mk['住民税の設定'] = df_mk['住民税の設定'].replace(r_tax_dict)
    df_mk['控除対象扶養親族等の数'] = df[['一般扶養親族', '特定扶養親族', '同居老親等', 'その他老人', '源泉count']].sum(axis=1)

    return df_mk

st.title('給与R4→Mykomon')

st.markdown("""
    #### 給与R4からの出力手順
    1.02.「設定」→24.「従業員/一覧入力」を選択 \n
    2.Excel(F12)を選択→「はい」をクリック \n
    3.ファイルの種類を(*.xlsx)に変更して保存 \n
    4.下部の指定場所にアップロード
    """)

uploaded_file = st.file_uploader("Excelをアップロードしてください。", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, header=1)
    df_processed = process_data(df)
    
    st.dataframe(df_processed)
    
    # Excelへの出力
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_processed.to_excel(writer, index=False, sheet_name='Sheet1')
    
    processed_data = output.getvalue()

    st.download_button(
        label="Download Processed Data as Excel",
        data=processed_data,
        file_name='processed_data.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
