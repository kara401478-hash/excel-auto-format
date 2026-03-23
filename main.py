import pandas as pd
import os

def load_excel(filepath):
    """Excelファイルを読み込む"""
    try:
        df = pd.read_excel(filepath)
        print(f"✅ Excelファイルを読み込みました: {filepath}")
        print(f"   行数: {len(df)}, 列数: {len(df.columns)}")
        return df
    except Exception as e:
        print(f"❌ 読み込みエラー: {e}")
        return None

def remove_duplicates(df):
    """重複データを検出・削除する"""
    duplicates = df[df.duplicated()]
    print(f"\n🔍 重複データ検出: {len(duplicates)}件")
    df_clean = df.drop_duplicates()
    print(f"✅ 重複削除後: {len(df_clean)}行")
    return df_clean, duplicates

def clean_data(df):
    """空白セルの処理"""
    before = len(df)
    df = df.dropna(how='all')
    df = df.fillna('')
    after = len(df)
    print(f"✅ 空白行削除: {before - after}行削除")
    return df

def format_data(df):
    """データ整形"""
    # 文字列の前後空白を除去
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].str.strip()
    print("✅ 文字列の前後空白を除去しました")
    return df

def save_excel(df, filepath):
    """Excelファイルを保存する"""
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    df.to_excel(filepath, index=False)
    print(f"✅ 保存しました: {filepath}")

def main():
    # サンプルデータ生成
    sample_data = {
        '名前': ['田中太郎', '鈴木花子', '田中太郎', '佐藤次郎', '鈴木花子', '山田三郎'],
        '部署': ['営業', '総務', '営業', '開発', '総務', '開発'],
        'メール': ['tanaka@example.com', 'suzuki@example.com', 
                  'tanaka@example.com', 'sato@example.com',
                  'suzuki@example.com', 'yamada@example.com'],
        '売上': [100000, 150000, 100000, 200000, 150000, 180000]
    }
    os.makedirs('data', exist_ok=True)
    df_sample = pd.DataFrame(sample_data)
    df_sample.to_excel('data/sample.xlsx', index=False)
    print("✅ サンプルExcelを生成しました: data/sample.xlsx")

    # メイン処理
    df = load_excel('data/sample.xlsx')
    if df is not None:
        df_clean, duplicates = remove_duplicates(df)
        df_clean = clean_data(df_clean)
        df_clean = format_data(df_clean)

        save_excel(df_clean, 'output/formatted.xlsx')
        if len(duplicates) > 0:
            save_excel(duplicates, 'output/duplicates.xlsx')

        print(f"\n🎉 完了！outputフォルダを確認してください。")
        print(f"   整形後データ: {len(df_clean)}行")
        print(f"   削除した重複: {len(duplicates)}行")

if __name__ == "__main__":
    main()
