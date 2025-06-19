import pandas as pd

def convert_and_export(
    src_df: pd.DataFrame,
    mapping_info: list,     # List of (src_col, dst_col, dtype_str)
    output_path: str
):
    """
    mapping_info: list of tuples like (src_col, dst_col, dtype)
    dtype: one of 'str', 'int', 'float', 'bool', 'date'
    """

    result_df = pd.DataFrame()

    for src_col, dst_col, dtype in mapping_info:
        if src_col not in src_df.columns:
            raise ValueError(f"移行元カラム「{src_col}」が見つかりません。")

        series = src_df[src_col]

        # 型変換
        try:
            if dtype == "str":
                converted = series.astype(str)
            elif dtype == "int":
                converted = pd.to_numeric(series, errors='coerce').astype("Int64")  # null対応
            elif dtype == "float":
                converted = pd.to_numeric(series, errors='coerce').astype(float)
            elif dtype == "bool":
                converted = series.astype(bool)
            elif dtype == "date":
                converted = pd.to_datetime(series, errors='coerce').dt.strftime('%Y-%m-%d')
            else:
                raise ValueError(f"未対応のデータ型: {dtype}")
        except Exception as e:
            raise ValueError(f"{src_col} → {dst_col} の型変換エラー: {str(e)}")

        result_df[dst_col] = converted

    # CSV出力
    result_df.to_csv(output_path, index=False, encoding='utf-8-sig')