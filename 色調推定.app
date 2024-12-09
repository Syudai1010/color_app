import streamlit as st
import pandas as pd
from io import BytesIO
import colour

# 色調→マンセル値の対応表（「灰」を追加）
color_to_munsell = {
    "暗灰褐": "10YR 3/2",
    "暗灰": "N 3",
    "褐灰": "10YR 5/2",
    "灰褐": "10YR 4/2",
    "茶褐": "7.5YR 3/4",
    "茶灰": "7.5YR 5/2",
    "白灰": "N 8",
    "緑灰": "5G 6/1",
    "黄灰": "2.5Y 7/2",
    "淡暗灰": "N 4",
    "灰白": "N 7",
    "黒灰": "N 2",
    "淡黄褐": "2.5Y 6/4",
    "暗褐": "7.5YR 3/2",
    "黄褐": "10YR 6/4",
    "青灰": "5PB 5/2",
    "乳灰": "N 9",
    "暗茶": "5YR 3/2",
    "灰": "N 5"  # 「灰」を追加
}

def munsell_to_rgb(munsell_str):
    """Convert Munsell value to RGB using colour-science library."""
    try:
        munsell_str = munsell_str.strip()
        # 中立色の場合、グレースケールを計算
        if munsell_str.startswith("N"):
            # マンセルの値部分を抽出
            parts = munsell_str.split()
            if len(parts) != 2:
                raise ValueError(f"Invalid neutral Munsell notation: '{munsell_str}'")
            value_str = parts[1]
            # Valueを取得
            if "/" in value_str:
                value = float(value_str.split("/")[0])
            else:
                value = float(value_str)
            # Valueは0から10の範囲なので、0〜255にスケール
            grayscale = int(round((value / 10) * 255))
            return (grayscale, grayscale, grayscale)
        else:
            # Munsell値をCIE xyYに変換
            xyY = colour.notation.munsell.munsell_colour_to_xyY(munsell_str)
            # CIE xyYからsRGBに変換
            rgb = colour.XYZ_to_sRGB(colour.xyY_to_XYZ(xyY))
            # 値が0未満または1を超える場合はクリップ
            rgb = [min(max(channel, 0), 1) for channel in rgb]
            # 0〜255スケールに変換
            rgb_255 = tuple(int(round(channel * 255)) for channel in rgb)
            return rgb_255
    except Exception as e:
        st.warning(f"マンセル値 '{munsell_str}' の変換中にエラーが発生しました: {e}")
        # エラー時のデフォルト値
        return (None, None, None)

def color_name_to_munsell_and_rgb(color_name):
    """Convert color name to Munsell and RGB."""
    base_color = color_name.split("～")[0]  # 「～」で分割し、最初の色を基にする
    munsell_value = color_to_munsell.get(base_color, None)
    if munsell_value is None:
        return (None, None, None, None)  # 対応する色名がなければNone
    # マンセル値からRGB値へ変換
    rgb = munsell_to_rgb(munsell_value)
    return (munsell_value, *rgb)

# Streamlitアプリ開始
st.title("色調→マンセル値→RGB値 変換アプリ")

# ファイルアップローダー
uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Excel読み込み
        df = pd.read_excel(uploaded_file)

        # 「色調」列が存在するか確認
        if "色調" not in df.columns:
            st.error("アップロードされたファイルに「色調」列が見つかりません。")
        else:
            # 色調→マンセル値→RGB変換
            df[["マンセル値", "R", "G", "B"]] = df["色調"].apply(
                lambda x: pd.Series(color_name_to_munsell_and_rgb(x))
            )

            # 結果表示
            st.write("変換結果")
            st.dataframe(df)

            # 結果をExcelでダウンロードできるようにする
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
            output.seek(0)

            st.download_button(
                label="変換結果をダウンロード",
                data=output,
                file_name="output_with_rgb.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"ファイルの処理中にエラーが発生しました: {e}")
