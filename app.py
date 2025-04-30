import streamlit as st  # StreamlitをインポートしてWebアプリケーションを構築
import pandas as pd  # pandasをインポートしてデータフレーム操作を行う
import tempfile  # 一時ファイル作成用モジュールをインポート
import os  # OSパス操作用のモジュールをインポート
import pythoncom  # COM初期化用のモジュールをインポート
import win32com.client as win32  # Win32 COMクライアントをインポートしてExcel操作を行う
import io  # メモリ上のバイトバッファを扱うモジュールをインポート
from PyPDF2 import PdfReader, PdfWriter  # PDF読み書き用のクラスをインポート
from reportlab.pdfgen import canvas  # PDF生成用のCanvasをインポート


def main():  # メイン処理の定義開始
    pythoncom.CoInitialize()  # COMライブラリを初期化する
    st.title("Excel to PDF 変換ツール")  # アプリのタイトルを表示

    uploaded_file = st.file_uploader("Excelファイルをアップロード", type=['xlsx'])  # Excelファイルアップロード用ウィジェット
    if uploaded_file:  # ファイルがアップロードされた場合
        # ファイルが変わったらDataFrameと入力状態をリセット
        if 'df' not in st.session_state or st.session_state.get('file_name') != uploaded_file.name:
            df = pd.read_excel(uploaded_file, engine='openpyxl')  # アップロードされたExcelを読み込む
            st.session_state['df'] = df  # 読み込んだDataFrameをセッションに保存
            st.session_state['file_name'] = uploaded_file.name  # ファイル名をセッションに保存
            st.session_state['inputs'] = {}  # 空セル入力用セッション状態を初期化

        df = st.session_state['df']  # セッション上のDataFrameを取得

        st.subheader("Excelデータプレビュー")  # サブヘッダー表示
        st.dataframe(df)  # 読み込んだExcelデータをテーブル表示

        # 空白セルを検出
        mask = df.isna()  # DataFrameの空白セル位置をブールマスクで取得
        blank_positions = [(i, col) for i in df.index for col in df.columns if mask.loc[i, col]]  # 空セルの (行, 列) をリスト化

        if blank_positions:  # 空白セルがある場合
            st.subheader("空白セルの編集")  # サブヘッダーを表示
            for i, col in blank_positions:  # 各空セル位置についてループ
                key = f"{i}_{col}"  # セッションキーを生成
                default = st.session_state['inputs'].get(key, "")  # 既存入力値を取得
                value = st.text_input(f"行 {i+2}, 列 {col}", value=default, key=key)  # テキスト入力ウィジェットを表示
                st.session_state['inputs'][key] = value  # 入力値をセッションに保存
                if value:  # 入力があれば
                    df.at[i, col] = value  # DataFrameの該当セルに値をセット
            st.session_state['df'] = df  # 更新後のDataFrameをセッションに保存
        else:  # 空白セルがない場合
            st.success("空白セルはありません。正常に読み込まれました！")  # 成功メッセージを表示

        st.subheader("オーバーレイテキストの追加")  # オーバーレイ用テキスト入力の見出し
        overlay_text = st.text_area("PDF に重ねるテキスト", value="", height=100)  # テキストエリアで文字列を入力
        x_pos = st.number_input("X座標 (pts)", min_value=0, value=100)  # テキスト表示X座標を数値入力
        y_pos = st.number_input("Y座標 (pts)", min_value=0, value=750)  # テキスト表示Y座標を数値入力

        if st.button("PDFに変換"):  # PDF変換ボタンが押された場合
            with st.spinner("PDFを生成中..."):  # 処理中スピナーを表示
                # 一時Excelファイルに保存
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:  # 一時ファイルを生成
                    excel_path = tmp.name  # ファイルパスを変数に保存
                    df.to_excel(excel_path, index=False)  # DataFrameをExcelファイルとして保存

                # COMでExcelを起動してPDFへ変換
                excel_app = win32.Dispatch("Excel.Application")  # Excelアプリケーションを起動
                excel_app.Visible = False  # Excelウィンドウを非表示設定
                wb = excel_app.Workbooks.Open(os.path.abspath(excel_path))  # 一時Excelを開く
                pdf_path = os.path.splitext(excel_path)[0] + ".pdf"  # PDF出力先パスを生成
                wb.ExportAsFixedFormat(0, pdf_path)  # PDF形式でエクスポート
                wb.Close(False)  # ブックを閉じる
                excel_app.Quit()  # Excelアプリケーションを終了

                # オーバーレイテキストがあればPDFに重ねる
                final_pdf = pdf_path  # 最終的なPDFパスを初期設定
                if overlay_text:  # テキストが入力されている場合
                    existing_pdf = PdfReader(open(pdf_path, "rb"))  # 生成PDFを読み込む
                    output = PdfWriter()  # 新規PDFライターを作成
                    for i in range(len(existing_pdf.pages)):  # 全ページに対してループ
                        page = existing_pdf.pages[i]  # ページを取得
                        packet = io.BytesIO()  # メモリバッファを作成
                        w = page.mediabox.upper_right[0]  # ページ幅を取得
                        h = page.mediabox.upper_right[1]  # ページ高さを取得
                        can = canvas.Canvas(packet, pagesize=(w, h))  # Canvasを初期化
                        can.drawString(x_pos, y_pos, overlay_text)  # テキストを描画
                        can.save()  # Canvasを保存
                        packet.seek(0)  # バッファ位置を先頭に移動
                        overlay_pdf = PdfReader(packet)  # バッファからPDFを読み込む
                        overlay_page = overlay_pdf.pages[0]  # オーバーレイPDFのページを取得
                        page.merge_page(overlay_page)  # 元ページに合成
                        output.add_page(page)  # 出力用PDFにページを追加
                    overlayed_pdf_path = os.path.splitext(excel_path)[0] + "_overlay.pdf"  # 合成後PDFのパスを設定
                    with open(overlayed_pdf_path, "wb") as f:  # 新規ファイルを書き込みモードで開く
                        output.write(f)  # 合成PDFを出力
                    final_pdf = overlayed_pdf_path  # 最終PDFパスを更新

                # PDFを読み込んでダウンロード
                with open(final_pdf, "rb") as f:  # 最終PDFをバイナリモードで開く
                    pdf_data = f.read()  # データを読み込む
                st.download_button("Download PDF", data=pdf_data, file_name="output.pdf", mime="application/pdf")  # PDFダウンロードボタンを表示
                st.success("PDFの生成が完了しました！")  # 完了メッセージを表示


if __name__ == "__main__":  # このスクリプトが直接実行された場合
    main()  # メイン処理を呼び出す 