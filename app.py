from datetime import date, datetime, timedelta
import io
import os
import urllib.request
import pandas as pd
import requests
import streamlit as st

# ReportLab関連のインポート
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

# --------------------------------------------------
# 1. ページ初期設定
# --------------------------------------------------
st.set_page_config(page_title="受注一括帳票ジェネレーター (仕入・外注・売上)", layout="wide")

st.title("📑 受注一括帳票ジェネレーター (@pocket API連携版)")
st.markdown("💡 @pocketの詳細画面から連携起動、または受注番号を指定して**「仕入発注書」「外注発注書」「売上納品書」**を一括生成・PDFダウンロードできます。")

# --------------------------------------------------
# 2. 日本語フォント登録処理
# --------------------------------------------------
def setup_japanese_font():
    local_font_path = "NotoSansJP-Regular.ttf"
    if os.path.exists(local_font_path):
        try:
            pdfmetrics.registerFont(TTFont("JPFont", local_font_path))
            return "JPFont"
        except Exception:
            pass

    win_fonts = [
        os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts", "msgothic.ttc"),
        os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts", "meiryo.ttc"),
        os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts", "YuGothM.ttc"),
    ]
    for font_path in win_fonts:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont("JPFont", font_path, subfontIndex=0))
                return "JPFont"
            except Exception:
                pass

    font_url = "https://github.com/google/fonts/raw/main/ofl/notosansjp/NotoSansJP-Regular.ttf"
    try:
        urllib.request.urlretrieve(font_url, local_font_path)
        if os.path.exists(local_font_path):
            pdfmetrics.registerFont(TTFont("JPFont", local_font_path))
            return "JPFont"
    except Exception:
        pass

    return "Helvetica"

FONT_NAME = setup_japanese_font()

# --------------------------------------------------
# 3. セキュリティ設定（どのようなキー名でも自動取得する柔軟対応）
# --------------------------------------------------
api_key = ""
try:
    # 考えられるキー名を順番に探索して最初に見つかったものを採用
    possible_keys = ["pockets_api_key", "app36_api_key", "ATPOCKET_API_TOKEN", "api_key", "atpocket_api_key"]
    for k in possible_keys:
        if k in st.secrets:
            api_key = st.secrets[k]
            break
    
    # それでも見つからない場合で、Secrets自体に何か値が入っていれば最初のものを取得
    if not api_key and len(st.secrets.keys()) > 0:
        first_key = list(st.secrets.keys())[0]
        api_key = st.secrets[first_key]
except Exception:
    pass

# 固定の接続情報
APP_BASE_HOST = "https://app050.at-pocket.com/seihon02_bb"
APP_ID = 51

# URLのクエリパラメータから「受注番号」を取得
query_params = st.query_params
order_id_from_url = query_params.get("order_id", "")

# --------------------------------------------------
# 4. サイドバー設定
# --------------------------------------------------
st.sidebar.header("⚙️ 検索コントロール")
input_order_id = st.sidebar.text_input("受注番号 (field-1)", value=order_id_from_url)
fetch_btn = st.sidebar.button("🔍 @pocketからデータを検索する", type="primary")

auto_fetch = bool(order_id_from_url)


# --------------------------------------------------
# 5. フィールド値抽出ヘルパー
# --------------------------------------------------
def extract_val(raw_val):
    if raw_val is None:
        return ""
    if isinstance(raw_val, dict):
        return str(raw_val.get("value", raw_val.get("name", ""))).strip()
    elif isinstance(raw_val, list) and len(raw_val) > 0:
        extracted = []
        for item in raw_val:
            if isinstance(item, dict):
                extracted.append(str(item.get("value", item.get("name", ""))))
            else:
                extracted.append(str(item))
        return " ".join([e for e in extracted if e and e.lower() != "none"])
    else:
        val_str = str(raw_val).strip()
        if val_str.lower() == "none":
            return ""
        return val_str


# --------------------------------------------------
# 6. APIデータ取得関数
# --------------------------------------------------
def fetch_pocket_record_by_order_no(api_key, base_host, app_id, order_no):
    url = f"{base_host}/api/apps/{app_id}/records"
    headers = {"X-At-Pocket-API-Key": api_key, "Accept": "application/json"}
    
    params = {
        "query": f'field-1 = "{order_no}"'
    }

    try:
        response = requests.get(url, headers=headers, params=params, timeout=15)
        if response.status_code != 200:
            response = requests.get(url, headers=headers, timeout=15)

        if response.status_code == 200:
            data = response.json()
            records = data.get("records", data.get("data", []))
            
            for r in records:
                inner = r.get("record", r)
                curr_order_no = extract_val(inner.get("field-1", ""))
                if str(curr_order_no) == str(order_no):
                    return inner
            
            if records and not order_no:
                return records[0].get("record", records[0])
            return None
        else:
            st.error(f"APIエラーが発生しました（ステータスコード: {response.status_code}）")
            return None
    except Exception as e:
        st.error(f"通信エラーが発生しました: {e}")
        return None


# --------------------------------------------------
# 7. サブテーブルパーサー
# --------------------------------------------------
def parse_sub_table_items(record, sub_table_field, field_map):
    sub_list = record.get(sub_table_field, "")
    parsed_items = []
    
    if isinstance(sub_list, list) and len(sub_list) > 0:
        for i, sub_item in enumerate(sub_list):
            sub_rec = sub_item.get("record", sub_item) if isinstance(sub_item, dict) else {}

            def get_val(f_key):
                if not f_key:
                    return ""
                val = sub_rec.get(f_key)
                if val is not None and val != "":
                    return extract_val(val)
                val_container = record.get(f_key)
                if isinstance(val_container, list) and len(val_container) > i:
                    return extract_val(val_container[i])
                return ""

            item_dict = {k: get_val(v) for k, v in field_map.items()}
            parsed_items.append(item_dict)
    return parsed_items


# --------------------------------------------------
# 8. PDF生成関数
# --------------------------------------------------
def generate_document_pdf(doc_type, header_info, items):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", fontName=FONT_NAME, fontSize=16, leading=20, textColor=colors.HexColor("#1E3A8A"), alignment=1)
    sub_style = ParagraphStyle("Sub", fontName=FONT_NAME, fontSize=9, leading=14, textColor=colors.HexColor("#1E293B"))
    th_style = ParagraphStyle("TH", fontName=FONT_NAME, fontSize=8, leading=10, textColor=colors.white, alignment=1)
    td_text = ParagraphStyle("TDText", fontName=FONT_NAME, fontSize=8, leading=11, textColor=colors.HexColor("#1E293B"))
    td_center = ParagraphStyle("TDCenter", fontName=FONT_NAME, fontSize=8, leading=11, textColor=colors.HexColor("#1E293B"), alignment=1)

    elements = []
    elements.append(Paragraph(f"<b>{doc_type}</b>", title_style))
    elements.append(Spacer(1, 10))

    info_text = f"<b>受注番号:</b> {header_info.get('order_id', '')} ｜ <b>書名:</b> {header_info.get('book_title', '')} ｜ <b>得意先:</b> {header_info.get('customer', '')}"
    elements.append(Paragraph(info_text, sub_style))
    elements.append(Spacer(1, 15))

    if "発注書" in doc_type:
        headers = ["宛先 / 担当", "納期希望日", "内容明細", "数量"]
        col_widths = [120, 90, 205, 100]
        table_data = [[Paragraph(h, th_style) for h in headers]]
        
        for item in items:
            table_data.append([
                Paragraph(item.get("vendor", ""), td_text),
                Paragraph(item.get("date", ""), td_center),
                Paragraph(item.get("content", ""), td_text),
                Paragraph(item.get("qty", ""), td_center),
            ])
    else:
        headers = ["品名", "内容", "数量", "売上単価"]
        col_widths = [180, 155, 85, 95]
        table_data = [[Paragraph(h, th_style) for h in headers]]
        
        for item in items:
            table_data.append([
                Paragraph(item.get("name", ""), td_text),
                Paragraph(item.get("content", ""), td_text),
                Paragraph(item.get("qty", ""), td_center),
                Paragraph(item.get("price", ""), td_center),
            ])

    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor("#CBD5E1")),
    ]))
    
    elements.append(t)
    doc.build(elements)
    buffer.seek(0)
    return buffer


# --------------------------------------------------
# 9. メイン画面の処理
# --------------------------------------------------
target_order_id = input_order_id.strip()

if (fetch_btn or auto_fetch) and target_order_id:
    if not api_key:
        st.error("❌ APIキーが設定されていません。Streamlit Community Cloudの Secrets 設定画面でAPIキーが正しく登録されているか確認してください。")
    else:
        with st.spinner(f"受注番号 [{target_order_id}] のデータを@pocketから取得中..."):
            record = fetch_pocket_record_by_order_no(api_key, APP_BASE_HOST, APP_ID, target_order_id)
            
            if not record:
                st.warning(f"⚠️ 指定された受注番号「{target_order_id}」のデータが見つかりませんでした。")
            else:
                header_info = {
                    "order_id": extract_val(record.get("field-1", target_order_id)),
                    "book_title": extract_val(record.get("field-2", "")),
                    "customer": extract_val(record.get("field-3", ""))
                }
                
                sales_map = {"name": "field-10_3", "content": "field-10_8", "qty": "field-10_4", "price": "field-10_5"}
                sales_items = parse_sub_table_items(record, "field-10", sales_map)
                
                outsource_map = {"vendor": "field-12_1", "date": "field-12_2", "content": "field-12_3", "qty": "field-12_5"}
                outsource_items = parse_sub_table_items(record, "field-12", outsource_map)
                
                purchase_map = {"vendor": "field-11_1", "date": "field-11_2", "content": "field-11_3", "qty": "field-11_4"}
                purchase_items = parse_sub_table_items(record, "field-11", purchase_map)

                st.success(f"✅ データ取得成功！ （書名: {header_info['book_title']} / 得意先: {header_info['customer']}）")

                tab1, tab2, tab3 = st.tabs(["📦 仕入発注書", "🛠️ 外注発注書", "💰 売上納品書"])

                with tab1:
                    st.subheader("📦 仕入発注書 (field-11)")
                    if purchase_items:
                        st.dataframe(pd.DataFrame(purchase_items), use_container_width=True)
                        pdf_p = generate_document_pdf("仕入発注書", header_info, purchase_items)
                        st.download_button("📥 仕入発注書PDFをダウンロード", data=pdf_p, file_name=f"仕入発注書_{target_order_id}.pdf", mime="application/pdf", key="dl_purchase")
                    else:
                        st.info("仕入明細はありません。")

                with tab2:
                    st.subheader("🛠️ 外注発注書 (field-12)")
                    if outsource_items:
                        st.dataframe(pd.DataFrame(outsource_items), use_container_width=True)
                        pdf_o = generate_document_pdf("外注発注書", header_info, outsource_items)
                        st.download_button("📥 外注発注書PDFをダウンロード", data=pdf_o, file_name=f"外注発注書_{target_order_id}.pdf", mime="application/pdf", key="dl_outsource")
                    else:
                        st.info("外注明細はありません。")

                with tab3:
                    st.subheader("💰 売上納品書 (field-10)")
                    if sales_items:
                        st.dataframe(pd.DataFrame(sales_items), use_container_width=True)
                        pdf_s = generate_document_pdf("売上納品書", header_info, sales_items)
                        st.download_button("📥 売上納品書PDFをダウンロード", data=pdf_s, file_name=f"売上納品書_{target_order_id}.pdf", mime="application/pdf", key="dl_sales")
                    else:
                        st.info("売上明細はありません。")

else:
    st.info("👈 @pocketの詳細画面にある「一括帳票を出力」ボタンを押すか、左側に受注番号を入力してください。")