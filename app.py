import sqlite3
import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from rembg import remove
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import streamlit as st

# 폰트 등록
pdfmetrics.registerFont(TTFont("NanumGothic", "NanumGothic.ttf"))
pdfmetrics.registerFont(TTFont("NotoSansKR", "NotoSansKR-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Roboto", "Roboto-Regular.ttf"))

# Streamlit UI 스타일링
st.markdown("""
<style>
    .main-title {
        color: #FF5733;
        font-size: 36px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 20px;
        font-family: 'NotoSansKR';
    }
    .section-header {
        color: #333;
        font-size: 24px;
        font-weight: bold;
        margin-top: 20px;
        margin-bottom: 10px;
        font-family: 'Roboto';
    }
    .stButton > button {
        background-color: #FF5733;
        color: white;
        border-radius: 8px;
        padding: 12px 24px;
        font-size: 16px;
        border: none;
        margin: 5px;
        font-family: 'NanumGothic';
    }
    .stButton > button:hover {
        background-color: #C70039;
    }
    .stSelectbox, .stTextInput, .stMultiselect {
        background-color: #f9f9f9;
        border-radius: 8px;
        padding: 8px;
        font-family: 'NotoSansKR';
    }
    .stImage {
        border: 2px solid #ddd;
        border-radius: 8px;
        padding: 5px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .stWarning, .stSuccess, .stError {
        background-color: #f0f0f0;
        border-radius: 8px;
        padding: 12px;
        margin: 10px 0;
        font-family: 'NanumGothic';
    }
    .expander-content {
        background-color: #f9f9f9;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

# 엑셀 파일에서 데이터 읽기
def fetch_pos_data():
    try:
        df = pd.read_excel("items.xlsx")
        expected_columns = ["Name", "Price", "PriceQualityText"]
        optional_columns = ["AdditionalPrice"]
        missing_columns = [col for col in expected_columns if col not in df.columns]
        if missing_columns:
            st.error(f"엑셀 파일에 다음 필수 열이 누락되었습니다: {missing_columns}")
            return []
        st.write("엑셀 파일 열 이름:", df.columns.tolist())
        items = df.to_dict("records")
        return items
    except Exception as e:
        st.error(f"엑셀 파일 읽기 오류: {str(e)}")
        return []

# 로컬 이미지 확인
def get_local_image(item_name):
    mapping = {
        "생수 500ml": "water.png",
        "초콜릿 바": "chocolate.png",
        "감자칩 100g": "chips.png",
        "우유 1L": "milk.png",
        "라면 5봉": "ramen.png",
        "커피 200ml": "coffee.png",
        "비누 100g": "soap.png",
        "치약 100g": "toothpaste.png",
        "샴푸 500ml": "shampoo.png",
        "세제 1L": "detergent.png",
        "CJ 명가 재래김/파래김": "cj_seaweed.png"  # CJ 명가 이미지 추가
    }
    filename = mapping.get(item_name)
    if filename and os.path.exists(f"images/{filename}"):
        return f"images/{filename}"
    return None

# 이미지 누끼 처리
def process_image(input_path, item_name):
    if not input_path:
        return None
    output_path = f"images/processed_{item_name}.png"
    input_img = Image.open(input_path)
    output_img = remove(input_img)
    output_img.save(output_path)
    return output_path

# 디스플레이 미리보기 생성 (A4 1/3 크기, 단일 품목)
def create_display_preview(template_id, item, title, footer_text, show_price_quality=False):
    width, height = 595, 280  # A4 1/3 크기 (72dpi 기준, 210mm x 99mm)
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype("NotoSansKR-Regular.ttf", 20)
    price_font = ImageFont.truetype("Roboto-Regular.ttf", 36)
    small_font = ImageFont.truetype("NanumGothic.ttf", 16)
    smaller_font = ImageFont.truetype("NanumGothic.ttf", 12)

    # 전체 테두리
    draw.rectangle((0, 0, width - 1, height - 1), outline="black", width=2)

    # 상단 문구 ("가격은 거꾸로 품질은 제대로", 엑셀 체크 여부)
    y_offset = 0
    if show_price_quality:
        draw.rectangle((0, 0, width, 30), fill=(0, 51, 102))  # 다크 블루 배경
        draw.text((width / 2, 15), "가격은 거꾸로 품질은 제대로", font=smaller_font, fill="white", anchor="mm")
        y_offset = 30

    # 제목 ("가격역주행", 사용자 입력)
    draw.rectangle((0, y_offset, width, y_offset + 50), fill=(255, 165, 0))  # 주황색 배경
    draw.text((width / 2, y_offset + 25), title, font=font, fill="white", anchor="mm")

    # 품목명 (엑셀 데이터)
    margin = 10
    x, y = margin, y_offset + 60
    draw.text((x, y), item["Name"], font=small_font, fill="black")

    # 가격 (엑셀 데이터)
    y += 30
    draw.text((x, y), f"₩{item['Price']:,.0f} (각)", font=price_font, fill="red")
    if item.get("AdditionalPrice"):
        y += 30
        draw.text((x, y), f"10g당 {item['AdditionalPrice']}원", font=smaller_font, fill="black")

    # 품목 이미지 (CJ명가 스타일, 4개 이미지 배치)
    if item["ProcessedImagePath"]:
        try:
            item_img = Image.open(item["ProcessedImagePath"])
            item_img.thumbnail((60, 60))
            for i in range(4):  # 4개 이미지 배치
                img.paste(item_img, (width - 70 - i * 70, y_offset + 60), item_img if item_img.mode == "RGBA" else None)
        except Exception as e:
            st.error(f"이미지 로드 오류: {str(e)}")
            draw.text((width - 70, y_offset + 60), "[이미지 없음]", font=smaller_font, fill="black")
    else:
        st.warning(f"{item['Name']}의 이미지가 없습니다.")
        draw.text((width - 70, y_offset + 60), "[이미지 없음]", font=smaller_font, fill="black")

    # 하단 문구 (사용자 입력)
    draw.rectangle((0, height - 20, width, height), fill=(255, 165, 0))
    draw.text((width / 2, height - 10), footer_text, font=smaller_font, fill="white", anchor="mm")

    preview_path = f"preview_display_{template_id}.png"
    img.save(preview_path)
    return preview_path

# 디스플레이 PDF 생성
def render_display(template_id, item, output_path, title, footer_text, show_price_quality=False):
    c = canvas.Canvas(output_path, pagesize=(210 * mm, 99 * mm))
    width, height = 210 * mm, 99 * mm
    margin = 5 * mm

    # 전체 테두리
    c.setStrokeColor(colors.black)
    c.rect(0, 0, width, height, stroke=1, fill=0)

    # 상단 문구 ("가격은 거꾸로 품질은 제대로", 엑셀 체크 여부)
    y_offset = 0
    if show_price_quality:
        c.setFillColorRGB(0, 0.2, 0.4)  # 다크 블루 배경
        c.rect(0, height - 10 * mm, width, 10 * mm, fill=1, stroke=0)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("NanumGothic", 12)
        c.drawCentredString(width / 2, height - 7 * mm, "가격은 거꾸로 품질은 제대로")
        y_offset = 10 * mm

    # 제목 ("가격역주행", 사용자 입력)
    c.setFillColorRGB(1, 0.65, 0)
    c.rect(0, height - y_offset - 15 * mm, width, 15 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("NotoSansKR", 20)
    c.drawCentredString(width / 2, height - y_offset - 10 * mm, title)

    # 품목명
    x, y = margin, height - y_offset - 25 * mm
    c.setFont("NanumGothic", 16)
    c.drawString(x, y, item["Name"])

    # 가격
    y -= 10 * mm
    c.setFillColorRGB(1, 0, 0)
    c.setFont("Roboto", 36)
    c.drawString(x, y, f"₩{item['Price']:,.0f} (각)")
    c.setFillColorRGB(0, 0, 0)
    if item.get("AdditionalPrice"):
        y -= 10 * mm
        c.setFont("NanumGothic", 12)
        c.drawString(x, y, f"10g당 {item['AdditionalPrice']}원")

    # 품목 이미지 (CJ명가 스타일, 4개 이미지 배치)
    if item["ProcessedImagePath"]:
        try:
            for i in range(4):
                c.drawImage(item["ProcessedImagePath"], width - 20 * mm - i * 20 * mm, height - y_offset - 25 * mm, 15 * mm, 15 * mm)
        except:
            c.drawString(width - 20 * mm, height - y_offset - 25 * mm, "[이미지 없음]")

    # 하단 문구
    c.setFillColorRGB(1, 0.65, 0)
    c.rect(0, 0, width, 5 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("NanumGothic", 12)
    c.drawCentredString(width / 2, 2.5 * mm, footer_text)
    c.save()

# 전단지 미리보기 생성 (여러 품목)
def create_flyer_preview(template_id, items, title, footer_text):
    width, height = 595, 842  # A4 크기 (72dpi)
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype("NotoSansKR-Regular.ttf", 20)
    small_font = ImageFont.truetype("Roboto-Regular.ttf", 12)
    smaller_font = ImageFont.truetype("NanumGothic.ttf", 10)

    # 주황색 배경 제목
    draw.rectangle((0, 0, width, 50), fill=(255, 165, 0))
    draw.text((width / 2, 25), title, font=font, fill="white", anchor="mm")

    # 템플릿별 품목 표시
    if template_id == "1":  # 5개 품목
        cell_width, cell_height = 110, 150
        margin = 20
        x, y = margin, 70
        for i, item in enumerate(items[:5]):
            draw.rectangle((x, y, x + cell_width, y + cell_height), outline="black")
            if item["ProcessedImagePath"]:
                try:
                    item_img = Image.open(item["ProcessedImagePath"])
                    item_img.thumbnail((90, 90))
                    img.paste(item_img, (x + 10, y + 10), item_img if item_img.mode == "RGBA" else None)
                except:
                    draw.text((x + 10, y + 10), "[이미지 없음]", font=smaller_font, fill="black")
            draw.text((x + 10, y + 110), item["Name"][:20], font=small_font, fill="black")
            draw.text((x + 10, y + 125), f"₩{item['Price']:,.0f} (각)", font=smaller_font, fill="red")
            if item.get("AdditionalPrice"):
                draw.text((x + 10, y + 140), f"10g당 {item['AdditionalPrice']}원", font=smaller_font, fill="black")
            x += cell_width + 10
            if (i + 1) % 5 == 0:
                x = margin
                y += cell_height + 10

    # 하단 주황색 배경 텍스트
    draw.rectangle((0, height - 30, width, height), fill=(255, 165, 0))
    draw.text((width / 2, height - 15), footer_text, font=smaller_font, fill="white", anchor="mm")

    preview_path = f"preview_flyer_{template_id}.png"
    img.save(preview_path)
    return preview_path

# 전단지 PDF 생성
def render_flyer(template_id, items, output_path, title, footer_text):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    margin = 10 * mm
    cell_width = (width - 6 * margin) / 5
    cell_height = 40 * mm

    # 주황색 배경 제목
    c.setFillColorRGB(1, 0.65, 0)
    c.rect(0, height - 15 * mm, width, 15 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("NotoSansKR", 20)
    c.drawCentredString(width / 2, height - 12 * mm, title)

    if template_id == "1":
        x, y = margin, height - 30 * mm
        for i, item in enumerate(items[:5]):
            c.setStrokeColor(colors.black)
            c.rect(x, y - cell_height, cell_width, cell_height)
            if item["ProcessedImagePath"]:
                try:
                    c.drawImage(item["ProcessedImagePath"], x + 5 * mm, y - 15 * mm, 25 * mm, 25 * mm)
                except:
                    c.drawString(x + 5 * mm, y - 15 * mm, "[이미지 없음]")
            c.setFont("Roboto", 12)
            c.drawString(x + 5 * mm, y - 25 * mm, item["Name"][:20])
            c.setFillColorRGB(1, 0, 0)
            c.setFont("NanumGothic", 10)
            c.drawString(x + 5 * mm, y - 30 * mm, f"₩{item['Price']:,.0f} (각)")
            c.setFillColorRGB(0, 0, 0)
            if item.get("AdditionalPrice"):
                c.drawString(x + 5 * mm, y - 35 * mm, f"10g당 {item['AdditionalPrice']}원")
            x += cell_width + margin
            if (i + 1) % 5 == 0:
                x = margin
                y -= cell_height + 5 * mm

    # 하단 주황색 배경 텍스트
    c.setFillColorRGB(1, 0.65, 0)
    c.rect(0, 0, width, 10 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("NanumGothic", 8)
    c.drawCentredString(width / 2, 5 * mm, footer_text)
    c.save()

# Streamlit UI
def main():
    st.markdown('<div class="main-title">전단지 제작 시스템</div>', unsafe_allow_html=True)

    conn = sqlite3.connect("products.db")
    c = conn.cursor()

    # DB 초기화
    c.execute("""CREATE TABLE IF NOT EXISTS Products (
        ProductID INTEGER PRIMARY KEY AUTOINCREMENT,
        Name TEXT NOT NULL,
        Price REAL NOT NULL,
        AdditionalPrice TEXT,
        ImagePath TEXT,
        ProcessedImagePath TEXT,
        LastUpdated DATETIME
    )""")
    conn.commit()

    # 미리보기 섹션 추가
    st.markdown('<div class="section-header">디스플레이 시안 미리보기</div>', unsafe_allow_html=True)
    with st.expander("디스플레이 시안 확인", expanded=True):
        # 임시 품목 데이터 (시안 기반)
        temp_item = {
            "Name": "CJ 명가 재래김/파래김 (20봉, 김 원산지 국산, 각)",
            "Price": 5990,
            "AdditionalPrice": "749",
            "ProcessedImagePath": None,
            "PriceQualityText": True
        }
        title = "가겨운 가구로 품질은 제대로 가격역주행"
        footer_text = "행사기간: 5/20(화) - 기획상품 재고소진시종료"
        preview_path = create_display_preview("1", temp_item, title, footer_text, show_price_quality=True)
        st.image(preview_path, caption="디스플레이 시안 미리보기")

    # 품목 동기화 섹션
    with st.expander("품목 동기화 (포스기 API)", expanded=True):
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown('<div class="section-header">데이터 동기화</div>', unsafe_allow_html=True)
        with col2:
            if st.button("API 데이터 가져오기", key="sync_button"):
                try:
                    items = fetch_pos_data()
                    if not items:
                        st.warning("엑셀 파일에서 데이터를 읽지 못했습니다.")
                        return
                    for item in items:
                        image_path = get_local_image(item["Name"])
                        processed_path = process_image(image_path, item["Name"])
                        if not processed_path:
                            st.warning(f"{item['Name']}의 이미지를 처리하지 못했습니다. image_path: {image_path}")
                        c.execute("INSERT OR REPLACE INTO Products (Name, Price, AdditionalPrice, ImagePath, ProcessedImagePath) VALUES (?, ?, ?, ?, ?)",
                                  (item["Name"], item["Price"], item.get("AdditionalPrice", ""), image_path, processed_path))
                    conn.commit()
                    st.success("데이터 동기화 완료")
                except Exception as e:
                    st.error(f"데이터 동기화 오류: {str(e)}")

    # 품목 목록 및 선택
    with st.expander("품목 목록", expanded=True):
        st.markdown('<div class="section-header">등록된 품목</div>', unsafe_allow_html=True)
        c.execute("SELECT * FROM Products")
        items = [{"Name": row[1], "Price": row[2], "AdditionalPrice": row[3], "ProcessedImagePath": row[4]} for row in c.fetchall()]
        item_names = [item["Name"] for item in items]
        selected_items = st.multiselect("출력할 품목 선택", item_names, default=item_names[:1])  # 단일 품목만 선택 가능

    # 출력 모델 및 템플릿 선택
    with st.expander("출력 설정", expanded=True):
        st.markdown('<div class="section-header">출력 모델 및 템플릿</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            model = st.selectbox("출력 모델 선택", ["디스플레이 (1개 품목)", "전단지 (여러 품목)"])
        with col2:
            template = st.selectbox("템플릿 선택", ["1번 템플릿"])  # 현재는 1번 템플릿만 지원
        template_id = template.split("번")[0]

    # 제목 및 하단 텍스트 설정
    with st.expander("제목 및 하단 텍스트 설정", expanded=True):
        st.markdown('<div class="section-header">텍스트 설정</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            title = st.text_input("제목 입력", "가겨운 가구로 품질은 제대로 가격역주행")
        with col2:
            footer_text = st.text_input("하단 텍스트 입력", "행사기간: 5/20(화) - 기획상품 재고소진시종료")

    # 출력물 생성
    with st.expander("출력물 미리보기 및 생성", expanded=True):
        if model == "디스플레이 (1개 품목)":
            if selected_items:
                for idx, item in enumerate(selected_items):
                    selected_item_data = next((i for i in items if i["Name"] == item), None)
                    if selected_item_data:
                        show_price_quality = selected_item_data.get("PriceQualityText", False)
                        st.subheader(f"디스플레이 미리보기: {selected_item_data['Name']}")
                        preview_path = create_display_preview(template_id, selected_item_data, title, footer_text, show_price_quality)
                        st.image(preview_path, caption=f"디스플레이 {template_id} 미리보기")
                        col1, col2 = st.columns([1, 1])
                        with col1:
                            if st.button(f"디스플레이 생성: {selected_item_data['Name']}", key=f"display_button_{selected_item_data['Name']}_{idx}"):
                                output_path = f"display_{selected_item_data['Name']}.pdf"
                                render_display(template_id, selected_item_data, output_path, title, footer_text, show_price_quality)
                                with open(output_path, "rb") as f:
                                    st.download_button(f"PDF 다운로드: {selected_item_data['Name']}", f, file_name=output_path, key=f"download_display_{selected_item_data['Name']}_{idx}")
            else:
                st.warning("품목을 선택해주세요.")
        else:  # 전단지 (여러 품목)
            if selected_items:
                st.subheader("전단지 미리보기")
                selected_items_data = [item for item in items if item["Name"] in selected_items]
                preview_path = create_flyer_preview(template_id, selected_items_data, title, footer_text)
                st.image(preview_path, caption=f"전단지 {template_id} 미리보기")
                col1, col2 = st.columns([1, 1])
                with col1:
                    if st.button("전단지 생성", key="flyer_generate"):
                        output_path = f"flyer_{template_id}.pdf"
                        render_flyer(template_id, selected_items_data, output_path, title, footer_text)
                        with open(output_path, "rb") as f:
                            st.download_button("PDF 다운로드", f, file_name=output_path, key="flyer_download")
            else:
                st.warning("품목을 선택해주세요.")

    conn.close()

if __name__ == "__main__":
    main()