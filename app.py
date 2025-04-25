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

# 나눔고딕 폰트 등록
pdfmetrics.registerFont(TTFont("NanumGothic", "NanumGothic.ttf"))
font_path = "NanumGothic.ttf"

# 엑셀 파일에서 데이터 읽기 (열 이름 디버깅 추가)
def fetch_pos_data():
    try:
        df = pd.read_excel("items.xlsx")
        expected_columns = ["Name", "Price"]
        optional_columns = ["AdditionalPrice"]
        missing_columns = [col for col in expected_columns if col not in df.columns]
        if missing_columns:
            st.error(f"엑셀 파일에 다음 필수 열이 누락되었습니다: {missing_columns}")
            return []
        # 열 이름 확인 및 디버깅 출력
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
        "세제 1L": "detergent.png"
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

# 디스플레이 미리보기 생성 (A4 1/3 크기, 1개 품목)
def create_display_preview(template_id, item, title, footer_text):
    width, height = 595, 280  # A4 1/3 크기 (72dpi 기준, 210mm x 99mm)
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(font_path, 20)
    small_font = ImageFont.truetype(font_path, 16)
    smaller_font = ImageFont.truetype(font_path, 12)

    # 주황색 배경 제목
    draw.rectangle((0, 0, width, 40), fill=(255, 165, 0))
    draw.text((width / 2, 20), title, font=font, fill="white", anchor="mm")

    # 품목 표시
    margin = 10
    x, y = margin, 50
    draw.rectangle((x, y, x + 80, y + 80), outline="black")  # 정사각형 이미지 영역
    if item["ProcessedImagePath"]:
        try:
            item_img = Image.open(item["ProcessedImagePath"])
            item_img.thumbnail((80, 80))
            img.paste(item_img, (x, y), item_img if item_img.mode == "RGBA" else None)
        except:
            draw.text((x + 10, y + 10), "[이미지 없음]", font=smaller_font, fill="black")
    draw.text((x + 90, y + 20), f"₩{item['Price']:,.0f} (각)", font=small_font, fill="red")
    if item.get("AdditionalPrice"):
        draw.text((x + 90, y + 40), f"10g당 {item['AdditionalPrice']}원", font=smaller_font, fill="black")

    # 하단 주황색 배경 텍스트
    draw.rectangle((0, height - 20, width, height), fill=(255, 165, 0))
    draw.text((width / 2, height - 10), footer_text, font=smaller_font, fill="white", anchor="mm")

    preview_path = f"preview_display_{template_id}.png"
    img.save(preview_path)
    return preview_path

# 전단지 미리보기 생성 (여러 품목)
def create_flyer_preview(template_id, items, title, footer_text):
    width, height = 595, 842  # A4 크기 (72dpi)
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(font_path, 20)
    small_font = ImageFont.truetype(font_path, 12)
    smaller_font = ImageFont.truetype(font_path, 10)

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

# 디스플레이 PDF 생성
def render_display(template_id, item, output_path, title, footer_text):
    c = canvas.Canvas(output_path, pagesize=(210 * mm, 99 * mm))  # A4 1/3 크기
    width, height = 210 * mm, 99 * mm
    margin = 5 * mm

    # 주황색 배경 제목
    c.setFillColorRGB(1, 0.65, 0)
    c.rect(0, height - 10 * mm, width, 10 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("NanumGothic", 20)
    c.drawCentredString(width / 2, height - 7 * mm, title)

    # 품목 표시
    x, y = margin, height - 20 * mm
    c.setStrokeColor(colors.black)
    c.rect(x, y - 20 * mm, 20 * mm, 20 * mm)
    if item["ProcessedImagePath"]:
        try:
            c.drawImage(item["ProcessedImagePath"], x + 2 * mm, y - 18 * mm, 16 * mm, 16 * mm)
        except:
            c.drawString(x + 2 * mm, y - 18 * mm, "[이미지 없음]")
    c.setFont("NanumGothic", 16)
    c.drawString(x + 25 * mm, y - 5 * mm, f"₩{item['Price']:,.0f} (각)")
    c.setFillColorRGB(1, 0, 0)
    c.setFont("NanumGothic", 12)
    if item.get("AdditionalPrice"):
        c.drawString(x + 25 * mm, y - 15 * mm, f"10g당 {item['AdditionalPrice']}원")
    c.setFillColorRGB(0, 0, 0)

    # 하단 주황색 배경 텍스트
    c.setFillColorRGB(1, 0.65, 0)
    c.rect(0, 0, width, 5 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    beep = lambda x: ord(x) if len(x) == 1 else sum(ord(c) for c in x)
    c.setFont("NanumGothic", 12)
    c.drawCentredString(width / 2, 2.5 * mm, footer_text)
    c.save()

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
    c.setFont("NanumGothic", 20)
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
            c.setFont("NanumGothic", 12)
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
    st.title("전단지 제작 시스템")

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

    # 품목 동기화
    st.header("품목 동기화 (포스기 API)")
    if st.button("API 데이터 가져오기"):
        try:
            items = fetch_pos_data()
            if not items:
                st.warning("엑셀 파일에서 데이터를 읽지 못했습니다.")
                return
            for item in items:
                image_path = get_local_image(item["Name"])
                processed_path = process_image(image_path, item["Name"])
                c.execute("INSERT OR REPLACE INTO Products (Name, Price, AdditionalPrice, ImagePath, ProcessedImagePath) VALUES (?, ?, ?, ?, ?)",
                          (item["Name"], item["Price"], item.get("AdditionalPrice", ""), image_path, processed_path))
            conn.commit()
            st.success("데이터 동기화 완료")
        except Exception as e:
            st.error(f"데이터 동기화 오류: {str(e)}")

    # 품목 목록 및 선택
    st.header("품목 목록")
    c.execute("SELECT * FROM Products")
    items = [{"Name": row[1], "Price": row[2], "AdditionalPrice": row[3], "ProcessedImagePath": row[4]} for row in c.fetchall()]
    item_names = [item["Name"] for item in items]
    selected_items = st.multiselect("출력할 품목 선택", item_names, default=item_names)

    # 출력 모델 선택
    model = st.selectbox("출력 모델 선택", ["디스플레이 (1개 품목)", "전단지 (여러 품목)"])

    # 템플릿 선택
    template = st.selectbox("템플릿 선택", [
        "1번 템플릿",
        "2번 템플릿",
        "3번 템플릿",
        "4번 템플릿",
        "5번 템플릿"
    ])
    template_id = template.split("번")[0]

    # 제목 및 하단 텍스트 설정
    title = st.text_input("제목 입력", "가겨운 가구로 품질은 제대로 가격역주행")
    footer_text = st.text_input("하단 텍스트 입력", "행사기간: 5/20(화) - 기획상품 재고소진시종료")

    # 선택된 품목 필터링
    selected_items_data = [item for item in items if item["Name"] in selected_items]

    if model == "디스플레이 (1개 품목)":
        if selected_items_data:
            for item in selected_items_data:
                st.subheader(f"디스플레이 미리보기: {item['Name']}")
                preview_path = create_display_preview(template_id, item, title, footer_text)
                st.image(preview_path, caption=f"디스플레이 {template_id} 미리보기")
                if st.button(f"디스플레이 생성: {item['Name']}", key=f"display_button_{item['Name']}"):
                    output_path = f"display_{item['Name']}.pdf"
                    render_display(template_id, item, output_path, title, footer_text)
                    with open(output_path, "rb") as f:
                        st.download_button(f"PDF 다운로드: {item['Name']}", f, file_name=output_path, key=f"download_button_{item['Name']}")
        else:
            st.warning("품목을 선택해주세요.")

    else:  # 전단지 (여러 품목)
        if selected_items_data:
            st.subheader("전단지 미리보기")
            preview_path = create_flyer_preview(template_id, selected_items_data, title, footer_text)
            st.image(preview_path, caption=f"전단지 {template_id} 미리보기")
            if st.button("전단지 생성"):
                output_path = f"flyer_{template_id}.pdf"
                render_flyer(template_id, selected_items_data, output_path, title, footer_text)
                with open(output_path, "rb") as f:
                    st.download_button("PDF 다운로드", f, file_name=output_path)
        else:
            st.warning("품목을 선택해주세요.")

    conn.close()

if __name__ == "__main__":
    main()