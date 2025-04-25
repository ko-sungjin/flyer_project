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

# 엑셀 파일에서 데이터 읽기
def fetch_pos_data():
    df = pd.read_excel("items.xlsx")
    items = df.to_dict("records")
    return items

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
    if filename and os.path.exists(f"C:/flyer_project/images/{filename}"):
        return f"C:/flyer_project/images/{filename}"
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

# 미리보기 이미지 생성
def create_preview(template_id, items):
    width, height = 595, 842  # A4 at 72dpi
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(font_path, 20)
    small_font = ImageFont.truetype(font_path, 12)
    smaller_font = ImageFont.truetype(font_path, 10)

    if template_id == "1":  # 4x2 그리드
        draw.text((width / 2, 30), "마트 세일 전단지", font=font, fill="black", anchor="mm")
        cell_width, cell_height = 140, 100
        margin = 30
        x, y = margin, 80
        for i, item in enumerate(items[:8]):
            draw.rectangle((x, y, x + cell_width, y + cell_height), outline="blue")
            if item["ProcessedImagePath"]:
                try:
                    item_img = Image.open(item["ProcessedImagePath"])
                    item_img.thumbnail((80, 80))
                    img.paste(item_img, (x + 10, y + 10), item_img if item_img.mode == "RGBA" else None)
                except:
                    draw.text((x + 10, y + 10), "[이미지 없음]", font=smaller_font, fill="black")
            draw.text((x + 10, y + 70), item["Name"][:20], font=small_font, fill="black")
            draw.text((x + 10, y + 85), f"₩{item['Price']:,.0f}", font=smaller_font, fill="black")
            x += cell_width
            if (i + 1) % 4 == 0:
                x = margin
                y += cell_height

    elif template_id == "2":  # 3x3 그리드
        draw.text((width / 2, 30), "특별 할인", font=font, fill="black", anchor="mm")
        cell_width, cell_height = 180, 80
        margin = 30
        x, y = margin, 80
        for i, item in enumerate(items[:9]):
            draw.rectangle((x, y, x + cell_width, y + cell_height), outline="black", fill=(240, 240, 240))
            if item["ProcessedImagePath"]:
                try:
                    item_img = Image.open(item["ProcessedImagePath"])
                    item_img.thumbnail((60, 60))
                    img.paste(item_img, (x + 10, y + 10), item_img if item_img.mode == "RGBA" else None)
                except:
                    draw.text((x + 10, y + 10), "[이미지 없음]", font=smaller_font, fill="black")
            draw.text((x + 10, y + 50), item["Name"][:15], font=smaller_font, fill="black")
            draw.text((x + 10, y + 65), f"₩{item['Price']:,.0f}", font=smaller_font, fill="black")
            x += cell_width
            if (i + 1) % 3 == 0:
                x = margin
                y += cell_height

    elif template_id == "3":  # 세로 리스트
        draw.text((width / 2, 30), "오늘의 할인 품목", font=font, fill="black", anchor="mm")
        margin = 30
        y = 80
        for item in items[:10]:
            if item["ProcessedImagePath"]:
                try:
                    item_img = Image.open(item["ProcessedImagePath"])
                    item_img.thumbnail((50, 50))
                    img.paste(item_img, (margin, y), item_img if item_img.mode == "RGBA" else None)
                except:
                    draw.text((margin, y), "[이미지 없음]", font=smaller_font, fill="black")
            draw.text((margin + 60, y + 10), item["Name"][:30], font=small_font, fill="black")
            draw.text((margin + 60, y + 25), f"₩{item['Price']:,.0f}", font=smaller_font, fill="black")
            draw.line((margin, y + 50, width - margin, y + 50), fill="black")
            y += 60

    preview_path = f"preview_template_{template_id}.png"
    img.save(preview_path)
    return preview_path

# ReportLab 템플릿 (PDF 생성)
def render_template_1(items, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    margin = 10 * mm
    cell_width = (width - 2 * margin) / 4
    cell_height = (height - 3 * margin) / 4

    c.setFont("NanumGothic", 20)
    c.drawCentredString(width / 2, height - 20 * mm, "마트 세일 전단지")

    x, y = margin, height - 40 * mm
    for i, item in enumerate(items):
        c.setStrokeColor(colors.blue)
        c.rect(x, y - cell_height, cell_width, cell_height)
        if item["ProcessedImagePath"]:
            try:
                c.drawImage(item["ProcessedImagePath"], x + 10, y - 30 * mm, 30 * mm, 30 * mm)
            except:
                c.drawString(x + 10, y - 30 * mm, "[이미지 없음]")
        c.setFont("NanumGothic", 12)
        c.drawString(x + 10, y - 40 * mm, item["Name"][:20])
        c.setFont("NanumGothic", 10)
        c.drawString(x + 10, y - 45 * mm, f"₩{item['Price']:,.0f}")
        x += cell_width
        if (i + 1) % 4 == 0:
            x = margin
            y -= cell_height
        if y < margin + cell_height:
            c.showPage()
            y = height - 40 * mm
    c.save()

def render_template_2(items, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    margin = 10 * mm
    cell_width = (width - 2 * margin) / 3
    cell_height = (height - 3 * margin) / 3

    c.setFont("NanumGothic", 18)
    c.drawCentredString(width / 2, height - 20 * mm, "특별 할인")

    x, y = margin, height - 40 * mm
    for i, item in enumerate(items):
        c.setStrokeColor(colors.black)
        c.rect(x, y - cell_height, cell_width, cell_height, stroke=1, fill=0)
        c.setFillColor(colors.lightgrey)
        c.rect(x, y - cell_height, cell_width, cell_height, stroke=0, fill=1)
        if item["ProcessedImagePath"]:
            try:
                c.drawImage(item["ProcessedImagePath"], x + 10, y - 25 * mm, 25 * mm, 25 * mm)
            except:
                c.drawString(x + 10, y - 25 * mm, "[이미지 없음]")
        c.setFont("NanumGothic", 10)
        c.drawString(x + 10, y - 35 * mm, item["Name"][:15])
        c.setFont("NanumGothic", 8)
        c.drawString(x + 10, y - 40 * mm, f"₩{item['Price']:,.0f}")
        x += cell_width
        if (i + 1) % 3 == 0:
            x = margin
            y -= cell_height
        if y < margin + cell_height:
            c.showPage()
            y = height - 40 * mm
    c.save()

def render_template_3(items, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    margin = 10 * mm
    item_height = 30 * mm

    c.setFont("NanumGothic", 20)
    c.drawCentredString(width / 2, height - 20 * mm, "오늘의 할인 품목")

    y = height - 40 * mm
    for item in items[:10]:
        if item["ProcessedImagePath"]:
            try:
                c.drawImage(item["ProcessedImagePath"], margin, y - 20 * mm, 20 * mm, 20 * mm)
            except:
                c.drawString(margin, y - 20 * mm, "[이미지 없음]")
        c.setFont("NanumGothic", 12)
        c.drawString(margin + 25 * mm, y - 15 * mm, item["Name"][:30])
        c.setFont("NanumGothic", 10)
        c.drawString(margin + 25 * mm, y - 20 * mm, f"₩{item['Price']:,.0f}")
        c.line(margin, y - 25 * mm, width - margin, y - 25 * mm)
        y -= item_height
        if y < margin:
            c.showPage()
            y = height - 40 * mm
    c.save()

def create_flyer(template_id, items, output_path="flyer.pdf"):
    templates = {
        "1": render_template_1,
        "2": render_template_2,
        "3": render_template_3
    }
    templates[template_id](items, output_path)

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
        ImagePath TEXT,
        ProcessedImagePath TEXT,
        LastUpdated DATETIME
    )""")
    conn.commit()

    # 품목 목록
    st.header("품목 목록")
    c.execute("SELECT * FROM Products")
    items = [{"Name": row[1], "Price": row[2], "ProcessedImagePath": row[4]} for row in c.fetchall()]
    for item in items:
        st.write(f"{item['Name']} - ₩{item['Price']:,.0f}")
        if item["ProcessedImagePath"]:
            st.image(item["ProcessedImagePath"], width=100)

    # 품목 동기화
    st.header("품목 동기화 (포스기 API)")
    if st.button("API 데이터 가져오기"):
        items = fetch_pos_data()
        for item in items:
            image_path = get_local_image(item["Name"])
            processed_path = process_image(image_path, item["Name"])
            c.execute("INSERT OR REPLACE INTO Products (Name, Price, ImagePath, ProcessedImagePath) VALUES (?, ?, ?, ?)",
                      (item["Name"], item["Price"], image_path, processed_path))
        conn.commit()
        st.success("데이터 동기화 완료")

    # 전단지 생성 및 미리보기
    st.header("전단지 생성")
    template = st.selectbox("템플릿 선택", [
        "1: 4x2 그리드 (마트 세일)",
        "2: 3x3 그리드 (특별 할인)",
        "3: 세로 리스트 (오늘의 할인)"
    ])
    template_id = template.split(":")[0]

    if st.button("미리보기"):
        c.execute("SELECT Name, Price, ProcessedImagePath FROM Products")
        items = [{"Name": row[0], "Price": row[1], "ProcessedImagePath": row[2]} for row in c.fetchall()]
        preview_path = create_preview(template_id, items)
        st.image(preview_path, caption=f"템플릿 {template_id} 미리보기")

    if st.button("전단지 생성"):
        c.execute("SELECT Name, Price, ProcessedImagePath FROM Products")
        items = [{"Name": row[0], "Price": row[1], "ProcessedImagePath": row[2]} for row in c.fetchall()]
        create_flyer(template_id, items, "flyer.pdf")
        st.success("전단지 생성 완료: flyer.pdf")
        with open("flyer.pdf", "rb") as f:
            st.download_button("PDF 다운로드", f, file_name="flyer.pdf")

    conn.close()

if __name__ == "__main__":
    main()