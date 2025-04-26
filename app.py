import pandas as pd
import os
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from PIL import Image

# 엑셀 파일에서 데이터 읽기
def load_excel_data(file_path):
    df = pd.read_excel(file_path)
    return df.to_dict(orient="records")

# 로컬 이미지 불러오기
def get_local_image(item):
    image_file = item.get("ImageFile")
    if image_file and os.path.exists(f"images/{image_file}"):
        return f"images/{image_file}"
    return "images/no_image.png"  # 기본 이미지

# 디스플레이 PDF 생성
def render_display(item, output_path, title, footer_text):
    c = canvas.Canvas(output_path, pagesize=(210 * mm, 99 * mm))  # PDF 크기 설정
    width, height = 210 * mm, 99 * mm
    margin = 5 * mm

    # 전체 테두리 추가
    c.setStrokeColorRGB(0, 0, 0)
    c.rect(0, 0, width, height, stroke=1, fill=0)

    # 상단 문구 (주황색 배경, 흰 글씨)
    c.setFillColorRGB(1, 0.65, 0)  # 주황색 배경
    c.rect(0, height - 10 * mm, width, 10 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)  # 흰색 글씨
    c.setFont("NanumGothic", 20)
    c.drawCentredString(width / 2, height - 7 * mm, title)

    # 상품명 (엑셀에서 가져옴)
    x, y = margin, height - 20 * mm
    c.setFont("NanumGothic", 16)
    c.setFillColorRGB(0, 0, 0)  # 검은색 글씨
    c.drawString(x, y, item["Name"])

    # 이미지 (엑셀에서 매핑된 이미지, 지정된 위치와 크기)
    image_path = get_local_image(item)
    try:
        c.drawImage(image_path, width - 25 * mm, height - 25 * mm, 20 * mm, 20 * mm)
    except:
        c.drawString(width - 25 * mm, height - 25 * mm, "[이미지 없음]")

    # 하단 문구 (주황색 배경, 흰 글씨)
    c.setFillColorRGB(1, 0.65, 0)  # 주황색 배경
    c.rect(0, 0, width, 5 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)  # 흰색 글씨
    c.setFont("NanumGothic", 12)
    c.drawCentredString(width / 2, 2.5 * mm, footer_text)

    c.save()

# 메인 실행 함수
def generate_displays(excel_file, output_dir, title, footer_text):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    items = load_excel_data(excel_file)
    for i, item in enumerate(items):
        output_path = f"{output_dir}/display_{i+1}.pdf"
        render_display(item, output_path, title, footer_text)

# 실행 예시
if __name__ == "__main__":
    excel_file = "items.xlsx"  # 엑셀 파일 경로
    output_dir = "output_displays"  # 출력 폴더
    title = "가겨운 가구로 품질은 제대로 가격역주행"  # 상단 문구
    footer_text = "행사기간: 5/20(화) - 기획상품 재고소진시종료"  # 하단 문구
    
    generate_displays(excel_file, output_dir, title, footer_text)