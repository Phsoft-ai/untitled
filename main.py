# main.py

import base64
import io
import unicodedata
from fastapi import FastAPI, Response
from pydantic import BaseModel, Field
from typing import List, Optional, Dict

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from PIL import Image

# --- Pydantic 모델 정의 (기존과 동일) ---
class CenterPosition(BaseModel):
    dx: float
    dy: float

class TextItem(BaseModel):
    id: str
    text: str
    center_position: CenterPosition = Field(..., alias='centerPosition')
    font_size_pt: float = Field(..., alias='fontSizePt')
    measured_height_pt: Optional[float] = Field(None, alias='measuredHeightPt')
    color_value: int = Field(..., alias='colorValue')
    font_weight_bold: bool = Field(..., alias='fontWeightBold')
    font_family: Optional[str] = Field(None, alias='fontFamily')

class CanvasData(BaseModel):
    background_image_bytes: Optional[str] = Field(None, alias='backgroundImageBytes')
    canvas_size: Dict[str, float] = Field(..., alias='canvasSize')
    text_items: List[TextItem] = Field(..., alias='textItems')
    excel_data: List[Dict[str, str]] = Field(..., alias='excelData')

app = FastAPI()

def crop_image_to_ratio(img: Image.Image, target_ratio: float) -> Image.Image:
    # ... (기존과 동일)
    img_width, img_height = img.size
    img_ratio = img_width / img_height

    if abs(img_ratio - target_ratio) < 0.01:
        return img

    if img_ratio > target_ratio:
        new_width = int(target_ratio * img_height)
        offset = (img_width - new_width) // 2
        return img.crop((offset, 0, offset + new_width, img_height))
    else:
        new_height = int(img_width / target_ratio)
        offset = (img_height - new_height) // 2
        return img.crop((0, offset, img_width, offset + new_height))


def add_cards_on_slide(slide, chunk_data, text_items_template, canvas_size, cropped_background_stream, page_width_inch, page_height_inch):
    card_width_inch = page_width_inch / 2
    card_height_inch = page_height_inch / 2
    
    canvas_width_px = canvas_size['width']
    canvas_height_px = canvas_size['height']

    pixels_per_inch = canvas_width_px / card_width_inch

    grid_positions = [
        (0, 0),
        (card_width_inch, 0),
        (0, card_height_inch),
        (card_width_inch, card_height_inch)
    ]

    for i, card_data in enumerate(chunk_data):
        base_left_inch, base_top_inch = grid_positions[i]

        if cropped_background_stream:
            cropped_background_stream.seek(0)
            slide.shapes.add_picture(
                cropped_background_stream,
                Inches(base_left_inch),
                Inches(base_top_inch),
                width=Inches(card_width_inch),
                height=Inches(card_height_inch)
            )

        for item_template in text_items_template:
            if item_template.id == 'title':
                text_content = card_data.get('name', '')
            elif item_template.id == 'subtitle':
                text_content = card_data.get('group', '')
            else:
                text_content = item_template.text

            # --- 좌표 및 크기 계산 로직 (수정된 부분) ---
            center_pos_px = item_template.center_position
            cx_px = center_pos_px.dx
            cy_px = center_pos_px.dy

            # [수정 1] 좌표 보정 계수 및 전역 오프셋 정의
            HORIZONTAL_CORRECTION_FACTOR = 0.9
            VERTICAL_CORRECTION_FACTOR = 0.93
            
            # [수정 2] 전역 수직 오프셋 (단위: 포인트)
            # 모든 텍스트 요소를 아래로 이동시키려면 양수 값을, 위로 이동시키려면 음수 값을 입력하세요.
            # 이 값을 조절하여 전체 텍스트의 수직 위치를 한 번에 옮길 수 있습니다.
            VERTICAL_OFFSET_PT = 3.5 # 예: 10포인트만큼 아래로 이동

            # 보정 계수를 적용하여 좌표를 재계산
            corrected_cx_px = cx_px * HORIZONTAL_CORRECTION_FACTOR
            corrected_cy_px = cy_px * VERTICAL_CORRECTION_FACTOR

            # 오프셋을 픽셀 단위로 변환 (1 포인트 = 96/72 픽셀)
            vertical_offset_px = VERTICAL_OFFSET_PT * (96 / 72)

            # 1. Flutter의 중앙 기준 좌표를 카드 좌상단 기준의 절대 좌표(픽셀)로 변환
            # 이때, 위에서 계산한 '보정된' 좌표를 사용하고, 전역 오프셋을 더해줍니다.
            center_x_abs_px = (canvas_width_px / 2) + corrected_cx_px
            center_y_abs_px = (canvas_height_px / 2) - corrected_cy_px + vertical_offset_px # <<< 여기 오프셋이 추가되었습니다.

            font_size_pt = item_template.font_size_pt
            measured_height_pt = item_template.measured_height_pt or font_size_pt
            
            # 2. 텍스트 박스의 크기(픽셀) 결정
            text_box_width_px = canvas_width_px * 0.90
            text_box_height_px = (measured_height_pt * (96 / 72)) * 1.1

            # 3. 텍스트 박스의 좌측 상단 좌표(픽셀) 계산
            left_px = center_x_abs_px - (text_box_width_px / 2)
            top_px = center_y_abs_px - (text_box_height_px / 2)

            # 4. 계산된 픽셀 값들을 인치 단위로 변환
            final_left_inch = left_px / pixels_per_inch
            final_top_inch = top_px / pixels_per_inch
            box_width_inch = text_box_width_px / pixels_per_inch
            box_height_inch = text_box_height_px / pixels_per_inch
            
            # --- 이후 텍스트 박스 추가 및 스타일링 로직은 기존과 동일 ---
            txBox = slide.shapes.add_textbox(
                Inches(base_left_inch + final_left_inch),
                Inches(base_top_inch + final_top_inch),
                Inches(box_width_inch),
                Inches(box_height_inch)
            )
            
            tf = txBox.text_frame
            tf.margin_bottom = Pt(0)
            tf.margin_top = Pt(0)
            tf.margin_left = Pt(0)
            tf.margin_right = Pt(0)
            tf.word_wrap = True
            tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            tf.clear()

            p = tf.paragraphs[0]
            p.text = text_content
            p.alignment = PP_ALIGN.CENTER

            font = p.font
            default_font = 'Malgun Gothic'

            if item_template.font_family:
                try:
                    font_name = unicodedata.normalize('NFC', item_template.font_family)
                    font.name = font_name
                except Exception as e:
                    print(f"Warning: Could not set font '{item_template.font_family}'. Error: {e}. Falling back to '{default_font}'.")
                    font.name = default_font
            else:
                font.name = default_font

            font.size = Pt(item_template.font_size_pt)
            font.bold = item_template.font_weight_bold

            color_val = item_template.color_value
            font.color.rgb = RGBColor((color_val >> 16) & 0xFF, (color_val >> 8) & 0xFF, color_val & 0xFF)


@app.post("/generate-ppt")
async def generate_ppt(data: CanvasData):
    # ... (이후 코드는 기존과 동일)
    prs = Presentation()
    
    prs.slide_width = Inches(8.27)
    prs.slide_height = Inches(11.69)
    
    chunk_size = 4
    data_chunks = [data.excel_data[i:i + chunk_size] for i in range(0, len(data.excel_data), chunk_size)]

    page_width_inch = prs.slide_width.inches
    page_height_inch = prs.slide_height.inches
    card_width_inch = page_width_inch / 2
    card_height_inch = page_height_inch / 2
    card_ratio = card_width_inch / card_height_inch

    cropped_background_stream = None
    if data.background_image_bytes:
        try:
            img_bytes = base64.b64decode(data.background_image_bytes)
            original_image = Image.open(io.BytesIO(img_bytes))
            cropped_image = crop_image_to_ratio(original_image, card_ratio)
            cropped_background_stream = io.BytesIO()
            cropped_image.save(cropped_background_stream, format=original_image.format or 'PNG')
            cropped_background_stream.seek(0)
        except Exception as e:
            print(f"Error processing background image: {e}")
            
    for chunk in data_chunks:
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
        
        add_cards_on_slide(
            slide=slide,
            chunk_data=chunk,
            text_items_template=data.text_items,
            canvas_size=data.canvas_size,
            cropped_background_stream=cropped_background_stream,
            page_width_inch=page_width_inch,
            page_height_inch=page_height_inch,
        )

    file_stream = io.BytesIO()
    prs.save(file_stream)
    file_stream.seek(0)

    return Response(
        content=file_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": "attachment; filename=generated_A4.pptx"}
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
