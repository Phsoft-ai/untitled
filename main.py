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

# --- 🚀 FIX: CanvasData 모델에 canvas_aspect_ratio 필드 추가 ---
class CanvasData(BaseModel):
    background_image_bytes: Optional[str] = Field(None, alias='backgroundImageBytes')
    canvas_size: Dict[str, float] = Field(..., alias='canvasSize')
    # Flutter에서 보낸 canvasAspectRatio를 받기 위한 필드입니다. Optional로 설정하여 이전 버전 앱과 호환성을 유지합니다.
    canvas_aspect_ratio: Optional[float] = Field(None, alias='canvasAspectRatio')
    text_items: List[TextItem] = Field(..., alias='textItems')
    excel_data: List[Dict[str, str]] = Field(..., alias='excelData')
# --- FIX END ---


app = FastAPI()

def crop_image_to_ratio(img: Image.Image, target_ratio: float) -> Image.Image:
    # 기존과 동일한 함수
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
    # 기존과 동일한 함수
    card_width_inch = page_width_inch / 2
    card_height_inch = page_height_inch / 2
    
    canvas_width_px = canvas_size['width']
    canvas_height_px = canvas_size['height']

    # 픽셀-인치 변환 비율은 canvas_size와 카드 크기를 기준으로 계산합니다.
    # canvas_size의 비율이 바뀌면 이 값도 자동으로 보정됩니다.
    pixels_per_inch_w = canvas_width_px / card_width_inch
    pixels_per_inch_h = canvas_height_px / card_height_inch


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
            # 카드의 실제 크기에 맞춰 배경 이미지를 추가합니다.
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

            # --- 좌표 및 크기 계산 로직 (기존과 동일하게 유지) ---
            # Flutter의 canvasSize가 이미 변경된 비율을 반영하므로, 이 로직은 수정할 필요가 없습니다.
            center_pos_px = item_template.center_position
            cx_px = center_pos_px.dx
            cy_px = center_pos_px.dy

            HORIZONTAL_CORRECTION_FACTOR = 0.9
            VERTICAL_CORRECTION_FACTOR = 0.93
            VERTICAL_OFFSET_PT = 3.5

            corrected_cx_px = cx_px * HORIZONTAL_CORRECTION_FACTOR
            corrected_cy_px = cy_px * VERTICAL_CORRECTION_FACTOR
            
            vertical_offset_px = VERTICAL_OFFSET_PT * (96 / 72)
            
            center_x_abs_px = (canvas_width_px / 2) + corrected_cx_px
            center_y_abs_px = (canvas_height_px / 2) - corrected_cy_px + vertical_offset_px

            font_size_pt = item_template.font_size_pt
            measured_height_pt = item_template.measured_height_pt or font_size_pt
            
            text_box_width_px = canvas_width_px * 0.90
            text_box_height_px = (measured_height_pt * (96 / 72)) * 1.1

            left_px = center_x_abs_px - (text_box_width_px / 2)
            top_px = center_y_abs_px - (text_box_height_px / 2)
            
            # 픽셀-인치 변환 시, 가로/세로 각각의 비율을 사용합니다.
            final_left_inch = left_px / pixels_per_inch_w
            final_top_inch = top_px / pixels_per_inch_h
            box_width_inch = text_box_width_px / pixels_per_inch_w
            box_height_inch = text_box_height_px / pixels_per_inch_h
            
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
    prs = Presentation()
    
    # PPT 슬라이드 크기는 A4 용지로 고정
    prs.slide_width = Inches(8.27)
    prs.slide_height = Inches(11.69)
    
    chunk_size = 4
    data_chunks = [data.excel_data[i:i + chunk_size] for i in range(0, len(data.excel_data), chunk_size)]

    page_width_inch = prs.slide_width.inches
    page_height_inch = prs.slide_height.inches
    
    # --- 🚀 FIX: 배경 이미지를 자를 때 사용할 비율을 결정 ---
    # 1. 기본값은 A4 용지의 1/4 비율
    card_width_inch = page_width_inch / 2
    card_height_inch = page_height_inch / 2
    target_ratio = card_width_inch / card_height_inch
    
    # 2. 만약 Flutter에서 canvasAspectRatio 값을 보냈다면, 그 값을 우선적으로 사용
    if data.canvas_aspect_ratio is not None and data.canvas_aspect_ratio > 0:
        target_ratio = data.canvas_aspect_ratio
        # Flutter에서 보낸 비율에 맞춰 카드의 너비 또는 높이를 재계산합니다.
        # 이렇게 해야 배경 이미지가 카드에 정확히 맞춰집니다.
        if target_ratio > (card_width_inch / card_height_inch):
             # 새 비율이 기준보다 넓으면, 높이를 줄입니다.
             card_height_inch = card_width_inch / target_ratio
        else:
            # 새 비율이 기준보다 좁으면, 너비를 줄입니다.
            card_width_inch = card_height_inch * target_ratio

    # --- FIX END ---

    cropped_background_stream = None
    if data.background_image_bytes:
        try:
            img_bytes = base64.b64decode(data.background_image_bytes)
            original_image = Image.open(io.BytesIO(img_bytes))
            # 위에서 결정된 target_ratio를 사용하여 이미지를 자릅니다.
            cropped_image = crop_image_to_ratio(original_image, target_ratio)
            cropped_background_stream = io.BytesIO()
            cropped_image.save(cropped_background_stream, format=original_image.format or 'PNG')
            cropped_background_stream.seek(0)
        except Exception as e:
            print(f"Error processing background image: {e}")
            
    for chunk in data_chunks:
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
        
        # add_cards_on_slide 함수에 재계산된 카드 크기를 전달합니다.
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
