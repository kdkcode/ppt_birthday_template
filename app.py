import streamlit as st
import pandas as pd
import math
import os
import uuid
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image, ImageDraw, UnidentifiedImageError

############################
# 월별 테마 (1~12 예시)
############################
MONTHLY_THEMES = {
    1:  {'background': 'FFE6E6', 'title_color': 'FF3366', 'message': '부럽다 부러워 1월의 해피버쓰?!데이', 'sub_message': '생일자 여러분 생일 당일 오후 12시 30분 퇴근하세요'},
    2:  {'background': 'E6EEFF', 'title_color': '3366FF', 'message': '2월의 러블리 벌스데이',       'sub_message': '2월 생일자 여러분 축하합니다!'},
    3:  {'background': 'E6FFE6', 'title_color': '33CC33', 'message': '3월의 새싹같은 생일',       'sub_message': '봄날에 태어난 여러분!'},
    4:  {'background': 'FFFBE6', 'title_color': 'CC9900', 'message': '4월의 화사한 생일!',       'sub_message': '꽃처럼 피어난 4월의 주인공'},
    5:  {'background': 'F0F0F0', 'title_color': '333333', 'message': '5월의 가정의 달 탄생!',     'sub_message': '가족 같은 회사에서 함께해요'},
    6:  {'background': 'E6FFFF', 'title_color': '00CCCC', 'message': '6월의 시원한 생일이!',     'sub_message': '여름을 시원하게 만들어줄 당신'},
    7:  {'background': 'FFF0F5', 'title_color': 'FF0066', 'message': '7월의 뜨거운 생일',         'sub_message': '열정 가득 7월의 주인공'},
    8:  {'background': 'F5F5DC', 'title_color': '996600', 'message': '8월의 태양처럼! 생일',      'sub_message': '한여름 태양보다 뜨거운 축하'},
    9:  {'background': 'F0FFF0', 'title_color': '009966', 'message': '9월의 풍성한 생일',        'sub_message': '가을처럼 풍요로운 9월'},
    10: {'background': 'FFFACD', 'title_color': 'CC6600', 'message': '10월의 청명한 생일',       'sub_message': '맑고 높은 하늘처럼 빛나는'},
    11: {'background': 'F5F5F5', 'title_color': '666666', 'message': '11월의 감사하는 생일',     'sub_message': '가을의 끝, 감사의 마음과 함께'},
    12: {'background': 'FFEFFC', 'title_color': 'FF33CC', 'message': '12월의 따뜻한 생일',       'sub_message': '연말에 더 반짝이는 당신!'}
}

############################
# PPT 생성 클래스
############################
class BirthdaySlideGenerator:
    def __init__(self, month=1):
        # 새 PPT 시작
        self.prs = Presentation()
        
        # 슬라이드 크기(EMU)
        self.prs.slide_width = int(33.87 * 360000)  
        self.prs.slide_height = int(19.05 * 360000)
        
        # 빈 슬라이드 추가
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        self.SLIDE_WIDTH = self.prs.slide_width
        self.SLIDE_HEIGHT = self.prs.slide_height
        
        # 월별 테마 불러오기
        self.month = month
        if month in MONTHLY_THEMES:
            self.theme = MONTHLY_THEMES[month]
        else:
            self.theme = {
                'background': 'FFFFFF',
                'title_color': '000000',
                'message': f'{month}월의 생일자',
                'sub_message': '축하합니다!'
            }
    
    def create_circle_image(self, image_path, output_path, size=(300, 300)):
        """원형 마스크로 씌워 PNG 생성 (투명 배경)"""
        img = Image.open(image_path).convert("RGBA")
        img = img.resize(size, Image.Resampling.LANCZOS)
        
        mask = Image.new('L', size, 0)
        draw = ImageDraw.Draw(mask)
        draw.ellipse((0, 0, size[0], size[1]), fill=255)
        
        output = Image.new('RGBA', size, (0, 0, 0, 0))
        output.paste(img, (0,0), mask=mask)
        output.save(output_path, 'PNG')
        
        return output_path

    def add_decorations(self):
        """좌우 상단 간단 장식"""
        slide_width_in = self.SLIDE_WIDTH / 914400.0
        shapes = [MSO_SHAPE.PENTAGON, MSO_SHAPE.OVAL, MSO_SHAPE.DIAMOND]
        colors = ['FFD700', 'FF69B4', 'FF6B6B']  
        
        # 왼쪽
        for i in range(3):
            shape = self.slide.shapes.add_shape(
                shapes[i % len(shapes)],
                Inches(0.5 + i * 1.5),
                Inches(0.3),
                Inches(0.4),
                Inches(0.4)
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(colors[i % len(colors)])
            shape.line.width = Pt(0)
        
        # 오른쪽
        for i in range(3):
            x_pos = slide_width_in - (0.5 + i * 1.5)
            shape = self.slide.shapes.add_shape(
                shapes[i % len(shapes)],
                Inches(x_pos),
                Inches(0.3),
                Inches(0.4),
                Inches(0.4)
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(colors[i % len(colors)])
            shape.line.width = Pt(0)

    def add_profile(self, info, position, is_small=False):
        """
        info: dict {
          'name': str,         # 한글 이름
          'eng_name': str,     # 영문 이름 (파일명 매칭에 사용)
          'department': str,
          'birth_month': int,
          'birth_day': int,
          'image_path': str (or None)
        }
        position: (leftEMU, topEMU)
        is_small: True이면 사진/폰트 등을 더 작게
        """
        left, top = position
        
        # 사진 크기: 기본 1.5인치, is_small이면 1.2인치
        photo_inch = 1.5 if not is_small else 1.2
        photo_emu = Inches(photo_inch)
        
        # 1) 원형 사진 or "No Photo"
        valid_photo = False
        if info.get('image_path') and os.path.exists(info['image_path']):
            try:
                temp_path = f"temp_circle_{uuid.uuid4().hex}.png"
                size_tuple = (300,300) if not is_small else (240,240)
                self.create_circle_image(info['image_path'], temp_path, size=size_tuple)
                self.slide.shapes.add_picture(temp_path, left, top, height=photo_emu)
                os.remove(temp_path)
                valid_photo = True
            except (UnidentifiedImageError, OSError):
                pass
        
        if not valid_photo:
            # 원형 도형 + "No Photo"
            shape = self.slide.shapes.add_shape(
                MSO_SHAPE.OVAL, left, top, photo_emu, photo_emu
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(200, 200, 200)
            shape.line.width = Pt(0)
            tf = shape.text_frame
            tf.text = "No Photo"
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            for run in tf.paragraphs[0].runs:
                run.font.size = Pt(10 if is_small else 11)
                run.font.name = '맑은 고딕'
        
        # 2) 노란색 박스: birth_month/birth_day
        birth_str = f"{info.get('birth_month','')}/{info.get('birth_day','')}"
        bday_top = top + Inches(photo_inch * (0.8 if not is_small else 0.75))
        box_w = Inches(0.8 if not is_small else 0.6)
        box_h = Inches(0.3 if not is_small else 0.25)
        
        bday_box = self.slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left,
            bday_top,
            box_w,
            box_h
        )
        bday_box.fill.solid()
        bday_box.fill.fore_color.rgb = RGBColor(255, 200, 0)  # 노란색
        bday_box.line.width = Pt(0)
        
        tf = bday_box.text_frame
        tf.text = birth_str
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # 3) 텍스트(부서 + 영문명(한글명))
        dept = info.get('department','')
        eng_name = info.get('eng_name','')
        kor_name = info.get('name','')
        
        name_box_left = left - Inches(0.25)
        name_box_top = bday_top + Inches(0.4 if not is_small else 0.3)
        name_box_w = Inches(2.0 if not is_small else 1.6)
        name_box_h = Inches(0.8 if not is_small else 0.6)
        
        name_box = self.slide.shapes.add_textbox(name_box_left, name_box_top, name_box_w, name_box_h)
        tx = name_box.text_frame
        tx.word_wrap = True
        
        # "부서\n영문이름 (한글이름)"
        # 2줄로 나눈 다음, 첫 줄(부서)에만 별도 색 적용
        p = tx.paragraphs[0]
        p.text = ""
        
        # 1) 부서 paragraph
        para_dept = tx.add_paragraph()
        para_dept.alignment = PP_ALIGN.CENTER
        para_dept.text = dept
        
        # 부서만 파란색
        for run in para_dept.runs:
            run.font.size = Pt(11 if not is_small else 9)
            run.font.name = '맑은 고딕'
            run.font.color.rgb = RGBColor(0,128,255)
        
        # 2) 이름 paragraph
        para_name = tx.add_paragraph()
        para_name.alignment = PP_ALIGN.CENTER
        para_name.text = f"{eng_name} ({kor_name})"
        for run in para_name.runs:
            run.font.size = Pt(11 if not is_small else 9)
            run.font.name = '맑은 고딕'

    def create_layout(self, people_data):
        if not people_data:
            return
        
        # 배경색
        bg = self.slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor.from_string(self.theme['background'])
        
        # 제목
        title_w = Inches(12)
        title_h = Inches(1)
        title_left = (self.SLIDE_WIDTH - title_w) / 2
        title_top = Inches(0.8)
        
        title_box = self.slide.shapes.add_textbox(title_left, title_top, title_w, title_h)
        p = title_box.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.clear()
        
        run = p.add_run()
        run.text = f"{self.theme['message']}"
        run.font.size = Pt(36)
        run.font.bold = True
        run.font.name = '맑은 고딕'
        run.font.color.rgb = RGBColor.from_string(self.theme['title_color'])
        
        # 부제목
        subtitle_w = Inches(12)
        subtitle_h = Inches(0.5)
        subtitle_left = (self.SLIDE_WIDTH - subtitle_w) / 2
        subtitle_top = Inches(1.8)
        
        subtitle_box = self.slide.shapes.add_textbox(subtitle_left, subtitle_top, subtitle_w, subtitle_h)
        sp = subtitle_box.text_frame.paragraphs[0]
        sp.text = self.theme['sub_message']
        sp.alignment = PP_ALIGN.CENTER
        for r in sp.runs:
            r.font.size = Pt(14)
            r.font.name = '맑은 고딕'
        
        # 장식
        self.add_decorations()
        
        # 인원수에 따라 레이아웃/크기 변경
        n = len(people_data)
        is_small = (n >= 10)  # 10명 이상이면 작게
        
        # 예: 10명 이하 -> 2행5열까지, 그 이상 -> 행 수 늘려 5열
        #    6명 이하 -> (2,3) 등등은 자유롭게 변경 가능
        if n <= 6:
            rows, cols = (2, 3) if n > 3 else (1, n)
        elif n <= 10:
            rows, cols = (2, 5)
        else:
            cols = 5
            rows = math.ceil(n / cols)
        
        profile_w_in = 2.0 if not is_small else 1.6
        profile_h_in = 3.0 if not is_small else 2.6
        
        total_w = cols * profile_w_in
        total_h = rows * profile_h_in
        
        top_margin_in = 2.5
        usable_h_in = (self.SLIDE_HEIGHT / 914400.0) - top_margin_in
        
        start_top_in = top_margin_in + (usable_h_in - total_h) / 2
        if start_top_in < top_margin_in:
            start_top_in = top_margin_in
        
        start_left_in = ((self.SLIDE_WIDTH / 914400.0) - total_w) / 2
        
        start_left_emu = start_left_in * 914400
        start_top_emu = start_top_in * 914400
        
        # 실제 배치
        for i, person in enumerate(people_data):
            r = i // cols
            c = i % cols
            
            left_emu = start_left_emu + c * Inches(profile_w_in)
            top_emu  = start_top_emu + r * Inches(profile_h_in)
            
            self.add_profile(person, (left_emu, top_emu), is_small=is_small)

    def save(self):
        """PPT를 BytesIO로 반환"""
        buf = BytesIO()
        self.prs.save(buf)
        buf.seek(0)
        return buf

############################
# 이미지 매칭 함수 (eng_name -> 파일 찾기)
############################
def find_image_by_engname(eng_name, image_map):
    """
    eng_name: 예) "Mark" (대소문자 구분 없이 매칭)
    image_map: { "mark.png": "/path/mark.png", "dk.jpg": "/path/dk.jpg", ... }
    - 아래 규칙: eng_name + ".png"/".jpg"/".jpeg" 를 우선 찾아봄
    """
    if not eng_name:
        return None
    
    base = eng_name.strip().lower()  # "mark"
    possible_exts = [".png", ".jpg", ".jpeg"]
    for ext in possible_exts:
        candidate = base + ext  # ex) "mark.png"
        if candidate in image_map:  # image_map 키는 이미 lower() 상태
            return image_map[candidate]
    return None

############################
# Streamlit 앱
############################
def main():
    st.title("AI 생일자 PPT 자동 생성기 (eng_name = 파일명 매칭)")
    st.write("**엑셀에는 `eng_name`이 들어있고, 업로드된 파일명은 반드시 `eng_name.(png|jpg|jpeg)`**로 맞춰주십시오.")
    st.write("예: 엑셀 eng_name=Mark → 업로드 파일명=Mark.png / mArK.jpg 등 대소문자 무시")
    
    excel_file = st.file_uploader("엑셀 파일 업로드 (xlsx, xls)", type=["xlsx","xls"])
    image_files = st.file_uploader("이미지 파일들 (여러 개 가능)", type=["png","jpg","jpeg"], accept_multiple_files=True)
    
    selected_month = st.number_input("생일 달을 입력하세요 (1~12)", min_value=1, max_value=12, value=1)
    
    if st.button("PPT 생성"):
        if not excel_file:
            st.warning("엑셀 파일을 먼저 업로드하세요.")
            return

        # 이미지 임시 폴더
        temp_folder = f"temp_{uuid.uuid4().hex}"
        os.makedirs(temp_folder, exist_ok=True)

        # 1) 업로드한 이미지들을 lower() 키로 dict에 저장
        image_map = {}
        for f in image_files:
            filename_lower = f.name.lower()  # "mark.png"
            path = os.path.join(temp_folder, f.name)
            with open(path, "wb") as outfile:
                outfile.write(f.read())
            image_map[filename_lower] = path
        
        # 2) 엑셀 읽기
        df = pd.read_excel(excel_file)
        # 예: 컬럼 -> [name, eng_name, department, birth_month, birth_day]
        
        # 3) selected_month만 필터
        df = df[df['birth_month'] == selected_month]
        if df.empty:
            st.error(f"{selected_month}월 생일자가 없습니다.")
            return
        
        # 4) people_data 구성 (eng_name 기반으로 파일 찾기)
        people_data = []
        for _, row in df.iterrows():
            e_name = str(row.get('eng_name','')).strip()
            img_path = find_image_by_engname(e_name, image_map)  # 없으면 None
            
            info = {
                'name': row.get('name',''),
                'eng_name': e_name,
                'department': row.get('department',''),
                'birth_month': row.get('birth_month',''),
                'birth_day': row.get('birth_day',''),
                'image_path': img_path
            }
            people_data.append(info)
        
        if not people_data:
            st.error(f"엑셀 행은 존재하나, {selected_month}월에 해당하는 eng_name과 매칭되는 이미지가 없습니다.")
            return
        
        # 5) PPT 생성
        generator = BirthdaySlideGenerator(month=selected_month)
        generator.create_layout(people_data)
        ppt_bytes = generator.save()

        # 6) 다운로드
        st.success(f"PPT 생성 완료 (총 {len(people_data)}명). 아래 버튼을 눌러 다운로드:")
        st.download_button(
            label="Download PPT",
            data=ppt_bytes,
            file_name="birthday_slide_result.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        # (선택) 임시 폴더 삭제
        # import shutil
        # shutil.rmtree(temp_folder, ignore_errors=True)

if __name__ == "__main__":
    main()
