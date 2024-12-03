import pyhwp

def extract_text_from_hwp(file_path):
    # HWP 파일을 엽니다
    hwp = pyhwp.Document(file_path)
    
    # 모든 문단에서 텍스트를 추출합니다
    paragraphs = []
    for section in hwp.bodytext.sections:
        for para in section.paragraphs:
            paragraphs.append(para.get_text())
            
    return '\n'.join(paragraphs)

file_path = "test.hwp"
extracted_text = extract_text_from_hwp(file_path)
print(extracted_text)
