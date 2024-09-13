from PIL import Image, ImageDraw

# 이미지를 엽니다
image_path = "C://Users//wpghk//Desktop//automatic//slides//슬라이드12.JPG"
image = Image.open(image_path)

# OCR을 적용할 영역의 좌표 (좌표는 필요에 따라 조정하세요)
# (left, upper, right, lower)
ocr_area = (0, 30, 1080, 600)

# 이미지를 복사하고, 해당 영역을 표시합니다
image_with_box = image.copy()
draw = ImageDraw.Draw(image_with_box)
draw.rectangle(ocr_area, outline="red", width=3)

# 결과 이미지를 저장합니다
output_image_path = "C://Users//wpghk//Desktop//automatic//slides//슬라이드12_with_box.JPG"
image_with_box.save(output_image_path)

# 저장된 이미지를 보여줍니다
image_with_box.show()

print(f"결과 이미지를 '{output_image_path}'에 저장했습니다.")
