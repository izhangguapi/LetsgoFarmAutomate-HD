import win32gui
import win32ui
import win32con
from ctypes import windll
from PIL import Image
import numpy as np
from paddleocr import PaddleOCR
import paddle





window_title = "元梦之星模拟器高清版"
window_class = "UnrealWindow"
hwnd = win32gui.FindWindow(window_class, window_title)

print(f"找到窗口：{window_title}，句柄为：{hwnd}，类型: {win32gui.GetClassName(hwnd)}")

_, _, client_width, client_height = win32gui.GetClientRect(hwnd)

image = screenshot_window(
    (
        int(client_width * 0.525),
        int(client_height * 0.6),
        int(client_width * 0.675),
        int(client_height * 0.675),
    )
)
# image.show()

device = "gpu" if paddle.is_compiled_with_cuda() else "cpu"
# 初始化 PaddleOCR 实例
ocr = PaddleOCR(
    lang="ch",
    use_doc_orientation_classify=False,
    use_doc_unwarping=False,
    use_textline_orientation=False,
    device=device,  # 如果有 GPU 可用，改为 "gpu"
)

# 对示例图像执行 OCR 推理
result = ocr.predict(np.array(image))

print(type(result))
print(result[0]["rec_texts"])

# 可视化结果并保存 json 结果
# for res in result:
#     res.print("rec_texts")
#     res.save_to_img("output")
#     res.save_to_json("output")
