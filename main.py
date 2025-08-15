import win32gui, win32api, win32con, win32com.client
# from paddleocr import PaddleOCR
# import numpy as np
# import paddle
import tools
import ctypes
import time
import os


def get_screen_size():
    """
    获取屏幕的分辨率
    返回值： (width, height)
    """
    width = win32api.GetSystemMetrics(0)
    height = win32api.GetSystemMetrics(1)
    return width, height


def get_game_size():
    """
    获取指定窗口的大小
    返回值： (width, height)
    """
    if hwnd:
        _, _, client_width, client_height = win32gui.GetClientRect(hwnd)
        return client_width, client_height
    else:
        return None


def calculate_target_position(size):
    """计算目标点击位置"""
    # 计算目标坐标 (基于屏幕尺寸的比例)
    target_x = int(size[0] * 0.721)
    target_y = int(size[1] * 0.271)
    return target_x, target_y


def letsgo_quit():
    """退出程序"""
    print("程序已退出")
    os.system("pause")
    quit()


def calculation_offset():
    """计算窗口偏移量"""
    x, y, window_width, window_height = win32gui.GetWindowRect(hwnd)
    client_width, client_height = get_game_size()

    print(f"窗口大小：{window_width}x{window_height}")
    print(f"客户区大小：{client_width}x{client_height}")
    borderx = (window_width - x - client_width) // 2
    bordery = window_height - y - client_height - borderx

    print(f"窗口边框大小：{borderx},{bordery}")

    return (client_width, client_height), x + borderx, y + bordery


def is_16_9():
    """检测屏幕是否为16:9"""
    # 获取屏幕和游戏窗口的大小
    screen_size = get_screen_size()
    print(f"屏幕分辨率：{screen_size}")
    game_size = get_game_size()
    print(f"游戏窗口分辨率：{game_size}")
    offset_x = 0
    offset_y = 0
    # 比较游戏窗口和屏幕的分辨率是否相等
    if screen_size == game_size:
        print(f"屏幕分辨率：{screen_size}")
        if screen_size[0] / 16 != screen_size[1] / 9:
            print(
                "屏幕分辨率不是16:9，当前暂不支持其他比例的屏幕，请将游戏窗口化并设置分辨率的比例为16:9"
            )
            letsgo_quit()
        return screen_size, offset_x, offset_y
    else:
        if game_size[0] / 16 != game_size[1] / 9:
            print("请将游戏分辨率的比例设为16:9，前往“更多->设置->画面->窗口”进行调整")
            letsgo_quit()
        # 计算偏移量
        return calculation_offset()


# 禁止最小化按钮
def disable_minimize_button():
    """禁用窗口的最小化按钮"""
    # 2. 获取当前窗口样式
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    # 3. 禁用最小化和最大化按钮（移除 WS_MINIMIZEBOX）
    new_style = style & ~win32con.WS_MINIMIZEBOX & ~win32con.WS_MAXIMIZEBOX
    # 4. 应用新样式
    win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, new_style)
    # 5. 刷新窗口（可选）
    win32gui.SetWindowPos(
        hwnd,
        0,
        0,
        0,
        0,
        0,
        win32con.SWP_NOSIZE
        | win32con.SWP_NOMOVE
        | win32con.SWP_NOZORDER
        | win32con.SWP_FRAMECHANGED,
    )


# def checkServerConnection():
#     """检查服务器连接状态"""
#     client_width, client_height = get_game_size()
#     image = tools.screenshot_window(
#         hwnd,
#         (
#             int(client_width * 0.525),
#             int(client_height * 0.6),
#             int(client_width * 0.675),
#             int(client_height * 0.675),
#         ),
#     )
#     # image.show()

#     device = "gpu" if paddle.is_compiled_with_cuda() else "cpu"
#     # 初始化 PaddleOCR 实例
#     ocr = PaddleOCR(
#         lang="ch",
#         use_doc_orientation_classify=False,
#         use_doc_unwarping=False,
#         use_textline_orientation=False,
#         device=device,  # 如果有 GPU 可用，改为 "gpu"
#     )

#     # 对示例图像执行 OCR 推理
#     result = ocr.predict(np.array(image))
#     if result[0]["rec_texts"][0] == "继续尝试":
#         return True
#     else:
#         return False


if __name__ == "__main__":
    # 检查管理员权限
    if not ctypes.windll.shell32.IsUserAnAdmin():
        print("请以管理员身份运行->右键点击程序->以管理员身份运行")
        letsgo_quit()
    # 获取游戏窗口标题和类名
    window_title = "元梦之星模拟器高清版"
    window_class = "UnrealWindow"

    # 判断游戏窗口是否存在
    hwnd = win32gui.FindWindow(window_class, window_title)
    if hwnd:
        print(
            f"找到窗口：{window_title}，句柄为：{hwnd}，类型: {win32gui.GetClassName(hwnd)}"
        )
        disable_minimize_button()
    else:
        print(f"未找到标题为'{window_title}'的窗口")
        letsgo_quit()

    # 等待用户输入循环秒数
    duration = int(input("请输入循环秒数后按回车（不得小于30）："))
    duration = max(duration, 30)

    # 获取游戏窗口大小和偏移量
    size, offset_x, offset_y = is_16_9()
    print(f"游戏窗口大小：{size}, 偏移量：{offset_x}, {offset_y}")
    # 初始化输入处理类
    kmc = tools.KeyboardMouseControl(hwnd, size, offset_x, offset_y)

    # 激活游戏窗口
    time.sleep(1)
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys("%")
    print("激活游戏窗口")
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(1)

    print("开始执行")
    while True:
        start_time = time.time()

        # 检查服务器连接状态
        # if checkServerConnection():
        kmc.mouse_move_click(0.63, 0.65)
        time.sleep(3)
        # ==================== 循环体开始 ====================
        # 在这里编写你的核心业务逻辑代码

        # 拒绝好友拉到身边
        kmc.mouse_move_click(0.721, 0.271)
        for i in ["w", "a", "s", "d"]:
            # 点击空白地方
            kmc.mouse_move_click(0.01, 0.99)
            # 按下 WASD 键
            kmc.press_key(i, 0.1)
            time.sleep(0.1)
        kmc.press_key("r", 0.2)
        kmc.key_down("a")
        kmc.press_key("space", 0.1)
        time.sleep(1)
        kmc.key_up("a")
        kmc.press_key("q", 0.1)
        # ==================== 循环体结束 ====================
        # 计算本次执行耗时
        execution_time = time.time() - start_time
        wait_time = duration - execution_time
        print(f"执行时间 {execution_time:.2f}s < {duration}s，等待 {wait_time:.2f}s")
        time.sleep(wait_time)
