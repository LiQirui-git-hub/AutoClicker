import os
import time
import threading
from pynput import mouse
import win32gui

from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text
from rich import box

import json

read_struct = json.load(open("Config.json"))

click_delay = read_struct["ClickDelay"]
hold_threshold = read_struct["HoldThreshold"]

console = Console()

is_clicking_left = False
is_clicking_right = False
is_self_clicking = False
mouse_controller = mouse.Controller()

hold_timer_left = None
hold_timer_right = None
press_mouse_pos_left = (0, 0)
press_mouse_pos_right = (0, 0)
press_window_pos = None

click_count_left = 0
start_time_left = 0.0
click_count_right = 0
start_time_right = 0.0


def get_window_at_mouse():
    try:
        x, y = mouse_controller.position
        hwnd = win32gui.WindowFromPoint((x, y))
        if hwnd:
            rect = win32gui.GetWindowRect(hwnd)
            return {
                "hwnd": hwnd,
                "pos": (rect[0], rect[1])
            }
        return None
    except Exception as e:
        console.print(f"[red]获取窗口信息失败: {e}[/red]")
        return None


def is_dragging_window(press_window, current_mouse_pos):
    if not press_window:
        return False

    current_window = get_window_at_mouse()
    if not current_window or current_window["hwnd"] != press_window["hwnd"]:
        return False

    dx = abs(current_window["pos"][0] - press_window["pos"][0])
    dy = abs(current_window["pos"][1] - press_window["pos"][1])
    return dx > 2 or dy > 2


def clicker_left():
    global is_clicking_left, click_count_left, is_self_clicking
    click_count_left = 0
    console.rule()
    console.print(Panel("[green bold][左键] 启动连点[/green bold]", expand=True))
    while is_clicking_left:
        is_self_clicking = True
        mouse_controller.press(mouse.Button.left)
        mouse_controller.release(mouse.Button.left)
        is_self_clicking = False
        click_count_left += 1
        time.sleep(click_delay)


def clicker_right():
    global is_clicking_right, click_count_right, is_self_clicking
    click_count_right = 0
    console.rule()
    console.print(Panel("[blue bold][右键] 启动连点[/blue bold]", expand=True))
    while is_clicking_right:
        is_self_clicking = True
        mouse_controller.press(mouse.Button.right)
        mouse_controller.release(mouse.Button.right)
        is_self_clicking = False
        click_count_right += 1
        time.sleep(click_delay)


def show_stats(button_type, duration, count, frequency):
    table = Table(box=box.ROUNDED, expand=True)
    table.add_column("指标", justify="center", style="yellow")
    table.add_column("数值", justify="center")

    table.add_row("按住时长", f"{duration:.2f} 秒")
    table.add_row("总点击次数", f"{count} 次")
    table.add_row("平均频率", f"{frequency:.0f} 次/秒")

    console.print(table)


def on_click(x, y, button, pressed):
    global is_clicking_left, is_clicking_right
    global hold_timer_left, hold_timer_right, press_mouse_pos_left, press_mouse_pos_right
    global press_window_pos, click_count_left, click_count_right, start_time_left, start_time_right
    global is_self_clicking

    if is_self_clicking:
        return

    try:
        if button == mouse.Button.left:
            if pressed:
                if not is_clicking_left:
                    press_mouse_pos_left = (x, y)
                    press_window_pos = get_window_at_mouse()

                    def delayed_start():
                        global is_clicking_left, start_time_left
                        current_mouse_pos = mouse_controller.position
                        if not is_dragging_window(press_window_pos, current_mouse_pos):
                            is_clicking_left = True
                            start_time_left = time.time()
                            threading.Thread(target=clicker_left, daemon=True).start()

                    hold_timer_left = threading.Timer(hold_threshold, delayed_start)
                    hold_timer_left.start()
            else:
                if hold_timer_left is not None and hold_timer_left.is_alive():
                    hold_timer_left.cancel()
                elif is_clicking_left:
                    is_clicking_left = False
                    duration = time.time() - start_time_left
                    frequency = click_count_left / duration if duration > 0 else 0
                    console.print(Panel("[red bold][左键] 停止连点[/red bold]", expand=True))
                    show_stats("左键", duration, click_count_left, frequency)
                    console.rule()

        elif button == mouse.Button.right:
            if pressed:
                if not is_clicking_right:
                    press_mouse_pos_right = (x, y)
                    press_window_pos = get_window_at_mouse()

                    def delayed_start():
                        global is_clicking_right, start_time_right
                        current_mouse_pos = mouse_controller.position
                        if not is_dragging_window(press_window_pos, current_mouse_pos):
                            is_clicking_right = True
                            start_time_right = time.time()
                            threading.Thread(target=clicker_right, daemon=True).start()

                    hold_timer_right = threading.Timer(hold_threshold, delayed_start)
                    hold_timer_right.start()
            else:
                if hold_timer_right is not None and hold_timer_right.is_alive():
                    hold_timer_right.cancel()
                elif is_clicking_right:
                    is_clicking_right = False
                    duration = time.time() - start_time_right
                    frequency = click_count_right / duration if duration > 0 else 0
                    console.print(Panel("[red bold][右键] 停止连点[/red bold]", expand=True))
                    show_stats("右键", duration, click_count_right, frequency)
                    console.rule()

    except Exception as e:
        console.print(Panel(f"[red]鼠标事件错误: {e}[/red]", expand=True))


def on_move(x, y):
    global hold_timer_left, hold_timer_right, is_clicking_left, is_clicking_right

    if hold_timer_left is not None and hold_timer_left.is_alive():
        if is_dragging_window(press_window_pos, (x, y)):
            hold_timer_left.cancel()

    if hold_timer_right is not None and hold_timer_right.is_alive():
        if is_dragging_window(press_window_pos, (x, y)):
            hold_timer_right.cancel()

    if is_clicking_left and is_dragging_window(press_window_pos, (x, y)):
        is_clicking_left = False
        console.print(Panel("[orange][左键] 检测到窗口拖动，停止连点[/orange]"))
        console.rule()
    if is_clicking_right and is_dragging_window(press_window_pos, (x, y)):
        is_clicking_right = False
        console.print(Panel("[orange][右键] 检测到窗口拖动，停止连点[/orange]"))
        console.rule()


def main():
    welcome_text = Text("鼠标连点器", style="bold magenta")
    welcome_panel = Panel(
        f"""
核心配置：
   • 长按判定：{hold_threshold} 秒
   • 连点间隔：{click_delay} 秒/次

使用说明：
   • 单纯长按鼠标 > {hold_threshold} 秒 → 启动连点
   • 拖动窗口/控件 → 不触发连点
   • 松开鼠标/拖动窗口 → 停止连点
   • 关闭控制台 → 退出程序
""",
        title=welcome_text,
        border_style="green",
        expand=True
    )
    console.print(welcome_panel)

    listener = mouse.Listener(on_click=on_click, on_move=on_move)
    listener.start()
    listener.join()


if __name__ == "__main__":
    os.system("cls")
    main()