import re
import time
from time import sleep
from typing import Union

import pywinauto
import pywinauto.base_wrapper
import pywinauto.findwindows
import win32com.client as win32
import win32con
import win32gui
from pywinauto import Application, mouse, win32functions

import process_utils

CDispatch = Union[win32.CDispatch, win32.dynamic.CDispatch]


class Colvir:
    def __init__(self, process_path: str, user: str, password: str):
        self.process_path = process_path
        self.user = user
        self.password = password
        self.app = self.open_colvir()

    def open_colvir(self) -> pywinauto.Application:
        app = None
        for _ in range(10):
            try:
                app = pywinauto.Application().start(cmd_line=self.process_path)
                self.login(app=app, user=self.user, password=self.password)
                self.check_interactivity(app=app)
                break
            except pywinauto.findwindows.ElementNotFoundError:
                if isinstance(app, Application) and self.change_password(app):
                    break
                process_utils.kill_all_processes("COLVIR")
                continue

        assert app is not None, Exception("max_retries exceeded")
        return app

    @staticmethod
    def change_password(app: pywinauto.Application) -> bool:
        attention_win = app.window(title="Внимание")
        if not attention_win.exists():
            return False
        attention_win["OK"].click()

        change_password_win = app.window(title_re="Смена пароля.+")
        change_password_win["Edit0"].set_text("ROBOTIZ2024_")
        change_password_win["Edit2"].set_text("ROBOTIZ2024_")
        change_password_win["OK"].click()

        confirm_win = app.window(title="Colvir Banking System", found_index=1)
        confirm_win.send_keystrokes("{ENTER}")

        mode_win = app.window(title="Выбор режима")
        return mode_win.exists()

    @staticmethod
    def login(app: pywinauto.Application, user: str, password: str) -> None:
        if not user or not password:
            raise ValueError("COLVIR_USR or COLVIR_PSW is not set")

        login_win = app.window(title="Вход в систему")

        login_username = login_win["Edit2"]
        login_password = login_win["Edit"]

        login_username.set_text(text=user)
        if login_username.window_text() != user:
            login_username.set_text("")
            login_username.type_keys(user, set_foreground=False)

        login_password.set_text(text=password)
        if login_password.window_text() != password:
            login_password.set_text("")
            login_password.type_keys(password, set_foreground=False)

        login_win["OK"].click()

        sleep(1)
        if login_win.exists() and app.window(title="Произошла ошибка").exists():
            raise pywinauto.findwindows.ElementNotFoundError()

    @staticmethod
    def check_interactivity(app: pywinauto.Application) -> None:
        choose_mode(app=app, mode="KREQDOC")
        sleep(1)

        close_window(win=app.window(title="Фильтр"), raise_error=True)

    def get_app(self) -> pywinauto.Application:
        assert self.app is not None
        return self.app


def set_focus_win32(win: pywinauto.WindowSpecification) -> None:
    if win.wrapper_object().has_focus():
        return

    handle = win.wrapper_object().handle

    mouse.move(coords=(-10000, 500))
    if win.is_minimized():
        if win.was_maximized():
            win.maximize()
        else:
            win.restore()
    else:
        win32gui.ShowWindow(handle, win32con.SW_SHOW)
    win32gui.SetForegroundWindow(handle)

    win32functions.WaitGuiThreadIdle(handle)


def set_focus(win: pywinauto.WindowSpecification, retries: int = 20) -> None:
    while retries > 0:
        try:
            if retries % 2 == 0:
                set_focus_win32(win)
            else:
                win.set_focus()
            break
        except (Exception, BaseException):
            retries -= 1
            time.sleep(5)
            continue

    if retries <= 0:
        raise Exception("Failed to set focus")


def press(
    win: pywinauto.WindowSpecification, key: str, pause: float = 0
) -> None:
    set_focus(win)
    win.type_keys(key, pause=pause, set_foreground=False)


def choose_mode(app: pywinauto.Application, mode: str) -> None:
    mode_win = app.window(title="Выбор режима")
    mode_win["Edit2"].set_text(text=mode)
    press(mode_win["Edit2"], "~")


def close_window(
    win: pywinauto.WindowSpecification, raise_error: bool = False
) -> None:
    if win.exists():
        win.close()
        return

    if raise_error:
        raise pywinauto.findwindows.ElementNotFoundError(
            f"Window {win} does not exist"
        )


def get_window(
    app: pywinauto.Application,
    title: str,
    wait_for: str = "exists",
    timeout: int = 20,
    regex: bool = False,
    found_index: int = 0,
) -> pywinauto.WindowSpecification:
    window = (
        app.window(title=title, found_index=found_index)
        if not regex
        else app.window(title_re=title, found_index=found_index)
    )
    window.wait(wait_for=wait_for, timeout=timeout)
    time.sleep(0.5)
    return window


def type_keys(
    window: pywinauto.WindowSpecification,
    keystrokes: str,
    step_delay: float = 0.1,
    delay_after: float = 0.5,
) -> None:
    set_focus(window)
    for command in list(filter(None, re.split(r"({.+?})", keystrokes))):
        try:
            window.type_keys(command, set_foreground=False)
        except pywinauto.base_wrapper.ElementNotEnabled:
            time.sleep(1)
            window.type_keys(command, set_foreground=False)
        time.sleep(step_delay)

    time.sleep(delay_after)


def find_and_click_button(
    app: pywinauto.Application,
    window: pywinauto.WindowSpecification,
    toolbar: pywinauto.WindowSpecification,
    target_button_name: str,
    horizontal: bool = True,
    offset: int = 5,
) -> None:
    status_win = app.window(title_re="Банковская система.+")
    rectangle = toolbar.rectangle()
    mid_point = rectangle.mid_point()
    window.move_mouse_input(coords=(mid_point.x, mid_point.y), absolute=True)

    start_point = rectangle.left if horizontal else rectangle.top
    end_point = rectangle.right if horizontal else rectangle.bottom

    x, y = mid_point.x, mid_point.y
    point = 0

    x_offset = offset if horizontal else 0
    y_offset = offset if not horizontal else 0

    i = 0
    while (
        status_win["StatusBar"].window_text().strip() != target_button_name
        or point >= end_point
    ):
        point = start_point + i * 5

        if horizontal:
            x = point
        else:
            y = point

        window.move_mouse_input(coords=(x, y), absolute=True)
        i += 1

    window.set_focus()
    sleep(1)
    window.click_input(
        button="left", coords=(x + x_offset, y + y_offset), absolute=True
    )
