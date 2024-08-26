import logging
import os
import sys
import time
import traceback
import warnings
from datetime import datetime
from functools import wraps
from time import sleep
from typing import Any, Callable, Tuple

import dotenv
import pandas as pd
import pyperclip
import pywinauto

import bpm

try:
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    import src.colvir_utils as colvir_utils
    import src.data as data
    import src.process_utils as process_utils
    from src.logger import setup_logger
    from src.mail import send_mail
    from src.notification import TelegramAPI, send_message
    from src.wiggle import wiggle_mouse
except Exception as exc:
    exception_traceback = traceback.format_exc()
    raise exc


def handle_error(func: Callable[..., Any]) -> Callable[..., Any]:
    @wraps(func)
    def wrapper(*args, **kwargs) -> Any:
        bot = kwargs.get("bot")

        try:
            return func(*args, **kwargs)
        except KeyboardInterrupt as error:
            raise error
        except (Exception, BaseException) as error:
            logging.exception(error)
            message = traceback.format_exc()

            if bot:
                send_message(bot, message)
            raise error

    return wrapper


def fill_filter_win(
    app: pywinauto.Application, year: str, order_id: str
) -> None:
    filter_win = colvir_utils.get_window(app=app, title="Фильтр")

    start_date = f"01.01.{year}"
    end_date = f"31.12.{year}"

    filter_win["Edit8"].set_text(start_date)
    assert filter_win["Edit8"].window_text() == start_date

    filter_win["Edit10"].set_text(end_date)
    assert filter_win["Edit10"].window_text() == end_date

    filter_win["Edit6"].set_text(order_id)
    assert filter_win["Edit6"].window_text() == order_id

    filter_win["OK"].click()

    sleep(1)


def get_kbk(text: str, rk: bool) -> Tuple[str, str]:
    text = text.lower()
    if rk is False:
        budget_type = "EXC"
        if "суточные" in text:
            kbk = "80302020201"
        elif "проезд" in text:
            kbk = "80302020202"
        elif "проживание" in text:
            kbk = "80302020203"
        elif "штраф" in text:
            kbk = "80213"
        elif "отмена" in text:
            kbk = "803030903"
        elif "сверх норм" in text:
            kbk = "80302020301"
        else:
            kbk = "70302020204"
            budget_type = "CPC"
    else:
        budget_type = "CPC"
        if "суточные" in text:
            kbk = "70302020101"
        elif "проезд" in text:
            kbk = "70302020102"
        elif "проживание" in text:
            kbk = "70302020103"
        elif "штраф" in text:
            kbk = "80213"
            budget_type = "EXC"
        elif "отмена" in text:
            kbk = "803030903"
            budget_type = "EXC"
        else:
            kbk = "70302020104"

    return kbk, budget_type


def new_finance(
    app: pywinauto.Application,
    journal_win: pywinauto.WindowSpecification,
    request: data.Request,
) -> None:
    colvir_utils.find_and_click_button(
        app=app,
        window=journal_win,
        toolbar=journal_win["Static2"],
        target_button_name="Создать новую финансовую запись",
        horizontal=False,
    )

    finance_win = colvir_utils.get_window(app=app, title="Финансовая запись")
    finance_win.set_focus()

    finance_win["Edit46"].click_input()
    finance_win["Edit46"].type_keys("{BACKSPACE}" * 8)
    sleep(0.5)
    finance_win["Edit46"].type_keys("28000504")
    sleep(0.5)
    finance_win["Edit42"].click_input()
    finance_win["Edit42"].type_keys("{BACKSPACE}")
    finance_win["Edit42"].type_keys("1")
    sleep(0.5)
    finance_win["Edit16"].click_input()
    finance_win["Edit16"].type_keys("{BACKSPACE}" * 20)
    sleep(0.5)
    finance_win["Edit16"].type_keys("KZ54907A185400000035")
    sleep(0.5)
    finance_win["Edit14"].click_input()
    finance_win["Edit14"].type_keys("{LEFT}" * len(request.oz))
    finance_win["Edit14"].type_keys(request.oz)
    sleep(0.5)
    finance_win["Edit32"].click_input()
    finance_win["Edit32"].type_keys("{BACKSPACE}" * 20)
    sleep(0.5)
    finance_win["Edit32"].type_keys("KZ45907A185400000003")
    sleep(1)
    finance_win["Edit30"].click_input()
    sleep(1)
    finance_win["Edit14"].click_input()
    sleep(1)
    finance_win["Edit30"].click_input()
    sleep(1)
    finance_win["TDBMemo"].type_keys(
        "{BACKSPACE}" * len(str(request.reimbursement)),
    )
    finance_win["TDBMemo"].type_keys(
        str(request.reimbursement), with_spaces=True, pause=0.05
    )

    for class_name in [
        "Edit46",
        "Edit42",
        "Edit16",
        "Edit14",
        "Edit32",
        "Edit30",
        "Edit14",
        "TDBMemo",
    ]:
        finance_win[class_name].click_input()
        sleep(1)

    colvir_utils.find_and_click_button(
        app=app,
        window=finance_win,
        toolbar=finance_win["Static3"],
        target_button_name="Сохранить изменения (PgDn)",
    )


def fill_order(
    app: pywinauto.Application,
    business_trip_order_win: pywinauto.WindowSpecification,
    now: datetime,
    request: data.Request,
    rk: bool,
) -> str:
    business_trip_order_win.menu_select("#0->#5->#0")
    confirm_pay_win = colvir_utils.get_window(
        app=app, title="Подтверждение", wait_for="exists enabled"
    )
    confirm_pay_win["&Да"].click()

    payment_win = colvir_utils.get_window(
        app=app,
        title="Оплата КОМАНДИРОВОК.+",
        wait_for="exists enabled",
        regex=True,
    )
    payment_win["OK"].click()

    wiggle_mouse(duration=1)

    business_trip_order_win.wait(wait_for="enabled", timeout=20)
    business_trip_order_win.menu_select("#0->#5->#1")

    confirm_pay_win = colvir_utils.get_window(
        app=app, title="Подтверждение", wait_for="exists enabled"
    )
    confirm_pay_win["&Да"].click()

    change_win = colvir_utils.get_window(
        app=app, title="Изменение/добавление позиции"
    )

    for row in request.rows:
        colvir_utils.find_and_click_button(
            app, change_win, change_win["Static4"], "Создать новую запись (Ins)"
        )

        colvir_utils.type_keys(
            window=change_win, keystrokes="{ENTER}{SPACE}{ENTER}{RIGHT}"
        )
        change_win.type_keys(row.name, with_spaces=True)
        colvir_utils.type_keys(window=change_win, keystrokes="{ENTER}{RIGHT}")
        change_win.type_keys(row.sum_tenge, with_spaces=True)
        colvir_utils.type_keys(window=change_win, keystrokes="{ENTER}{RIGHT}")
        change_win.type_keys("1", with_spaces=True)
        colvir_utils.type_keys(window=change_win, keystrokes="{ENTER}{RIGHT 2}")

        time.sleep(1)
        colvir_utils.type_keys(window=change_win, keystrokes="{ENTER}")
        time.sleep(1)
        change_win.type_keys("^{ENTER}")
        currency_win = colvir_utils.get_window(
            app=app, title="Валюты", wait_for="exists enabled"
        )
        colvir_utils.type_keys(
            window=currency_win, keystrokes="Z", step_delay=0.5
        )
        find_win = colvir_utils.get_window(
            app=app, title="Найти ", wait_for="exists enabled"
        )
        find_win["Edit2"].set_text(row.currency)
        find_win["OK"].click()
        currency_win["OK"].click()

        if "с ндс" in row.name.lower():
            colvir_utils.type_keys(
                window=change_win, keystrokes="{RIGHT 2}{ENTER}"
            )
            change_win.type_keys("^{ENTER}")
            nds_win = colvir_utils.get_window(
                app=app, title="Ставки НДС", wait_for="exists enabled"
            )
            colvir_utils.type_keys(
                window=nds_win, keystrokes="Z", step_delay=0.5
            )
            find_win = colvir_utils.get_window(
                app=app, title="Найти код", wait_for="exists enabled"
            )
            find_win["Edit2"].set_text("05")
            find_win["OK"].click()
            nds_win["OK"].click()

            colvir_utils.type_keys(
                window=change_win, keystrokes="{RIGHT 3}{ENTER}"
            )
        else:
            colvir_utils.type_keys(
                window=change_win, keystrokes="{RIGHT 5}{ENTER}"
            )

        kbk, budget_type = get_kbk(text=row.name, rk=rk)

        change_win.type_keys("^{ENTER}")
        kbk_win = colvir_utils.get_window(
            app=app, title="Классификатор", wait_for="exists"
        )
        colvir_utils.type_keys(window=kbk_win, keystrokes="{F9}")

        dictionary_win = colvir_utils.get_window(
            app=app, title="Справочник.+", regex=True
        )

        dictionary_win["Edit2"].set_text(budget_type)
        assert dictionary_win["Edit2"].window_text() == budget_type

        dictionary_win["Edit4"].set_text(kbk)
        assert dictionary_win["Edit4"].window_text() == kbk

        dictionary_win["OK"].click()

        result_win = colvir_utils.get_window(
            app=app, title="Бюджетная классификация.+", regex=True
        )
        result_win["OK"].click()

        colvir_utils.type_keys(window=change_win, keystrokes="{RIGHT}{ENTER}")
        change_win.type_keys("^{ENTER}")
        branches_win = colvir_utils.get_window(app=app, title="Подразделения")
        colvir_utils.type_keys(window=branches_win, keystrokes="{F7}")
        find_win = colvir_utils.get_window(app=app, title="Поиск")
        find_win["Edit2"].set_text('001. АО "Банк Развития Казахстана"')
        find_win["OK"].click()
        result_win = colvir_utils.get_window(app=app, title="Результаты поиска")
        result_win["Перейти"].click()
        branches_win["OK"].click()

        colvir_utils.type_keys(window=change_win, keystrokes="{RIGHT}{ENTER}")
        change_win.type_keys("^{ENTER}")
        debt_win = colvir_utils.get_window(
            app=app, title="Виды дебиторской.+", regex=True
        )
        colvir_utils.type_keys(window=debt_win, keystrokes="{F9}")
        filter_win = colvir_utils.get_window(app=app, title="Фильтр")
        filter_win["Edit8"].set_text(row.debt_type)
        sleep(1)
        filter_win["OK"].click_input()

        debt_win.wait(wait_for="active enabled")
        debt_win["OK"].click_input()

        colvir_utils.find_and_click_button(
            app, change_win, change_win["Static4"], "Сохранить изменения (PgDn)"
        )
        change_win.wait(wait_for="enabled")

    change_win["OK"].click()

    colvir_utils.find_and_click_button(
        app=app,
        window=business_trip_order_win,
        toolbar=business_trip_order_win["Static3"],
        target_button_name="Авансовый отчет",
    )

    report_win = colvir_utils.get_window(
        app=app, title="Авансовый отчет .+", regex=True
    )

    for _ in request.rows:
        report_win.type_keys("^C")
        current_row_text = pyperclip.paste()

        _, selection = [
            x.replace("\r", "").split("\t")
            for x in current_row_text.split("\n")
        ]
        required_row = next(
            row for row in request.rows if row.name == selection[0]
        )

        colvir_utils.find_and_click_button(
            app=app,
            window=report_win,
            toolbar=report_win["Static3"],
            target_button_name="Создать дочернюю запись",
        )

        colvir_utils.type_keys(window=report_win, keystrokes="{DOWN}{ENTER}")
        report_win.type_keys(required_row.name, with_spaces=True)
        colvir_utils.type_keys(
            window=report_win, keystrokes="{ENTER}{RIGHT 3}{ENTER}"
        )
        report_win.type_keys(required_row.name_num_date, with_spaces=True)
        colvir_utils.type_keys(window=report_win, keystrokes="{ENTER}{RIGHT 4}")
        report_win.type_keys(required_row.sum_tenge, with_spaces=True)
        colvir_utils.type_keys(window=report_win, keystrokes="{ENTER}{RIGHT}")
        if "с ндс" in required_row.name.lower():
            colvir_utils.type_keys(
                window=report_win, keystrokes="{ENTER}{SPACE}{ENTER}{RIGHT}"
            )
            colvir_utils.type_keys(window=report_win, keystrokes="12")
            colvir_utils.type_keys(
                window=report_win, keystrokes="{ENTER}{LEFT}"
            )

        colvir_utils.find_and_click_button(
            app=app,
            window=report_win,
            toolbar=report_win["Static3"],
            target_button_name="Сохранить изменения (PgDn)",
        )

        colvir_utils.type_keys(window=report_win, keystrokes="{UP}")

    report_win.close()

    business_trip_order_win.set_focus()
    time.sleep(1)
    business_trip_order_win.menu_select("#0->#5->#4")
    confirm_pay_win = colvir_utils.get_window(
        app=app, title="Подтверждение", wait_for="exists enabled"
    )
    confirm_pay_win["&Да"].click()

    approve_win = colvir_utils.get_window(
        app=app, title="Утвердить авансовый отчет", wait_for="exists enabled"
    )
    approve_win["Edit2"].set_text(now.strftime("%d.%m.%y"))
    assert approve_win["Edit2"].window_text() == now.strftime("%d.%m.%y")
    approve_win["OK"].click()

    time.sleep(2)
    wiggle_mouse(1)

    business_trip_order_win.menu_select("#0->#5->#2")
    confirm_accounting_win = colvir_utils.get_window(
        app=app, title="Подтверждение", wait_for="exists enabled"
    )
    confirm_accounting_win["&Да"].click()

    time.sleep(5)
    wiggle_mouse(1)

    error_win = app.window(title_re="Произошла ошибка")
    if error_win.exists():
        # make a screenshot
        raise Exception("")

    colvir_utils.find_and_click_button(
        app=app,
        window=business_trip_order_win,
        toolbar=business_trip_order_win["Static3"],
        target_button_name="Журнал выполненных операций",
    )

    # journal_win = colvir_utils.get_window(app=app, title="Журнал операций")
    # if request.reimbursement:
    #     new_finance(app=app, journal_win=journal_win, request=request)
    # journal_win.wait(wait_for="enabled")
    # journal_win.close()

    return business_trip_order_win["Edit46"].window_text().capitalize()


def get_from_env(key: str) -> str:
    value = os.getenv(key)
    assert isinstance(value, str), f"{key} not set in .env"
    return value


def parse_name(business_trip_order_win: pywinauto.WindowSpecification) -> str:
    full_name = business_trip_order_win["Edit18"].window_text()
    names = full_name.split(" ")
    name = names[0] + " " + names[1][0] + "."
    if len(names) > 2:
        name += names[2][0] + "."
    return name


def main(bot: TelegramAPI):
    warnings.simplefilter(action="ignore", category=UserWarning)
    dotenv.load_dotenv()

    project_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    setup_logger(project_folder=project_folder)

    now = datetime.now()
    logging.info("Start of the process...")

    logging.info(f"{bot=}")

    driver_path = get_from_env("DRIVER_PATH")
    bpm_user = get_from_env("BPM_USER")
    bpm_password = get_from_env("BPM_PASSWORD")
    colvir_path = get_from_env("COLVIR_PATH")
    colvir_user = get_from_env("COLVIR_USER")
    colvir_password = get_from_env("COLVIR_PASSWORD")

    logging.info(f"{driver_path=}")
    logging.info(f"{bpm_user=} {bpm_password=}")
    logging.info(f"{colvir_path=} {colvir_user=} {colvir_password=}")

    data_folder = os.path.join(project_folder, "data")
    attachment_folder_path = os.path.join(data_folder, "attachments")

    os.makedirs(data_folder, exist_ok=True)
    os.makedirs(attachment_folder_path, exist_ok=True)

    sample_json_path = os.path.join(data_folder, "sample.json")
    report_path = os.path.join(attachment_folder_path, "Отчет.xlsx")

    process_utils.kill_all_processes(proc_name="COLVIR")

    bpm.run(
        executable_path=driver_path,
        bpm_user=bpm_user,
        bpm_password=bpm_password,
        sample_json_path=sample_json_path,
    )

    requests = data.load_json_requests(sample_json_path)

    logging.info(f"{requests=}")

    colvir = colvir_utils.Colvir(
        process_path=colvir_path, user=colvir_user, password=colvir_password
    )
    app = colvir.get_app()
    colvir_utils.choose_mode(app=app, mode="KREQDOC")

    report_data = []
    for request in requests:
        order_report = {
            "№ Приказа": request.order_id,
            "Статус": "",
            "Отработан роботом": "",
        }

        fill_filter_win(
            app=app, year=now.strftime("%y"), order_id=request.order_id
        )

        confirm_order_not_exists_win = app.window(title="Подтверждение")
        if confirm_order_not_exists_win.exists():
            confirm_order_not_exists_win["&Нет"].click()
            order_report["Статус"] = "Приказ не найден"
            order_report["Отработан роботом"] = "Нет. Приказ не найден"
            report_data.append(order_report)
            logging.info(f"{order_report=}")
            continue

        main_win = colvir_utils.get_window(
            app=app, title="Список счетов к оплате", wait_for="exists enabled"
        )
        colvir_utils.type_keys(window=main_win, keystrokes="{ENTER}")

        business_trip_order_win = colvir_utils.get_window(
            app=app, title="Распоряжение на командировку.+", regex=True
        )

        status = business_trip_order_win["Edit46"].window_text().capitalize()
        if status.lower() != "введен":
            order_report["Статус"] = status
            order_report["Отработан роботом"] = (
                "Нет. Приказ уже был отработан днями раньше, либо статус не равен "
                '"Введен"'
            )
            report_data.append(order_report)
            business_trip_order_win.close()
            main_win.close()
            colvir_utils.choose_mode(app=app, mode="KREQDOC")
            logging.info(f"{order_report=}")
            continue

        if request.reimbursement:
            request.reimbursement.name = parse_name(business_trip_order_win)

        status = fill_order(
            app=app,
            business_trip_order_win=business_trip_order_win,
            now=now,
            request=request,
            rk=request.rk,
        )
        order_report["Статус"] = status
        order_report["Отработан роботом"] = "Да"
        report_data.append(order_report)

        business_trip_order_win.close()
        main_win.close()
        colvir_utils.choose_mode(app=app, mode="KREQDOC")
        logging.info(f"{order_report=}")

    logging.info(f"{report_data=}")
    df = pd.DataFrame(report_data)
    df.to_excel(report_path, index=False)

    process_utils.kill_all_processes("COLVIR")

    send_mail(
        subject='Отчет "Учет командировочных"',
        body='Отчет "Учет командировочных"',
        attachment_folder_path=attachment_folder_path,
    )

    logging.info("Finished")


if __name__ == "__main__":
    telegram_bot = TelegramAPI()
    main(bot=telegram_bot)
