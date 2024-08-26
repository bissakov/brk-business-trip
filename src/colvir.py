import os
import warnings
from time import sleep

import dotenv

from src import colvir_utils, data, process_utils
from src.main import fill_filter_win, get_from_env


def main():
    warnings.simplefilter(action="ignore", category=UserWarning)
    dotenv.load_dotenv()

    process_utils.kill_all_processes(proc_name="COLVIR")

    colvir_path = get_from_env("COLVIR_PATH")
    colvir_user = get_from_env("COLVIR_USER")
    colvir_password = get_from_env("COLVIR_PASSWORD")

    project_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_folder = os.path.join(project_folder, "data")

    sample_json_path = os.path.join(data_folder, "sample.json")
    requests = data.load_json_requests(sample_json_path)
    request = requests[0]

    if not request.reimbursement:
        return

    request.reimbursement.name = "Капенов Т.Н."

    colvir = colvir_utils.Colvir(
        process_path=colvir_path, user=colvir_user, password=colvir_password
    )
    app = colvir.get_app()
    colvir_utils.choose_mode(app=app, mode="KREQDOC")

    fill_filter_win(app=app, year="24", order_id="410 - I")

    main_win = colvir_utils.get_window(
        app=app, title="Список счетов к оплате", wait_for="exists enabled"
    )
    colvir_utils.type_keys(window=main_win, keystrokes="{ENTER}")

    business_trip_order_win = colvir_utils.get_window(
        app=app, title="Распоряжение на командировку.+", regex=True
    )

    colvir_utils.find_and_click_button(
        app=app,
        window=business_trip_order_win,
        toolbar=business_trip_order_win["Static3"],
        target_button_name="Журнал выполненных операций",
    )

    journal_win = colvir_utils.get_window(app=app, title="Журнал операций")

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

    # finance_win.menu_select("#0->#1")

    # colvir_utils.find_and_click_button(
    #     app=app,
    #     window=finance_win,
    #     toolbar=finance_win["Static3"],
    #     target_button_name="Сохранить изменения (PgDn)",
    # )
    journal_win.close()


if __name__ == "__main__":
    main()
