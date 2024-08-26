import dataclasses
import json
import logging
import time
from typing import Dict, List, Optional, Union

import selenium.webdriver.chrome.service as chrome_service
from selenium.common import NoSuchElementException
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.wait import WebDriverWait

import src.data as data

Sample = Dict[str, List[Union[str, int, List[List[str]]]]]


def driver_init(executable_path: str) -> Chrome:
    service = chrome_service.Service(executable_path=executable_path)
    options = ChromeOptions()
    prefs = {"profile.default_content_setting_values.notifications": 2}
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")
    driver = Chrome(service=service, options=options)
    return driver


def login(
    driver: Chrome, wait: WebDriverWait, bpm_user: str, bpm_password
) -> None:
    driver.get("https://bpm.kdb.kz/?s=obj_a&gid=873&reset_page=1")

    user_input = wait.until(
        ec.presence_of_element_located((By.NAME, "u_login"))
    )
    user_input.send_keys(bpm_user)

    psw_input = wait.until(ec.presence_of_element_located((By.NAME, "pwd")))
    psw_input.send_keys(bpm_password)

    submit_button = wait.until(
        ec.presence_of_element_located((By.NAME, "submit"))
    )
    submit_button.click()


def parse_table(driver: Chrome) -> List[Dict[str, str]]:
    table = driver.find_element(
        By.CSS_SELECTOR, ".udf_box_content.udf_box_content_84661 .obj_table"
    )

    headers_element = table.find_element(
        By.CSS_SELECTOR, "tr.obj_tbl_header:not(.js_hidden)"
    )
    headers = [header for header in headers_element.text.split("\n") if header]

    cells = [
        cell.text.strip()
        for cell in table.find_elements(
            By.CSS_SELECTOR, "tr[data-row] > td > .obj_table_value"
        )
    ]

    rows = []
    for i in range(0, len(cells), len(headers)):
        row = dict(zip(headers, cells[i : i + len(headers)]))
        if not row.get("Наименование расхода") and not row.get(
            "Наименование, №, дата подтверждающего документа"
        ):
            continue
        rows.append(row)
    return rows


def fill_reimbursement(
    driver: Chrome, request: data.Request
) -> Optional[data.Reimbursement]:
    oz_num = float(request.oz.replace(" ", ""))

    if oz_num <= 0:
        return None

    city = driver.find_element(
        By.CSS_SELECTOR,
        '[data-field-label="Место командирования/обучения"] [class="field_view"]',
    ).text

    start_date = driver.find_element(
        By.CSS_SELECTOR, '[data-field-label="Дата начала"] [class="field_view"]'
    ).text

    end_date = driver.find_element(
        By.CSS_SELECTOR,
        '[data-field-label="Дата окончания"] [class="field_view"]',
    ).text

    order_date = driver.find_element(
        By.CSS_SELECTOR,
        '[data-field-label="Дата подписания"] [class="field_view"]',
    ).text

    reimbursement = data.Reimbursement(
        city=city,
        start_date=start_date,
        end_date=end_date,
        order_id=request.order_id,
        order_date=order_date,
    )

    return reimbursement


def find_element(
    parent: Union[Chrome, WebElement],
    by: str,
    value: str,
    default: Optional[str] = None,
) -> str:
    try:
        element = parent.find_element(by, value)
        return element.text
    except NoSuchElementException as error:
        logging.exception(error)
        if not default:
            raise error
        return default


def run(
    executable_path: str,
    bpm_user: str,
    bpm_password: str,
    sample_json_path: str,
) -> List[data.Request]:
    driver = driver_init(executable_path)

    requests: List[data.Request] = []

    wait = WebDriverWait(driver, timeout=10)

    with driver:
        login(
            driver=driver,
            wait=wait,
            bpm_user=bpm_user,
            bpm_password=bpm_password,
        )

        state_filter_input = wait.until(
            ec.presence_of_element_located(
                (By.CSS_SELECTOR, '[data-col-id="4680"]')
            )
        )
        state_filter_input.send_keys("На исполнении (НУ ДБУ)")

        time.sleep(3)

        urls = [
            url.get_attribute("href")
            for url in driver.find_elements(
                By.CSS_SELECTOR,
                ".js_dbl_click_text_select.js_list_dflt_col_5 > a",
            )
        ]

        for url in urls:
            if not url:
                continue

            driver.get(url)

            wait.until(
                ec.presence_of_element_located(
                    (By.CSS_SELECTOR, "div.form_table")
                )
            )

            order_id = find_element(
                driver,
                By.CSS_SELECTOR,
                '[data-field-label="№ Приказа"] [class="field_view"]',
            )

            rk = (
                find_element(
                    driver,
                    By.CSS_SELECTOR,
                    '[data-field-label="За пределами РК"] [class="udf_field_el_value"]',
                )
                == "—"
            )

            ob = find_element(
                driver,
                By.CSS_SELECTOR,
                '[data-field-label="Оплачено Банком и/или с корпоративной карты"] [class="udf_field_el_value"]',
                default="0.00",
            )

            ppz = (
                find_element(
                    driver,
                    By.CSS_SELECTOR,
                    '[data-field-label="Получено по заявке на денежный аванс"] [class="udf_field_el_value"]',
                    default="0.00",
                )
                != "0.00"
            )

            oz = find_element(
                driver,
                By.CSS_SELECTOR,
                '[data-field-label="Остаток задолженности (+)/Перерасход (-)"] [class="udf_field_el_value"]',
                default="0.00",
            )

            order_type = find_element(
                driver,
                By.CSS_SELECTOR,
                '[data-field-label="Вид заявки"] [class="udf_field_el_value"]',
            )

            request = data.Request(
                order_id=order_id,
                rk=rk,
                ob=ob,
                ppz=ppz,
                oz=oz,
                order_type=order_type,
                reimbursement=None,
                rows=[],
            )

            request.reimbursement = fill_reimbursement(
                driver=driver, request=request
            )

            rows = parse_table(driver=driver)
            for row in rows:
                request_row = data.Row(
                    name=row["Наименование расхода"],
                    name_num_date=row[
                        "Наименование, №, дата подтверждающего документа"
                    ],
                    sum_tenge=row["Сумма расходов в тенге"],
                    currency=row["Валюта"],
                    debt_type="2",
                )

                name = request_row.name.lower()
                if "проезд" in name or "сервисный" in name:
                    request_row.debt_type = "10"
                elif not request.ppz and (
                    "суточные" in name or "проживание" in name
                ):
                    request_row.debt_type = "39"
                else:
                    request_row.debt_type = "2"

                request.rows.append(request_row)
            requests.append(request)

    with open(sample_json_path, "w", encoding="utf-8") as f:
        json.dump(
            [dataclasses.asdict(request) for request in requests],
            f,
            indent=4,
            ensure_ascii=False,
        )

    return requests
