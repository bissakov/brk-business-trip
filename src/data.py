import dataclasses
import json
import logging
import re
from datetime import datetime
from typing import Dict, List, Optional, Union, cast

JSON_Row = Dict[str, str]
JSON_Rows = List[JSON_Row]
JSON_Reimbursement = Dict[str, str]
JSON_Request = Dict[str, Union[str, bool, JSON_Reimbursement, JSON_Rows]]


@dataclasses.dataclass
class Row:
    name: str
    name_num_date: str
    sum_tenge: str
    currency: str
    debt_type: str


@dataclasses.dataclass
class Reimbursement:
    name: str = "{name}"
    city: str = "{city}"
    start_date: str = "{start_date}"
    end_date: str = "{end_date}"
    order_id: str = "{order_id}"
    order_date: str = "{order_date}"

    def __str__(self):
        template = (
            "{name} К удержанию остаток неиспольз. аванса по командировке в {city} согл. отч. о понес. расх. за {"
            "start_date}-{end_date} по приказу № {order_id} от {order_date}"
        )
        return template.format(
            name=self.name,
            city=self.city,
            start_date=self.start_date,
            end_date=self.end_date,
            order_id=self.order_id,
            order_date=self.order_date,
        )

    def __repr__(self):
        return str(self)


@dataclasses.dataclass
class Request:
    order_id: str
    rk: bool
    ob: str
    ppz: bool
    oz: str
    order_type: str
    reimbursement: Optional[Reimbursement]
    rows: List[Row]


def is_num(num_str: str) -> bool:
    try:
        num = float(num_str.replace(" ", ""))
        return isinstance(num, float)
    except ValueError:
        return False


def is_dt_format_correct(date_str: str, date_fmt: str) -> bool:
    try:
        datetime.strptime(date_str, date_fmt)
        return True
    except ValueError:
        return False


def parse_request(json_request: JSON_Request) -> Optional[Request]:
    # NOTE: № Приказа
    order_id = json_request.get("order_id")
    if (
        order_id is None
        or not isinstance(order_id, str)
        or not re.fullmatch(r"\d+\s?-\s?I", order_id)
    ):
        logging.error(f"Error. order_id mismatch: {order_id}")
        return None

    # NOTE: За пределами РК (наоборот)
    rk = json_request.get("rk")
    if rk is None or not isinstance(rk, bool):
        logging.error(f"Error. rk mismatch: {rk}")
        return None

    # NOTE: Оплачено Банком и/или с корпоративной карты
    ob = json_request.get("ob")
    if ob is None or not isinstance(ob, str) or not is_num(ob):
        logging.error(f"Error. ob mismatch: {ob}")
        return None

    # NOTE: Получено по заявке на денежный аванс (существует ли поле в приказе)
    ppz = json_request.get("ppz")
    if ppz is None or not isinstance(ppz, bool):
        logging.error(f"Error. ppz mismatch: {ppz}")
        return None

    # NOTE: Остаток задолженности (+)/Перерасход (-)
    oz = json_request.get("oz")
    if oz is None or not isinstance(oz, str) or not is_num(oz):
        logging.error(f"Error. oz mismatch: {oz}")
        return None

    # NOTE: Вид заявки
    order_type = json_request.get("order_type")
    if order_type is None or not isinstance(order_type, str):
        logging.error(f"Error. order_type mismatch: {order_type}")
        return None

    request = Request(
        order_id=order_id,
        rk=rk,
        ob=ob,
        ppz=ppz,
        oz=oz,
        order_type=order_type,
        reimbursement=None,
        rows=[],
    )

    json_reimbursement = json_request.get("reimbursement")

    if json_reimbursement:
        json_reimbursement = cast(JSON_Reimbursement, json_reimbursement)

        # NOTE: Имя сотрудника
        name = json_reimbursement.get("names")
        if name is None or not isinstance(name, str) or name != "{name}":
            logging.error(f"Error. name mismatch: {name}")
            return None

        # NOTE: Место командирования/обучения
        city = json_reimbursement.get("city")
        if city is None or not isinstance(city, str):
            logging.error(f"Error. city mismatch: {city}")
            return None

        # NOTE: Дата начала
        start_date = json_reimbursement.get("start_date")
        if (
            start_date is None
            or not isinstance(start_date, str)
            or not is_dt_format_correct(start_date, "%d.%m.%Y")
        ):
            logging.error(f"Error. start_date mismatch: {start_date}")
            return None

        # NOTE: Дата окончания
        end_date = json_reimbursement.get("end_date")
        if (
            end_date is None
            or not isinstance(end_date, str)
            or not is_dt_format_correct(end_date, "%d.%m.%Y")
        ):
            logging.error(f"Error. end_date mismatch: {end_date}")
            return None

        # NOTE: Дата подписания
        order_date = json_reimbursement.get("order_date")
        if (
            order_date is None
            or not isinstance(order_date, str)
            or not is_dt_format_correct(order_date, "%d.%m.%Y")
        ):
            logging.error(f"Error. order_date mismatch: {order_date}")
            return None

        request.reimbursement = Reimbursement(
            city=city,
            start_date=start_date,
            end_date=end_date,
            order_id=request.order_id,
            order_date=order_date,
        )

    json_rows = json_request.get("rows")
    if json_rows is None or not isinstance(json_rows, list):
        logging.error(f"Error. rows mismatch: {json_rows}")
        return None

    for json_row in json_rows:
        json_row = cast(JSON_Row, json_row)

        # NOTE: Наименование расхода
        name = json_row.get("name")
        if name is None or not isinstance(name, str):
            logging.error(f"Error. row name mismatch: {name}")
            return None

        # NOTE: Наименование, №, дата подтверждающего документа
        name_num_date = json_row.get("name_num_date")
        if name_num_date is None or not isinstance(name_num_date, str):
            logging.error(f"Error. row name_num_date mismatch: {name_num_date}")
            return None

        # NOTE: Сумма расходов в тенге
        sum_tenge = json_row.get("sum_tenge")
        if (
            sum_tenge is None
            or not isinstance(sum_tenge, str)
            or not is_num(sum_tenge)
        ):
            logging.error(f"Error. row sum_tenge mismatch: {sum_tenge}")
            return None

        # NOTE: Валюта
        currency = json_row.get("currency")
        if (
            currency is None
            or not isinstance(currency, str)
            or currency != "KZT"
        ):
            logging.error(f"Error. row currency mismatch: {currency}")
            return None

        # NOTE: Тип задолженности
        debt_type = json_row.get("debt_type")
        if debt_type is None or not isinstance(debt_type, str):
            logging.error(f"Error. row debt_type mismatch: {debt_type}")
            return None

        row = Row(
            name=name,
            name_num_date=name_num_date,
            sum_tenge=sum_tenge,
            currency=currency,
            debt_type=debt_type,
        )
        request.rows.append(row)

    return request


def load_json_requests(sample_json_path: str) -> List[Request]:
    requests = []
    with open(sample_json_path, "r", encoding="utf-8") as f:
        sample_json = json.load(f)

    for json_request in sample_json:
        json_request = cast(JSON_Request, json_request)
        request = parse_request(json_request)
        requests.append(request)

    return requests
