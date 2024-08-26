import logging
import os
import urllib.parse
from typing import Dict, Optional, Tuple

import requests
import requests.adapters


class TelegramAPI:
    def __init__(self) -> None:
        self.session = requests.Session()
        self.session.mount(
            "http://", requests.adapters.HTTPAdapter(max_retries=5)
        )
        self.api_url = "https://api.telegram.org/bot{token}/"

    def reload_session(self) -> None:
        self.session = requests.Session()
        self.session.mount(
            "http://", requests.adapters.HTTPAdapter(max_retries=5)
        )

    def send_message(
        self,
        token: str,
        chat_id: str,
        message: str,
        use_session: bool = True,
    ) -> bool:
        self.api_url = self.api_url.format(token=token)
        send_data: Dict[str, Optional[str]] = {"chat_id": chat_id}
        files = None

        url = urllib.parse.urljoin(self.api_url, "sendMessage")
        send_data["text"] = message

        if use_session:
            response = self.session.post(url, data=send_data, files=files)
        else:
            response = requests.post(url, data=send_data, files=files)

        method = url.split("/")[-1]
        data = "" if not hasattr(response, "json") else response.json()
        logging.info(
            f"Response for '{method}': {response}\n"
            f"Is 200: {response.status_code == 200}\n"
            f"Data: {data}"
        )
        response.raise_for_status()
        return response.status_code == 200


def get_secrets() -> Tuple[str, str]:
    token = os.getenv("TOKEN")
    if token is None:
        raise EnvironmentError("Environment variable 'TOKEN' is not set.")

    chat_id = os.getenv("CHAT_ID")
    if chat_id is None:
        raise EnvironmentError("Environment variable 'CHAT_ID' is not set.")

    return token, chat_id


def send_with_retry(
    bot: TelegramAPI,
    token: str,
    chat_id: str,
    message: str,
) -> bool:
    retry = 0
    while retry < 5:
        try:
            use_session = retry < 5
            success = bot.send_message(token, chat_id, message, use_session)
            return success
        except (
            requests.exceptions.ConnectionError,
            requests.exceptions.SSLError,
            requests.exceptions.HTTPError,
        ) as e:
            bot.reload_session()
            logging.exception(e)
            logging.warning(f"{e} intercepted. Retry {retry + 1}/10")
            retry += 1

    return False


def send_message(
    bot: TelegramAPI,
    message: str,
) -> None:
    token, chat_id = get_secrets()
    send_with_retry(bot, token, chat_id, message)
