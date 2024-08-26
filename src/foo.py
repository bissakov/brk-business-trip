import os

import dotenv

from src.mail import send_mail


def main():
    dotenv.load_dotenv()

    project_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    data_folder = os.path.join(project_folder, "data")
    attachment_folder_path = os.path.join(data_folder, "attachments")

    send_mail(
        subject="asd mail",
        body="test",
        attachment_folder_path=attachment_folder_path,
    )


if __name__ == "__main__":
    main()
