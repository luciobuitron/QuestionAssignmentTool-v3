import os
from datetime import datetime


LOG_FOLDER = "logs"
LOG_FILE = os.path.join(LOG_FOLDER, "qat.log")


def log(level, message):

    os.makedirs(LOG_FOLDER, exist_ok=True)

    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    with open(
        LOG_FILE,
        "a",
        encoding="utf-8"
    ) as file:

        message = str(message).replace("\n", " ").replace("\r", "")

        file.write(
            f"{timestamp} | {level} | {message}\n"
        )