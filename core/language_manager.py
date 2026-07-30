import json
import os


class LanguageManager:

    def __init__(self, language_code="en"):
        self.language_code = language_code
        self.translations = {}
        self.load_language(language_code)

    def load_language(self, language_code):

        file_path = os.path.join(
            "languages",
            f"{language_code}.json"
        )

        with open(file_path, "r", encoding="utf-8") as file:
            self.translations = json.load(file)

        self.language_code = language_code

    def get(self, key):

        return self.translations.get(key, key)