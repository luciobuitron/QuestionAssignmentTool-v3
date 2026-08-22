import json

from core.path_utils import resource_path


class LanguageManager:

    def __init__(self, language_code="en"):
        self.language_code = language_code
        self.translations = {}
        self.load_language(language_code)

    def load_language(self, language_code):

        file_path = resource_path(
            f"languages/{language_code}.json"
        )

        with open(file_path, "r", encoding="utf-8") as file:
            self.translations = json.load(file)

        self.language_code = language_code

    def get(self, key):

        return self.translations.get(key, key)