"""Файл с кастомными исключениями. """


class EmptyException(Exception):
    def __init__(self, text):
        self.txt = text
