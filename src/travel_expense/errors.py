class ExportError(Exception):
    """利用者に表示できる出力エラー。"""


class ValidationError(ExportError):
    """入力内容が出力条件を満たさない。"""


class TemplateError(ExportError):
    """原本の構造が想定と異なる。"""


class OutputExistsError(ExportError):
    """出力先が既に存在する。"""
