from pathlib import Path


class DataSource:

    @classmethod
    def data_path(cls, book_name) -> Path:
        return Path(__file__).parent.parent.parent.joinpath(book_name)
