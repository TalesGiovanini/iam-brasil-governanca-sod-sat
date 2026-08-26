from typing import Protocol


class AIProvider(Protocol):
    name: str

    def generate(self, prompt: str, **kwargs: object) -> str:
        ...
