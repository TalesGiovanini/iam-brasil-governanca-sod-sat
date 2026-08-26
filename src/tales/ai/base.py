from typing import Protocol


class InferenceProvider(Protocol):
    name: str

    def generate(self, prompt: str, **kwargs: object) -> str:
        ...
