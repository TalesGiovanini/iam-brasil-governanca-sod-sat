class DisabledProvider:
    name = "disabled"

    def generate(self, prompt: str, **kwargs: object) -> str:
        raise RuntimeError("AI provider is disabled. This operation requires AI.")
