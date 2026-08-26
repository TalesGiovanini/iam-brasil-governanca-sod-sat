class NullProvider:
    name = "null"

    def generate(self, prompt: str, **kwargs: object) -> str:
        raise RuntimeError("Módulo de inferência não configurado. Esta operação requer um provedor externo.")
