from pathlib import Path
from tales.core.policies import ensure_within_roots


def list_files(target: str, allowed_roots: list[str]) -> list[str]:
    path = ensure_within_roots(Path(target), [Path(p) for p in allowed_roots])
    if not path.exists():
        raise FileNotFoundError(path)
    if path.is_file():
        return [str(path)]
    return [str(p) for p in sorted(path.rglob("*")) if p.is_file()]
