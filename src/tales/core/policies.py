from pathlib import Path


class PolicyViolation(PermissionError):
    pass


def ensure_within_roots(target: Path, roots: list[Path]) -> Path:
    resolved = target.resolve()
    for root in roots:
        root_resolved = root.resolve()
        try:
            resolved.relative_to(root_resolved)
            return resolved
        except ValueError:
            continue
    raise PolicyViolation(f"Path not allowed by policy: {resolved}")
