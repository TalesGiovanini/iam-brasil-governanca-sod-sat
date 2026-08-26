from dataclasses import dataclass


@dataclass(frozen=True)
class SystemStatus:
    core: str = "ONLINE"
    agent_engine: str = "BOOTSTRAP"
    knowledge_base: str = "BOOTSTRAP"
    ai_enabled: bool = False
    ai_provider: str = "disabled"


def get_status() -> SystemStatus:
    return SystemStatus()
