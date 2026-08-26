from dataclasses import dataclass


@dataclass(frozen=True)
class SystemStatus:
    core: str = "ONLINE"
    agent_engine: str = "BOOTSTRAP"
    knowledge_base: str = "BOOTSTRAP"
    inference_enabled: bool = False
    inference_provider: str = "null"


def get_status() -> SystemStatus:
    return SystemStatus()
