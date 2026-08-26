from tales.core.engine import get_status


def test_status_ai_is_disabled_by_default():
    status = get_status()
    assert status.core == "ONLINE"
    assert status.ai_enabled is False
    assert status.ai_provider == "disabled"
