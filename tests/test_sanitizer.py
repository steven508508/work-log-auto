from daily_sync import sanitize

def test_masks_email_address():
    assert "[Contact]" in sanitize("Email john@example.com")

def test_sensitive_keyword_becomes_internal_task():
    assert sanitize("Confidential HR Review") == " Internal Task"

def test_project_mapping():
    assert sanitize("Project DeathStar meeting") == "Infrastructure Upgrade meeting"
