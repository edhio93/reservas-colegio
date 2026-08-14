from pathlib import Path


def test_core_structure_exists():
    root = Path(__file__).resolve().parents[1]
    required = [
        "schedule_app.py",
        "core/config.py",
        "core/permissions.py",
        "core/navigation.py",
        "services/supabase.py",
        "services/gemini.py",
        "repositories/reservas.py",
        "components/responsive.py",
    ]
    missing = [item for item in required if not (root / item).exists()]
    assert not missing, f"Faltan archivos: {missing}"
