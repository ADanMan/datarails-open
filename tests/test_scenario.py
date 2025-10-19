from app import scenario


def test_apply_adjustments_percentage():
    rows = [
        ("2024-01", "Sales", "Revenue", 1000.0, "USD", ""),
        ("2024-01", "Marketing", "Revenue", 500.0, "USD", ""),
    ]

    adjustments = [scenario.ScenarioAdjustment(department="Sales", percentage_change=0.1)]

    result = scenario.apply_adjustments(rows, adjustments)

    assert result[0][3] == 1100.0
    assert result[1][3] == 500.0
