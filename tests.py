import pytest

def test_answer():
    assert OD_df['Total_Journeys'].isna().sum() == 0

