import pytest
from calendar_utils import first_day_of_the_month, get_week_days, shade_cells_if_needed
from docx import Document
import datetime

# Test wyznaczania pierwszego dnia miesiąca
def test_first_day_of_the_month():
    assert first_day_of_the_month('Styczeń', 2024) == 'Poniedziałek'
    assert first_day_of_the_month('Luty', 2024) == 'Czwartek'
    assert first_day_of_the_month('Marzec', 2023) == 'Środa'
    assert first_day_of_the_month('Grudzień', 2023) == 'Piątek'

# Test poprawnego tworzenia listy dni tygodnia
def test_get_week_days():
    assert get_week_days('Poniedziałek', 7) == ['Poniedziałek', 'Wtorek', 'Środa', 'Czwartek', 'Piątek', 'Sobota', 'Niedziela']
    assert get_week_days('Sobota', 3) == ['Sobota', 'Niedziela', 'Poniedziałek']

# Test kolorowania odpowiednich komórek w tabeli
def test_shade_cells_if_needed():
    doc = Document()
    table = doc.add_table(rows=2, cols=3)  # 2 wiersze, 3 kolumny (Dzień miesiąca, Dzień tygodnia, Ciśnienie)
    
    # Wypełnij wiersze przykładowymi danymi
    row = table.rows[0]
    row.cells[0].text = "1"
    row.cells[1].text = "Sobota"
    row.cells[2].text = "Ciśnienie"

    # Sprawdzenie, czy wiersz dla soboty (weekendu) jest szary
    shade_cells_if_needed(row, 'Sobota', False)
    for cell in row.cells:
        assert 'D9D9D9' in cell._element.xml  # Kolor szary dla weekendu

    # Sprawdzenie dla dnia powszedniego (np. Wtorek)
    row = table.rows[1]
    row.cells[0].text = "6"
    row.cells[1].text = "Wtorek"
    row.cells[2].text = "Ciśnienie"
    shade_cells_if_needed(row, 'Wtorek', False)
    
    for cell in row.cells:
        assert 'D9D9D9' not in cell._element.xml  # Brak koloru szarego dla dnia powszedniego

    # Sprawdzenie kolorowania dla dnia specjalnego (11 listopada 2024 to poniedziałek)
    row = table.add_row()
    row.cells[0].text = "11"
    row.cells[1].text = "Poniedziałek"
    row.cells[2].text = "Ciśnienie"
    date = datetime.date(2024, 11, 11)
    shade_cells_if_needed(row, 'Poniedziałek', date)
    
    for cell in row.cells:
        assert 'D9D9D9' in cell._element.xml  # Kolor szary dla dnia specjalnego

if __name__ == '__main__':
    pytest.main()
