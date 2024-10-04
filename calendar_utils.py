
import datetime
from docx.oxml import OxmlElement
from docx.oxml.ns import qn



month_days = {
    'Styczeń': 31,
    'Luty': 28,
    'Marzec': 31,
    'Kwiecień': 30,
    'Maj': 31,
    'Czerwiec': 30,
    'Lipiec': 31,
    'Sierpień': 31,
    'Wrzesień': 30,
    'Październik': 31,
    'Listopad': 30,
    'Grudzień': 31
}


def first_day_of_the_month(mon, year):
    """
    Zwraca pierwszy dzień tygodnia dla danego miesiąca i roku.
    """
    months = {
        'Styczeń': 1,
        'Luty': 2,
        'Marzec': 3,
        'Kwiecień': 4,
        'Maj': 5,
        'Czerwiec': 6,
        'Lipiec': 7,
        'Sierpień': 8,
        'Wrzesień': 9,
        'Październik': 10,
        'Listopad': 11,
        'Grudzień': 12
    }
    month_number = months[mon]
    data = datetime.date(year, month_number, 1)
    day_of_week = data.weekday()  # 0 = poniedziałek, 6 = niedziela
    days_of_week = ['Poniedziałek', 'Wtorek', 'Środa', 'Czwartek', 'Piątek', 'Sobota', 'Niedziela']
    return days_of_week[day_of_week]

def get_week_days(first_day, no_of_rows):
    """
    Zwraca listę dni tygodnia na podstawie pierwszego dnia miesiąca.
    """
    days_of_week = ['Poniedziałek', 'Wtorek', 'Środa', 'Czwartek', 'Piątek', 'Sobota', 'Niedziela']
    start = days_of_week.index(first_day)
    res = []
    for i in range(no_of_rows):  # minus 2 bo pierwszy wiersz to nagłówek
        index = (start + i) % len(days_of_week)
        res.append(days_of_week[index])
    return res

def shade_cells_if_needed(row, week_day, is_special_day):
    """
    Koloruje odpowiednie komórki, jeśli dzień to weekend lub dzień specjalny.
    """
    is_weekend = week_day.upper() in ['SOBOTA', 'NIEDZIELA']
    
    if is_weekend or is_special_day:
        for cell in row.cells:
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), 'D9D9D9')
            cell._element.get_or_add_tcPr().append(shading_elm)
