# spec_days.py

def add_special_days():
    special_days = []
    while True:
        choice = input("Dodaj dzień specjalny T/N: ").upper()
        if choice == 'N':
            break
        elif choice == 'T':
            day = int(input("Dzień specjalny (numer): "))
            special_days.append(day)
    return special_days
