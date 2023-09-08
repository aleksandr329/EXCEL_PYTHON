from constants import *
from function import *

while True:
    print(start)
    start1 = input('Какую операцию хотите выполнить? ')

    if start1.lower() in 'задолжность':
        scan()

    if start1.lower() in 'внести':
        zapis()

    if start1.lower() in 'завершить':
        break

    if start1.lower() in 'организация':
        name = input('Введите название организации: ')
        scan3(name)

    if start1.lower() in 'инн':
        inn = int(input('Введите инн организации: '))
        scan4(inn)

    if start1.lower() in 'номер счета':
        number = input('Введите номер счета: ')
        scan5(number)
