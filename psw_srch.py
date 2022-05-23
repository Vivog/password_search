import itertools
from string import digits, punctuation, ascii_letters
from datetime import datetime
import time
import win32com.client as client

def pasw_s():
    try:
        print('Привет пользователь!\nУкажи из скольки символов должен состоять пароль\n'
          'Например: 2-4 : ', end='')
        psw_length = input()
        psw_length = [int(item.strip()) for item in psw_length.split('-')]
    except:
        print('Не вверно ввел что-то...')
    combinations = 1
    try:
        print('Укажи какие символы должен содержать пароль\n'
              'Если только цифры, то введи 1\n'
              'Если только буквы, то введи 2\n'
              'Если только цифры и буквы, то введи 3\n'
              'Если цифры, буквы и спецсимволы, то введи 4')
    except:
        print('Не вверно ввел что-то...')
    try:
        psw_type = int(input('Твой ответ: '))
        if psw_type == 1:
            print('Пароль может состоять из этого набора символов:')
            possible_symbols = digits
        elif psw_type == 2:
            print('Пароль может состоять из этого набора символов:')
            possible_symbols = ascii_letters
        elif psw_type == 3:
            print('Пароль может состоять из этого набора символов:')
            possible_symbols = digits + ascii_letters
        elif psw_type == 4:
            print('Пароль может состоять из этого набора символов:')
            possible_symbols = digits + ascii_letters + punctuation
        else:
            print('Не вверно ввел что-то...')
        print(possible_symbols, len(possible_symbols))
        if psw_length[0] == psw_length[1]:
            combinations = pow(len(possible_symbols),psw_length[0])
        else :
            for i in psw_length:
                combinations *= pow(len(possible_symbols), i)
    except:
        print('Не вверно ввел что-то...')
    return (psw_length, possible_symbols)

#     iter password
pasw_length, possible_symbols = pasw_s()

def search_pasw(psw_length=pasw_length, possible_symbols=possible_symbols):
    start_search = time.time()
    print("Старт поиска пароля ", datetime.fromtimestamp(start_search).strftime('%H:%M:%S'))
    atempt_count = 0
    for psw_length in range(psw_length[0], psw_length[1]+1):
        for password in itertools.product(possible_symbols, repeat=psw_length):
            password = ''.join(password)
            atempt_count += 1
            if open_file(password, atempt_count, start_search):
                continue
            else:
                break

def open_file(password, atempt_count, start_search):
    open_app = client.Dispatch("Excel.Application")
    try:
        work = open_app.Workbooks.Open(
            r'C:\Users\VIVOG\PycharmProjects\password_search\251.xlsx',
            False,
            True,
            None,
            password)
        print("Пароль найден ", datetime.fromtimestamp(time.time()).strftime('%H:%M:%S'))
        print("Затрачено времени: ", time.time() - start_search)
        work.Close()
        open_app.Quit()
        print(f'Пароль - {password}, количество попыток - {atempt_count}')
        return 0
    except:
        return 1
def main():
    search_pasw()
if __name__ == "__main__":
    main()