import dop
import os

def main():
    bibr = input('''
    [1] Муниципальный этап
    [2] Ставропльский край
    [3] Ростовская областьфизика
    [4] Краснодарский край
    [5] Республика Дагестан
    [6] Республика Северная Осетия
    [7] Республика Карачаево-Черкессия
    [8] Республика Калмыкия
    [9] Чеченская республика
    [10] Анализ
    [0] Выход
    Введите значение:''')

    if bibr == "1":
        dop.Myn()
    elif bibr == "2":
        dop.Stav()
    elif bibr == "3": 
        dop.Rost()   
    elif bibr == "4": 
        dop.Kras()    
    elif bibr == "5": 
        dop.Dages()
    elif bibr == "6": 
        dop.Sever()
    elif bibr == "7":    
        dop.Kara()
    elif bibr == "8": 
        dop.Kalm()
    elif bibr == "9":
        dop.Cheche() 
    elif bibr == "10":
        dop.Analiz()
    elif bibr == "0":
        quit('выход')
    else:
        print("Такого значения нет")
        clear = lambda: os.system('cls')
        clear()
        main()

main()