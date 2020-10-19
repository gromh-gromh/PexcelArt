#Ебало опусти

from PIL import Image
import numpy as np
import xlsxwriter
import os

# Путь к изображению
inputDirectory = input('Укажите путь к исходному изображению: ')
extension = os.path.splitext(inputDirectory)[1]
while os.path.isfile(inputDirectory) == 0:
    inputDirectory = input('Изображение не найдено, укажите путь еще раз: ')
else:
    img = Image.open(inputDirectory)

# Выбор режима
if extension != ".png":
    print('Выберите режим\n[0] Фото режим (сохранение полутонов)\n[1] Минималистичный режим\n[2] Ч/Б')
    mode = int(input('Введите число: '))
    while mode != 0 and mode != 1 and mode != 2:
        mode = int(input('Неправильное значение, повторите ввод: '))
else:
    print('Выберите режим\n[1] Минималистичный режим\n[2] Ч/Б')
    mode = int(input('Введите число: '))
    while mode != 1 and mode != 2:
        mode = int(input('Неправильное значение, повторите ввод: '))

# Опциональная конвертация цветов для минималистичного режима
if mode == 1:
    inputColors = int(input('Укажите количество цветов в палитре конечного изображения (не более 256): '))
    while inputColors > 256 or inputColors <= 0:
        inputColors = int(input('Неверное количество цветов, укажите еще раз: '))
    else:
        img = img.convert('P', palette=Image.ADAPTIVE, colors=inputColors)

# Опциональная конвертация цветов для Ч/Б режима
if mode == 2:
    gray = img.convert('L')
    if extension != ".png":
        img = gray.point(lambda x: 0 if x < 128 else 256, '1')
    else:
        img = gray.point(lambda x: 0 if 128 > x > 0 else 256, '1')

# Компрессия
thumbnail = int(input('Укажите ограничение размера большей стороны конечного изображения (в пикселях): '))
size = thumbnail, thumbnail
img.thumbnail(size)

# Запись в массив
if mode == 1 or mode == 2:
    img = img.convert('RGB', palette=Image.ADAPTIVE, colors=255)
img_array = np.array(img)

# Создание xlsx документа
outWorkBook = xlsxwriter.Workbook("Pexcelart.xlsx")
outSheet = outWorkBook.add_worksheet()

# Попиксельная запись изображения в документ
i = -1
j = -1
for r in img_array:
    j = -1
    i += 1
    for c in r:
        j += 1

        outSheet.set_column(j, j, 2.14)  # Форматирование ячеек в квадраты
        if np.any(img_array[i, j]) != 0:
            color = '#%02X%02X%02X' % tuple(img_array[i, j])  # Конвертация RGB в HTML
            cell_format = outWorkBook.add_format()
            cell_format.set_bg_color(color)
            outSheet.write(i, j, '', cell_format)
        else:
            if mode == 2:
                cell_format = outWorkBook.add_format()
                cell_format.set_bg_color('black')
                outSheet.write(i, j, '', cell_format)

outWorkBook.close()
outLocation = os.getcwd()
print('Изображение успешно созданно в директории', outLocation, 'нажмите enter чтобы выйти')
input()
