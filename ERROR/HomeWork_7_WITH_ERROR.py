import datetime
import csv
import json

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm

#Задание 1. Автоматическая генерация отчета о машине в формате doc
#Ф-ия генерации DOC-файла
def Get_car_info (company, car, car_model, engine_volume, gearbox, price):
    template = 'car.docx' #Файл шаблона
    car_photo = 'car_img.png' # Файл изображения
    # Информация для наполнения
    docunet = from_template (company, car, car_model, engine_volume, gearbox, price, template, car_photo)

#Формирование шаблона
def from_template (company, car, car_model, engine_volume, gearbox, price, template, car_photo):
    template = DocxTemplate(template) #Передаем файл шаблона
    # Передаем информацию для заполнения шаблона
    context = Get_car_context(company, car, car_model, engine_volume, gearbox, price)

    #Вставка картинки
    img_size = Cm(10) #Указываем размер изображения
    car_img = InlineImage(template, car_photo, img_size) #Обрабатываем изображение
    context['car_img'] = car_img #Добавляем изображение в содержимое Информации для шаблона

    #Рендеринг передачи данных в шаблон
    template.render(context)
    # Задаем имя для файла при его генерации (название компании_дата_имя файла.формат)
    template.save(company + '_' + str(datetime.datetime.now().date()) + '_car.docx')


#Возврат словаря информации для заполнения шаблона
def Get_car_context(company, car, car_model, engine_volume, gearbox, price):
   return{
    'car_dealer': company,
    'car': car,
    'car_model': car_model,
    'engine_volume': engine_volume,
    'gearbox': gearbox,
    'price': price
   }

Get_car_info('Авто мир', 'Nissan', 'GT-R 2013', 3.8, 'АКПП', 2800000)

#Задание 2. Создать csv файл с данными о машине.
#ДРУГОЙ ЛОВАРЬ
car_dict = Get_car_context('Авто мир', 'Nissan', 'GT-R 2013', 3.8, 'АКПП', 2800000)
fieldnames = ['car_dealer', 'car', 'car_model', 'engine_volume', 'gearbox', 'price']


#Создаем сsv файл
with open('cars.csv', 'w') as file:
    write_info = csv.DictWriter(file, delimiter='|', fieldnames=fieldnames)
    write_info.writeheader()
    for i in range(len(car_dict)):
        write_info.writerow(car_dict[i])

#Задание 3. Создать json файл с данными о машине.
with open('cars.json', 'w') as file:
    json.dump(car_dict, file)