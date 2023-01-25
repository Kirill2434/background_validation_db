import csv
file = r'C:\source_data_4\0202.csv'
with open(file, mode='w', encoding='utf-8', errors='ignore') as f:
    writer = csv.writer(f, delimiter=',')
    f.writerow = (['Состояние обработки документа', 'УН документа', 'Код НО',
                   'Регномер', 'Дата', 'Номер', 'УН объекта НП',
                   'ФИД НП', 'ИНН', 'КПП', 'Наименование', 'УН НП', 'МНК',
                   'Дата вызова', 'Приемные дни, часы',	'Дата явки', 'Адм.штраф', 'Явка', 'GUID документа',
                   'Дата отправки',	'Дата получения', 'ун способа получения',
                   'Наименование способа получения', 'ун документа - основания МНК',
                   'ФИО инспектора-исполнителя', 'Фамилия инспектора-исполнителя', 'УН заявки',
                   'УН состояния обработки', 'Область контроля'])
    f.close()