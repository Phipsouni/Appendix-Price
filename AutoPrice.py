import pandas as pd
import re

# Чтение исходных данных из txt-файла
with open('Appendix.txt', 'r') as file:
    data = file.readlines()

# Функция для извлечения информации из каждой строки
def process_data(entry):
    match = re.match(r'([^,]+),\s*([\d/]+)\*?([\d/]*),\s*([\d.,-]+),\s*KD\s*(\d+)\s*(\d+)', entry)
    if match:
        group = match.group(1).strip()
        # Замена "VI" на "3" и "VII" на "4"
        group = group.replace("VII", "4").replace("VI", "3")
        starts = match.group(2).split('/')
        ends = match.group(3).split('/')
        price = match.group(6)
        combinations = [(start, end, group, price) for start in starts for end in ends]
        return combinations
    else:
        return None

# Обработка данных
formatted_data = [process_data(entry) for entry in data]
formatted_data = [item for sublist in formatted_data if sublist is not None for item in sublist]

# Создаем DataFrame
df = pd.DataFrame(formatted_data, columns=["Start", "End", "Group", "Price"])

# Имя для Excel файла
xlsx_file = "Price.xlsx"

# Записываем данные в Excel
df.to_excel(xlsx_file, index=False)

print("Данные успешно записаны в файл", xlsx_file)