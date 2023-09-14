
import pandas as pd
import random
from datetime import datetime, timedelta

# Создаем случайные данные для примера
data = {
    'Date': [datetime(2020, 1, 1) + timedelta(days=i) for i in range(365)],
    'Quantity Sold': [random.randint(10, 100) for _ in range(365)],
    'Price': [random.uniform(10, 100) for _ in range(365)],
    'Revenue': [],
    'Product Category': [random.choice(['Category A', 'Category B', 'Category C']) for _ in range(365)],
    'Promotions': [random.choice(['Yes', 'No']) for _ in range(365)]
}

# Вычисляем выручку
for index, row in enumerate(data['Quantity Sold']):
    data['Revenue'].append(row * data['Price'][index])

# Создаем DataFrame из данных
df = pd.DataFrame(data)

# Создаем файл Excel
df.to_excel('sales_data.xlsx', index=False)
