import json

with open('json.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Filtrar los que existen
existentes = [item for item in data if item.get('exists')]

# Mostrar resultado
for contacto in existentes:
    print(contacto.get('number'))