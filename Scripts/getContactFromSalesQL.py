import requests
import pandas as pd
import time


# Lista de URLs de perfiles de LinkedIn
string = """https://www.linkedin.com/in/soniareyescastro/
https://www.linkedin.com/in/marcos-penna-885a5619/
https://www.linkedin.com/in/wilson-david-tovar-monje-b8939098/
https://www.linkedin.com/in/rodolfo-r-2a969b62/
https://www.linkedin.com/in/carlos-eslava-b3b35466/
https://www.linkedin.com/in/diego-giraldo-/
https://www.linkedin.com/in/german-cardona-29a8522b/
https://www.linkedin.com/in/edisson-rocha-castro-456b8285/
https://www.linkedin.com/in/equinterov/
https://www.linkedin.com/in/petrobras-anderson-belém-ferreira-bertolossi-23348a92/
https://www.linkedin.com/in/aldo-ivan-lopez-villegas/
https://www.linkedin.com/in/erick-j-gonzalez-ventura-3a78212b/
https://www.linkedin.com/in/graciela-mirelle-mañón-arriaga-2747b264/
https://www.linkedin.com/in/mx-mario-alberto-fuentes-lópez-a0691b16b/
https://www.linkedin.com/in/césar-alberto-cruz-65a24a56/
https://www.linkedin.com/in/magali-barrera-colin-2b897316/
https://www.linkedin.com/in/joseluis-velazquez-ab313082/
https://www.linkedin.com/in/juanleandrolemus/
https://www.linkedin.com/in/miguel-toledo-1508ba13/
"""
URLs = string.split("\n")

# Base URL de la API
urlbase = "https://api-public.salesql.com/v1/persons/enrich/?linkedin_url="

# Encabezados para la API
headers = {
    "accept": "application/json",
    "Authorization": "Bearer wmMo7vj7Y3xm9qSd8v8WzexW3vnJf7lP"
}

# Lista para almacenar los resultados
results = []

# Procesar cada URL
for url in URLs:
    full_url = urlbase + url
    response = requests.get(full_url, headers=headers)
    print(response.json())
    if response.status_code == 200:  # Si la respuesta es exitosa
        data = response.json()
        
        # Extraer información organizada
        results.append({
            "UUID": data.get("uuid"),
            "First Name": data.get("first_name"),
            "Last Name": data.get("last_name"),
            "Full Name": data.get("full_name"),
            "LinkedIn URL": data.get("linkedin_url"),
            "Title": data.get("title"),
            "Headline": data.get("headline"),
            "Industry": data.get("industry"),
            "Emails": ", ".join([email["email"] for email in data.get("emails", [])]),
            "Phones": ", ".join([phone["phone"] for phone in data.get("phones", [])]),
            "Location": data.get("location"),
            "Organization Name": data["organization"].get("name") if data.get("organization") else None,
            "Organization Website": data["organization"].get("website") if data.get("organization") else None,
            "Organization LinkedIn URL": data["organization"].get("linkedin_url") if data.get("organization") else None,
            "Organization Industry": data["organization"].get("industry") if data.get("organization") else None,
            "Organization Location": data["organization"].get("location") if data.get("organization") else None,
        })
    else:
        # En caso de error, registrar un mensaje vacío
        results.append({
            "UUID": None,
            "First Name": None,
            "Last Name": None,
            "Full Name": None,
            "LinkedIn URL": url,
            "Title": None,
            "Headline": None,
            "Industry": None,
            "Emails": None,
            "Phones": None,
            "Location": None,
            "Organization Name": None,
            "Organization Website": None,
            "Organization LinkedIn URL": None,
            "Organization Industry": None,
            "Organization Location": None,
        })
        time.sleep(3)


# Convertir resultados a un DataFrame de pandas
df = pd.DataFrame(results)

# Guardar los datos en un archivo Excel
df.to_excel("TestDataFromSalesQLToNebula.xlsx", index=False)

print("Los datos han sido guardados en 'TestDataFromSalesQLToNebula.xlsx'")
