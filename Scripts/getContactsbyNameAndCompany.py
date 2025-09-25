import argparse
import requests
import arraysNameAndCargo
import time
import pandas as pd

def main(nameCampaign, date):
    urlbase = "https://api-public.salesql.com/v1/persons/enrich?full_name={}&organization_name={}"
    results = []
    for contact in arraysNameAndCargo.ARRAY:
        url = urlbase.format(contact["Nombre"].replace(" ", "%20"), contact["Company"].replace(" ", "%20"))
        print(url)
        headers = {
            "accept": "application/json",
            "Authorization": "Bearer wmMo7vj7Y3xm9qSd8v8WzexW3vnJf7lP"
        }
        response = requests.get(url, headers=headers)
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
            time.sleep(5)

    # Convertir resultados a un DataFrame de pandas
    df = pd.DataFrame(results)

    # Guardar los datos en un archivo Excel
    df.to_excel(f"{nameCampaign}{date}.xlsx", index=False)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process campaign data.')
    parser.add_argument('nameCampaign', type=str, help='The name of the campaign to be used in the filenames.')
    parser.add_argument('date', type=str, help='The date to be used in the filenames.')
    args = parser.parse_args()
    main(args.nameCampaign, args.date)
    print(f"File saved as {args.nameCampaign}{args.date}.xlsx")
