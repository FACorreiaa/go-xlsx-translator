import os
import time
import requests
import html
import openpyxl

start_time = time.time()

subscription_key = os.environ.get('AZURE_TRANSLATION_KEY')
endpoint = "https://api.cognitive.microsofttranslator.com/translate"

source_language = 'de'
target_language = 'en-en'

workbook = openpyxl.load_workbook('data.xlsx')

sheet = workbook.active

for row in sheet.iter_rows():
    for cell in row:
        source_text = cell.value

        if source_text is None:
            continue

        params = {
            "api-version": "3.0",
            "from": source_language,
            "to": target_language
        }

        headers = {
            "Ocp-Apim-Subscription-Key": subscription_key,
            "Content-Type": "application/json",
            "Ocp-Apim-Subscription-Region": "westeurope"
        }
        body = [{'text': source_text}]
        print(body)
        response = requests.post(endpoint, headers=headers, params=params, json=body)
        response.raise_for_status()

        translation = html.unescape(response.json()[0]['translations'][0]['text'])
        cell.value = translation
        print(translation)

workbook.save('output_file.xlsx')

print("--- Translation took %s seconds ---" % (time.time() - start_time))
