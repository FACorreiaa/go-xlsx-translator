# Excel Translator

The Excel Translator is a Python program that translates the values in the second column of an Excel file (XLSX). It utilizes the Azure Translation API to perform the translations. The translated values are then updated in the Excel file.

## Prerequisites

Before running the program, ensure that you have the following:

- Python programming language installed
- Azure Translation API subscription key
- Input Excel file (XLSX) with data in the first sheet

## Installation

1. Clone the repository:

```shell
git clone https://github.com/FACorreiaa/go-xlsx-translator
pip3 install openpyxl
python3 translator.py
```

## Usage

- Set your Azure Translation API subscription key in an environment variable named AZURE_TRANSLATION_KEY.

- Place the input Excel file (input.xlsx) with the values to be translated in the project directory.

- Run the program:

```shell
python3 translator.py
```

- The translated values will be updated in the second column of the input Excel file.

- The modified Excel file (output.xlsx) will be saved in the project directory.

# Configuration

## You can modify the following settings in the main.go file:

- endpoint: The Azure Translation API endpoint.
- targetLanguage: The target language for translation.

# Contributing

Contributions are welcome! If you find any issues or want to enhance the program, feel free to open a pull request.

# License

This project is licensed under the MIT License.

Feel free to modify the content as per your project's specific details and requirements.
