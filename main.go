package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"time"

	"github.com/xuri/excelize/v2"
)

type TranslationResponse struct {
	Rows []string
}

type TranslationRequest struct {
	Rows []string
}

func main() {
	start := time.Now()

	// Set your subscription key and endpoint.
	subscriptionKey := os.Getenv("AZURE_TRANSLATION_KEY")
	endpoint := "https://api.cognitive.microsofttranslator.com"

	// Set the target language.
	targetLanguage := "en"

	// Open the Excel file.
	xlFile, err := excelize.OpenFile("data.xlsx")
	if err != nil {
		fmt.Println("Error opening file:", err)
		return
	}

	// Get the name of the sheet and the desired column.
	sheetName := xlFile.GetSheetName(0)
	columnLetter := "food"

	rows, err := xlFile.GetRows(sheetName)
	if err != nil {
		fmt.Println("Error retrieving rows:", err)
		return
	}

	// Extract the values from the desired column.
	var columnValues []string
	for _, row := range rows {
		if len(row) > 1 {
			columnValues = append(columnValues, row[1]) // Assuming the second column (index 1) contains the values to translate
		}
	}

	// Translate each value in the column.
	for i, value := range columnValues {
		// Translate the value.
		translation, err := translateText(value, subscriptionKey, endpoint, targetLanguage)
		if err != nil {
			fmt.Println("Error translating rows:", err)
			return
		}
		// Update the translated value in the same cell.
		cell := fmt.Sprintf("%s%d", columnLetter, i+1)
		err = xlFile.SetCellValue(sheetName, cell, translation)
		if err != nil {
			log.Fatal("Error setting cell value:", err)
		}
	}

	// Save the modified Excel file.
	err = xlFile.SaveAs("output.xlsx")
	if err != nil {
		fmt.Println("Error saving file:", err)
		return
	}

	fmt.Println("Translation complete.")
	elapsed := time.Since(start)
	log.Printf("Translation took %s seconds", elapsed)
}

func translateText(text, subscriptionKey, endpoint, targetLanguage string) (string, error) {
	// Create the translation request body.
	//requestBody := []string

	// Encode the request body to JSON.
	requestBodyBytes := new(bytes.Buffer)
	json.NewEncoder(requestBodyBytes).Encode(text)
	// Create the HTTP request.
	request, err := http.NewRequest("POST", fmt.Sprintf("%s/translate?api-version=3.0&to=%s", endpoint, targetLanguage), requestBodyBytes)
	if err != nil {
		return "", err
	}
	request.Header.Set("Content-Type", "application/json")
	request.Header.Set("Ocp-Apim-Subscription-Key", subscriptionKey)
	request.Header.Set("Ocp-Apim-Subscription-Region", "westeurope")

	// Send the HTTP request.
	client := &http.Client{}
	response, err := client.Do(request)
	if err != nil {
		return "", err
	}
	defer response.Body.Close()

	// Read the response.
	responseBody, _ := io.ReadAll(response.Body)

	// Parse the translation response.
	var translationResponse []TranslationResponse
	err = json.Unmarshal(responseBody, &translationResponse)
	if err != nil {
		return "", err
	}

	// Extract the translated text.
	if len(translationResponse) > 0 && len(translationResponse[0].Rows) > 0 {
		return translationResponse[0].Rows[0], nil
	}

	return "", fmt.Errorf("no translations found")
}
