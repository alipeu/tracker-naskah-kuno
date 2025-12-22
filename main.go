package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"os"
	"strings"
	"text/template"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

type Manuskrip struct {
	NoPanggil string
	Judul     string
	BIBID     string
	Bahasa    string
	Aksara    string
	Media     string
	Halaman   string
	Dimensi   string
	Status    string
	LinkOPAC  string
}

// Retrieve a token, saves the token, then returns the generated client.
func getClient(config *oauth2.Config) *http.Client {
	// The file token.json stores the user's access and refresh tokens, and is
	// created automatically when the authorization flow completes for the first
	// time.
	tokFile := "token.json"
	tok, err := tokenFromFile(tokFile)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokFile, tok)
	}
	return config.Client(context.Background(), tok)
}

// Request a token from the web, then returns the retrieved token.
func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser then type the "+
		"authorization code: \n%v\n", authURL)

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("Unable to read authorization code: %v", err)
	}

	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		log.Fatalf("Unable to retrieve token from web: %v", err)
	}
	return tok
}

// Retrieves a token from a local file.
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

// Saves a token to a file path.
func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}

func getSheetsData() (map[string]any, error) {
	ctx := context.Background()
	b, err := os.ReadFile("credentials.json")
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	// If modifying these scopes, delete your previously saved token.json.
	config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets.readonly")
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	client := getClient(config)

	srv, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	// Prints the names and majors of students in a sample spreadsheet:
	// https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
	spreadsheetId := "1EDLv6f8ehprYBno9sRBrXvDXk07N_06KNM2c_VmxoYM"
	readRange := "rekap!A2:L"
	resp, err := srv.Spreadsheets.Values.Get(spreadsheetId, readRange).Do()
	if err != nil {
		log.Fatalf("Unable to retrieve data from sheet: %v", err)
	}

	pageData := make(map[string]any)
	manuskripsCount := 0
	unggahCount := 0
	postProcessingCount := 0
	pemotretanCount := 0
	penelusuranCount := 0

	if len(resp.Values) == 0 {
		fmt.Println("No data found.")
	} else {
		manuskrips := []Manuskrip{}
		for _, row := range resp.Values {
			manuskrip := new(Manuskrip)

			noPanggil := row[0].(string)
			if (noPanggil == "-") || (noPanggil == "#REF!") {
				continue
			}
			judul := row[1].(string)
			bibID := row[2].(string)
			bahasa := row[3].(string)
			aksara := row[4].(string)
			media := row[5].(string)
			halaman := row[6].(string)
			dimensi := row[7].(string)
			status := row[8].(string)
			linkOPAC := row[10].(string)

			if strings.Contains(status, "unggah") {
				status = "unggah"
			}

			if (status == "unggah") && (linkOPAC == "-") {
				status = "post processing"
			}

			switch status {
			case "unggah":
				unggahCount += 1
			case "post processing":
				postProcessingCount += 1
			case "pemotretan":
				pemotretanCount += 1
			default:
				penelusuranCount += 1
			}

			manuskrip.NoPanggil = noPanggil
			manuskrip.Judul = judul
			manuskrip.BIBID = bibID
			manuskrip.Bahasa = bahasa
			manuskrip.Aksara = aksara
			manuskrip.Media = media
			manuskrip.Halaman = halaman
			manuskrip.Dimensi = dimensi
			manuskrip.Status = status
			manuskrip.LinkOPAC = linkOPAC

			manuskrips = append(manuskrips, *manuskrip)
			manuskripsCount += 1
		}

		pageData["Manuskrips"] = manuskrips
		pageData["TotalManuskrips"] = manuskripsCount
		pageData["TotalUnggah"] = unggahCount
		pageData["TotalPostProcessing"] = postProcessingCount
		pageData["TotalPemotretan"] = pemotretanCount
		pageData["TotalPenelusuran"] = penelusuranCount
	}
	return pageData, nil
}

func main() {
	homeHandler := func(w http.ResponseWriter, r *http.Request) {
		tmpl := template.Must(template.ParseFiles("index.html"))
		pageData, err := getSheetsData()
		if err != nil {
			log.Fatalf("Unable to retrieve data from sheet: %v", err)
		}
		tmpl.Execute(w, pageData)
	}

	http.HandleFunc("/", homeHandler)
	http.Handle("/static/", http.StripPrefix("/static/", http.FileServer(http.Dir("static"))))

	log.Println("Server started at http://localhost:8080")
	log.Fatal(http.ListenAndServe(":8080", nil))
}
