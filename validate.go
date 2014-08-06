package main

import (
	"fmt"
	"io"
	"os"

	"path/filepath"

	"archive/zip"
	"bytes"

	"regexp"
	"strings"

	"github.com/franela/goreq"
	"time"

	"html/template"
	"github.com/skratchdot/open-golang/open"
	"log"
)

// define some custom regular expressions to find hyperlinks in our word documents
// we define this globally to avoid recompilation
var hyperlinkExpression = regexp.MustCompile(`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="(?P<url>.+?)"`)

// we define the name of our report
var reportName string = "report"

// define a custom document structure
type Document struct {
	Path       string
	IsValid    bool
	Hyperlinks []Hyperlink
}

// define a custom hyperlink structure
type Hyperlink struct {
	Url     string
	IsWorking bool
}

// define a custom report structre
type Report struct {
	ResultOfValidation bool
	Directories        []string
	Documents          *[]Document
	InvalidHyperlinks  []Hyperlink
	Date string
}

// our main function
func main() {

	// we are only interested in the current directory
	directories := []string{"."}

	// get current date and time
	currentTime := time.Now().String()
	currentTime = currentTime[:19]

	// get a list of all files in the directories specified
	wordDocuments := getDocxFilesInDirectory(".")

	// initialize our report structure
	report := Report{
		Directories: directories,
		Documents: &wordDocuments,
		Date: currentTime,
	}

	// go through all word documents and extract their hyperlinks
	for index := range wordDocuments {

		// set word document validity initially to valid
		wordDocuments[index].IsValid = true;

		// get hyperlinks from the word file
		hyperlinks := extractHyperlinksFromDocxFile(wordDocuments[index].Path)

		// iterate through all links
		for _, link := range hyperlinks {

			link.IsWorking = isHyperlinkWorking(link)

			wordDocuments[index].Hyperlinks = append(wordDocuments[index].Hyperlinks, link)

			if link.IsWorking == false {

				// word document is not valid
				wordDocuments[index].IsValid = false

				// remember our invalid hyperlinks
				report.InvalidHyperlinks = append(report.InvalidHyperlinks, link)

			}

		}
	}

	// create an html report with our data
	report.create()

	// open the report
	report.open()

}

// get paths for all word documents in the specified directories
func getDocxFilesInDirectory(rootDirectory string) []Document {

	// initialize a new slice of documents
	wordFiles := []Document{}

	// walk recursively through our directory
	filepath.Walk(rootDirectory, func(path string, fileInfo os.FileInfo, err error) error {

		// fmt.Println(strings.HasSuffix(f.Name, ".docx")

		if strings.HasSuffix(fileInfo.Name(), ".docx") {

			// create a new document
			file := Document{Path: path}

			// append the document to the list of existing documents
			wordFiles = append(wordFiles, file)
		}

		return nil

	})

	// return all documents found
	return wordFiles

}

// add a method to our document to extract all hyperlinks
func extractHyperlinksFromDocxFile(filePath string) []Hyperlink {

	// get file content
	content := getLinkFileContent(filePath)

	// find all hyperlinks in the document
	matches := extractHyperlinksFromContent(content)

	// store the hyperlinks in a the document reference
	return matches

}

func getLinkFileContent(filePath string) string {

	// open the docx file with our zip module (as it is basically a container)
	docxContainer, err := zip.OpenReader(filePath)
	if err != nil {
		fmt.Println("ERROR: could not open the file")
	}
	defer docxContainer.Close()

	// initialize a new buffer to read the file contents
	buffer := bytes.NewBuffer(nil)

	// go through all content files
	for _, file := range docxContainer.File {

		// links are stored in a special file (but without the name of the link)
		if file.Name == "word/_rels/document.xml.rels" {

			// open the file for reading
			fileContentReader, err := file.Open()

			if err != nil {
				fmt.Println("ERROR: coult not read file content")
			}
			defer fileContentReader.Close()

			// copy content of the file to our content buffer
			io.Copy(buffer, fileContentReader)

			if err != nil {
				fmt.Println("could not write file contents to console")
			}

		}
	}

	// return content as string
	return string(buffer.Bytes())

}

func extractHyperlinksFromContent(fileContent string) []Hyperlink {

	// find all matching links (the url of the hyperlink is matched by a capture group)
	matches := hyperlinkExpression.FindAllStringSubmatch(fileContent, -1)

	// initialize a new slice of strings of the same length as our matches
	var links []Hyperlink = make([]Hyperlink, len(matches))

	// we are only interested in the second element in our list as
	// the Submatch function returns the full match as first element and our capture group as second
	for index, match := range matches {
		links[index] = Hyperlink{Url: match[1], IsWorking: false}
	}

	// no filter out all microsoft links
	links = filterHyperlinks(links)

	return links

}

func filterHyperlinks(hyperlinks []Hyperlink) []Hyperlink {

	// regular expression to find microsoft links
	var microsoft string = `http://office.microsoft.com`

	// initialize an empty slice of strings
	var filteredLinks = []Hyperlink{}

	// check all links
	for _, link := range hyperlinks {

		// exclude any microsoft links
		isMicrosoft := strings.Contains(link.Url, microsoft)
		isEmpty := (link.Url == "")

		// include only non microsoft links
		if isMicrosoft == false && isEmpty == false {
			filteredLinks = append(filteredLinks, link)
		}

	}

	return filteredLinks

}

func isHyperlinkWorking(link Hyperlink) bool {

	// issue a GET request to the specified url and wait for response
	// set a timeout if there is no response
	_, err := goreq.Request{
		Uri:     link.Url,
		Timeout: 1500 * time.Millisecond,
	}.Do()

	if err != nil {
		// link was not found
		return false
	} else {
		// link was found
		return true
	}

}

func getAbsoluteFilePath (path string) string {

	// check if the path is already absolute
	if filepath.IsAbs(path) {
		// if so, return the path
		return path
	}

	// try to find the absolute path
	absolutePath, err := filepath.Abs(path)

	if err != nil {
		// we return the original path, if there was an error
		return path
	} else {
		return absolutePath
	}

}

// create a custom report
func (report *Report) create() bool {

	// open a new file to write our report to
	file, _ := os.Create(reportName + ".html")
	defer file.Close()

	functionMap := template.FuncMap {
		"absolutePath": getAbsoluteFilePath,
	}

	// load our template from the templat file
	tmpl, err := template.New("report").Funcs(functionMap).ParseFiles(reportName + ".tmpl")

	if err != nil {
		panic(err)
		return false

	} else {

		// fill our template with content and write it to the file
		err = tmpl.ExecuteTemplate(file, reportName+".tmpl", report)

		if err != nil {
			panic(err)
			fmt.Println("Could not fill the template with report data")
			return false
		}
	}

	return true

}

func (report *Report) open() {

	// open the report in the default browser
	err := open.Start(reportName + ".html")

	if err != nil {
		log.Fatal(err)
	}


}
