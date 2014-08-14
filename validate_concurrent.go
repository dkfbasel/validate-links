package main

import (
	"fmt"
	"io"
	"os"

	"sync"

	"path/filepath"

	"archive/zip"
	"bytes"

	"regexp"

	"github.com/franela/goreq"
	"time"

	"github.com/skratchdot/open-golang/open"
	"html/template"
	"log"
)

func main() {

	// fmt.Println("get documents")

	// var documents []*Document = getAndCheckFilesInDirectory(".")

	// for _, value := range documents {
	// 	fmt.Println(value.Path + " - " + value.Type)
	// }

	// fmt.Println("create report")
	// fmt.Println("be done")

	// measure execution time
	start := time.Now()

	fmt.Println("Checking documents. Please wait ..")

	// initialize our regular expressions
	initializeMatchers()

	// we are only interested in the current directory
	directories := []string{"."}

	// get current date and time
	currentTime := time.Now().String()
	currentTime = currentTime[:19]

	// get a list of all files in the directories specified
	documents := getAndCheckFilesInDirectory(".")

	// initialize our report structure
	report := Report{
		Directories: make([]string{"."}),
		Documents:   &documents,
		Date:        currentTime,
	}

	// create an html report with our data
	report.create()

	// open the report
	report.open()

	// measure the time of computing
	elapsed := time.Since(start)

	// inform user that process is finished
	log.Println("Finished! (it took %s", elapsed)

}

// define a custom document structure
type Document struct {
	Path       string
	Type       string
	IsValid    bool
	Hyperlinks []Hyperlink
}

// define a custom hyperlink structure
type Hyperlink struct {
	Url       string
	IsWorking bool
}

// define a custom report structre
type Report struct {
	ResultOfValidation bool
	Directories        []string
	Documents          *[]Document
	InvalidHyperlinks  []Hyperlink
	Date               string
}

// we define the name of our report
var reportName string = "report"

// define some custom regular expressions
var matchers map[string]*regexp.Regexp

// initialize our regular expresssions
func initializeMatchers() {

	// initialize our map of matchers
	matchers = make(map[string]*regexp.Regexp)

	// add our matching expressions
	matchers[".docx"] = regexp.MustCompile(`word/_rels/document.xml.rels`)
	matchers[".pptx"] = regexp.MustCompile(`ppt/slides/_rels/.*.xml.rels`)
	matchers["hyperlink"] = regexp.MustCompile(`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="(?P<url>.+?)"`)
	matchers["microsoft"] = regexp.MustCompile(`http://office.microsoft.com`)
	matchers["mailto"] = regexp.MustCompile(`mailto:.*`)
}

func getAndCheckFilesInDirectory(rootDirectory string) []*Document {

	// initialize a new slice of documents
	documents := []*Document{}

	var fileChannel chan Document = make(chan Document)

	var doneChannel chan bool = make(chan bool)

	var wg sync.WaitGroup

	// we have to wait until all file walking is done

	// walk recursively through our directory in a separate thread
	go walkDirectory(rootDirectory, fileChannel, &wg)

	for file := range fileChannel {

		// add a reference to the file to our documents slice
		documents = append(documents, &file)

		// remember to wait until the document is fully checked
		wg.Add(1)

		// check hyperlinks of the document
		go extractAndCheckHyperlinks(&file, &wg)

	}

	// file walking is done
	<-doneChannel

	// document checking is done
	wg.Wait()

	fmt.Println("we are done with these files")

	return documents

}

func walkDirectory(directory string, fileChannel chan Document, wg *sync.WaitGroup) {

	// walk recursively through the directory
	filepath.Walk(directory, func(path string, fileInfo os.FileInfo, err error) error {

		var fileName string = fileInfo.Name()

		var extension string = filepath.Ext(fileName)

		if extension == ".docx" || extension == ".pptx" {

			// create a pointer to new document with the corresponding type and path
			file := Document{Path: path, Type: filepath.Ext(fileName)}

			// send the file to the channel
			fileChannel <- file

		}

		// we are not expecting any errors (or not handling them at least)
		return nil

	})

	// close our fileChannel (no longer needed)
	close(fileChannel)

	// we are done walking the filepath
	wg.Done()

}

func extractAndCheckHyperlinks(doc *Document, wg *sync.WaitGroup) {
	fmt.Println("we are changing the type")
	doc.Type = "this is a new type"
	wg.Done()
}

func extractHyperlinkFromDocument(document Document) []Hyperlink {

	// get file content
	content := getLinkFileContent(document)

	// find all hyperlinks in the document
	matches := extractHyperlinksFromContent(content)

	// store the hyperlinks in a the document reference
	return matches

}

func getLinkFileContent(document Document) string {

	// open the docx file with our zip module (as it is basically a container)
	documentContainer, err := zip.OpenReader(document.Path)
	if err != nil {
		log.Println("ERROR: could not open the file")
	}
	defer documentContainer.Close()

	// initialize a new buffer to read the file contents
	buffer := bytes.NewBuffer(nil)

	// go through all content files
	for _, file := range documentContainer.File {

		// links are stored in a special file (but without the name of the link)
		if matchers[document.Type].MatchString(file.Name) {

			// open the file for reading
			fileContentReader, err := file.Open()

			if err != nil {
				log.Println("ERROR: coult not read file content")
			}
			defer fileContentReader.Close()

			// copy content of the file to our content buffer
			io.Copy(buffer, fileContentReader)

			if err != nil {
				log.Println("could not write file contents to console")
			}

		}
	}

	// return content as string
	return string(buffer.Bytes())

}

func extractHyperlinksFromContent(fileContent string) []Hyperlink {

	// find all matching links (the url of the hyperlink is matched by a capture group)
	matches := matchers["hyperlink"].FindAllStringSubmatch(fileContent, -1)

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

	// initialize an empty slice of strings
	var filteredLinks = []Hyperlink{}

	// check all links
	for _, link := range hyperlinks {

		// exclude any microsoft links
		isMicrosoft := matchers["microsoft"].MatchString(link.Url)
		isEmpty := (link.Url == "")
		isMail := matchers["mailto"].MatchString(link.Url)

		// include only non microsoft links
		if isMicrosoft == false && isEmpty == false && isMail == false {
			filteredLinks = append(filteredLinks, link)
		}

	}

	return filteredLinks

}

func isHyperlinkWorking(link Hyperlink) bool {

	// issue a GET request to the specified url and wait for response
	// set a timeout of 10 seconds if there is no response
	_, err := goreq.Request{
		Uri:     link.Url,
		Timeout: 10000 * time.Millisecond,
	}.Do()

	if err != nil {
		// link was not found
		return false
	} else {
		// link was found
		return true
	}

}

func getAbsoluteFilePath(path string) string {

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

	functionMap := template.FuncMap{
		"absolutePath": getAbsoluteFilePath,
	}

	// load our template from the templat file
	tmpl, err := template.New("report").Funcs(functionMap).Parse(reportTemplate)

	if err != nil {
		log.Println("Could not load template")
		return false

	} else {

		// fill our template with content and write it to the file
		err = tmpl.ExecuteTemplate(file, "report", report)

		if err != nil {
			log.Println("Could not fill the template with report data")
			return false
		}
	}

	return true

}

// open the report in the standard browser
func (report *Report) open() {

	// open the report in the default browser
	err := open.Start(reportName + ".html")

	if err != nil {
		log.Println("Could not open report")
	}

}

// we define the name of our report
var reportName string = "report"

