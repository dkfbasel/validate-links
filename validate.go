package main

import (
	"fmt"

	"path/filepath"

	"archive/zip"
	"bytes"
	"io"
	"os"

	"regexp"
	"strings"

	"github.com/franela/goreq"
	"time"
)

// define some custom regular expressions to find hyperlinks in our word documents
var (
	hyperlinkExpression = regexp.MustCompile(`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="(?P<url>.+?)"`)
	//hyperlinkExpression = regexp.MustCompile(`Target="(?P<url>.+?)"`)
)

// define a custom document structure
type Document struct {
	path         string
	hyperlinks   []string
}

// our main function
func main() {

	// get a list of all files in the directories specified
	wordDocuments := getAllWordDocumentPathsInDirectory(".")

	invalidLinks := []string{}

	// go through all word documents and extract their hyperlinks
	for _, document := range wordDocuments {

		fmt.Println(document.path)

		document.findHyperlinks()

		for _, link := range document.hyperlinks {

			if isHyperlinkWorking(link) == false {

				invalidLinks = append(invalidLinks, link)
				fmt.Println("- " + link)
			}

		}
	}

}

// get paths for all word documents in the specified directories
func getAllWordDocumentPathsInDirectory(rootDirectory string) []Document {

	// initialize a new slice of documents
	wordFiles := []Document{}

	// walk recursively through our directory
	filepath.Walk(rootDirectory, func(path string, fileInfo os.FileInfo, err error) error {

		// fmt.Println(strings.HasSuffix(f.Name, ".docx")

		if strings.HasSuffix(fileInfo.Name(), ".docx") {

			// create a new document
			file := Document{path: path}

			// append the document to the list of existing documents
			wordFiles = append(wordFiles, file)
		}

		return nil

	})

	// return all documents found
	return wordFiles

}

// add a method to our document to extract all hyperlinks
func (doc *Document) findHyperlinks() bool {

	// get file content
	content := getFileContentAsString(doc.path)

	// find all hyperlinks in the document (this will unfortunately also return some microsoft schema links)
	matches := extractHyperlinks(content)

	// store the hyperlinks in a the document reference
	doc.hyperlinks = matches

	// extraction of hyperlinks successfully completed
	return true

}

func getFileContentAsString(filePath string) string {

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

func extractHyperlinks(fileContent string) []string {

	// find all matching links (the url of the hyperlink is matched by a capture group)
	matches := hyperlinkExpression.FindAllStringSubmatch(fileContent, -1)

	// initialize a new slice of strings of the same length as our matches
	var links []string = make([]string, len(matches))

	// we are only interested in the second element in our list as
	// the Submatch function returns the full match as first element and our capture group as second
	for index, match := range matches {
		links[index] = match[1]
	}

	// no filter out all microsoft links
	links = filterHyperlinks(links)

	return links

}

func filterHyperlinks(hyperlinks []string) []string {

	// regular expression to find microsoft links
	var microsoft string = `http://office.microsoft.com`

	// initialize an empty slice of strings
	var filteredLinks = []string{}

	// check all links
	for _, link := range hyperlinks {

		// exclude any microsoft links
		isMicrosoft := strings.Contains(link, microsoft)
		isEmpty := (link == "")

		// include only non microsoft links
		if isMicrosoft == false && isEmpty == false {
			filteredLinks = append(filteredLinks, link)
		}

	}

	return filteredLinks

}

func isHyperlinkWorking(link string) bool {

	// issue a GET request to the specified url and wait for response
	// set a timeout if there is no response
	_, err := goreq.Request{
		Uri:     link,
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
