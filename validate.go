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

	"time"

	"github.com/franela/goreq"

	"html/template"
	"log"

	"github.com/skratchdot/open-golang/open"
)

func main() {

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

	var resultOfValidation bool = true

	for index, _ := range documents {

		// initialize document validity with true
		documents[index].IsValid = true

		for _, link := range documents[index].Hyperlinks {

			if link.IsWorking == false {
				documents[index].IsValid = false
				resultOfValidation = false
			}
		}
	}

	if resultOfValidation != false {
		resultOfValidation = true
	}

	// initialize our report structure
	report := Report{
		ResultOfValidation: resultOfValidation,
		Directories:        directories,
		Documents:          documents,
		Date:               currentTime,
	}

	// create an html report with our data
	report.create()

	// open the report
	report.open()

	// measure the time of computing
	elapsed := time.Since(start)

	// inform user that process is finished
	log.Printf("Finished! (it took %s\n", elapsed)

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

func (link *Hyperlink) validate(wg *sync.WaitGroup) {

	// issue a GET request to the specified url and wait for response
	// set a timeout of 10 seconds if there is no response
	_, err := goreq.Request{
		Uri:     link.Url,
		Timeout: 15000 * time.Millisecond,
	}.Do()

	if err != nil {
		// link was not found
		link.IsWorking = false
	} else {
		// link was found
		link.IsWorking = true
	}

	wg.Done()

}

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

func getAndCheckFilesInDirectory(rootDirectory string) []Document {

	// initialize a new slice of documents
	documents := []Document{}

	var fileChannel chan Document = make(chan Document)

	var wg sync.WaitGroup

	// we have to wait until all file walking is done
	wg.Add(1)

	// walk recursively through our directory in a separate thread
	// send each matching file into our file channel
	go walkDirectory(rootDirectory, fileChannel, &wg)

	// for each file we find, we get all links and validate them
	for file := range fileChannel {

		// remember to wait until the document is fully checked
		var checkingHyperlinks sync.WaitGroup

		checkingHyperlinks.Add(1)

		// check hyperlinks of the document
		go extractAndCheckHyperlinks(&file, &checkingHyperlinks)

		// wait until all hyperlinks are checked
		checkingHyperlinks.Wait()

		documents = append(documents, file)

	}

	// we are finished with finding and checking all elements
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

func extractAndCheckHyperlinks(file *Document, wg *sync.WaitGroup) {

	// get all hyperlinks from the document
	file.Hyperlinks = extractHyperlinksFromDocument(*file)

	// check all hyperlinks in a separate routine
	for index, _ := range file.Hyperlinks {

		wg.Add(1)

		fmt.Println("-- checking link: " + file.Hyperlinks[index].Url)

		go file.Hyperlinks[index].validate(wg)

	}

	wg.Done()

}

func extractHyperlinksFromDocument(document Document) []Hyperlink {

	// get the content of the file containing the links
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

func getAbsoluteFilePath(path string) string {

	// check if the path is already absolute
	if filepath.IsAbs(path) {
		// if so, return the path
		return path
	}

	// try to find the absolute path
	absolutePath, err := filepath.Abs(path)

	// we return the original path, if there was an error
	// otherwise we return the absolute path
	if err != nil {
		return path
	} else {
		return absolutePath
	}

}

// define a custom report structre
type Report struct {
	ResultOfValidation bool
	Directories        []string
	Documents          []Document
	InvalidHyperlinks  []Hyperlink
	Date               string
}

// create a custom html report
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

const reportTemplate = `<html>
<head>
<title>Check hyperlinks in docx and pptx files</title>
<meta charset="utf-8">
<meta name="author" content="Dr. med. Ramon Saccilotto, DKF, University Hospital Basel, Switzerland">

<style type="text/css">

* {
font-family: "Helvetica Neue", "Helvetica", "Calibri", "Arial", sans-serif;
font-weight: normal;
font-size: 14px;
}

a {
text-decoration: none;
color: inherit;
font-size: inherit;
}

body {
background-color: #eaeaea;
}

.container {
min-width: 600px;
margin: 40px 20px 20px 20px;
padding: 20px;
padding-bottom: 40px;
border: 1px solid #ccc;
background-color: #fff;
box-shadow: 0px 1px 1px rgba(74, 69, 69, 0.6), 0px -1px 1px rgba(50, 50, 50, 0.05);
}

.info {
margin: 20px;
}

.info p {
font-size: 12px;
opacity: 0.2;
}

.info:hover p{
opacity: 1;
transition: opacity 500ms;
}

h1 {
margin: 0px;
padding: 0px;
margin-bottom: 15px;
font-size: 20px;
border-bottom: 1px solid #ccc;
font-weight: bold;
padding-bottom: 10px;

}

ul {
margin: 0px;
padding: 0px;
list-style-type: none;
padding-left: 5px;
}

ul li {
position: relative;
padding-left: 28px;
}

ul.directories > li + li{
margin-top: 20px;
}

ul li:before {
content: "";
background-position: top left;
background-repeat: no-repeat;
display: block;
position: absolute;
left: 0px;
top: 0px;
width: 20px;
height: 20px;
}

ul.directories {
margin-bottom: 40px;
}

ul.directories li:before {
background-image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAj0lEQVQ4T2NkYGBgBGJ2IIaBv0DGbyQ+XiZI81Ig/gLE/6EqFYB0PRCfJMYQkAELgDgBSTEHkD2DCM0gvaewGUCEXriSdSADzgPxLSQvEGsASK8aiLgPxIrE6kJTdx9kwB0gViHTgDtUMeAo0HZrMl1wlNJYWDBqAAM4DNAzE7ERAtLLAyLQszOxBoDU/QQAylQgG9KLVSEAAAAASUVORK5CYII=);
}

ul.documents > li:before {
background-image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAABqElEQVQ4T52VPSiFURjH3VDyMRhEDK6Uj0wG+VgskhSDDJLMmCUD8p2S7RbJxKZMLGKgJBsGJKQsZCeUj9//dl699ziXk1O/3vs873P+5/8+57zvjaQkjhHCbCvnCs9JrrpuRKzkNXGfh+AKNUswZ9faghcUVHoInlHzAhswG64PBNNM8pBrgyX4Sfxu5STYCmuwA1PBfQnWmZVukzirIp/rEGwk9wRbcAyDqgkEe/i9D5dwB1lGoJqrNkqLhoccqt+vkAMtMA4TgWA3QTnMQA08wxWUQa9DUK4zrEXWiUsDwS4C9enIONrl+gGjsGcE1dtmS0ThpKmNb2hYUI/bDidQDxKdNovokaNQ4RDcJqeN+yGo5sagFh6NwH1IUJvQ5hAcSuZQ1uVAR0cTN81ktUEOi6DEIXhgcgkOl42Qoz7ehkJogk5HwYDtMJOEdlNjEfqtSerPKeQZl7ameq7x7TBcELx6+SQXXHYdOR0rnYhfBdMpKPYU1AH/02EBFfqa+IwOH4eppmc+gg8+Dn2E7BpnD2+oGv6PGnPmIWp/YH3/AlxrvpEc+wLSwV8VusxZtAAAAABJRU5ErkJggg==);
top: -2px;
}

ul.documents > li + li {
margin-top: 25px;
}

div.result {
border-width: 1px;
border-style: solid;
border-radius: 2px;
padding: 10px;
margin-bottom: 25px;
}

div.result.valid {
border-color: #20d420;
background-color: #dcffe7;
}

div.result.invalid {
border-color: #db2d2d;
background-color: #fff5f5;
}

h2 {
font-weight: bold;
}

ul.links li{
font-size: 12px;
}

ul.links > li + li {
margin-top: 15px;
}

ul.links li.invalid:before {
background-image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAA/UlEQVQ4T2NkoBAwUqifAa8B/xkYjIAW/AcqOo/LIpwGADVLAjVdBxkAxFpAhc+xGYLPgHVADYFQTeuACoOJNgBopT9Q8QY0Df5AQzahG4LhAqBmLqCiy0CsBMTToBqygPQ9INYBaviObAg2A3qACoqhihqhdD2U7gFqKMVpANB2Q6DkKSBmwWHAH6C4GXKswF0A1AxiHwJiGyQbJkPZuUhiR4BsO6BiUOwg0gGQB/LnVGwhjUUsC2jAdLgBQM0iQM5NIBZCU7wCakk4mvg7IF8daMgbsBeABiQAqflE2g5TFg/UvAhmAD9QFGQIB5GG/ABZCNT8ibaZiRjXAABQjy8Rw0RFZAAAAABJRU5ErkJggg==);
top: -1px;
}

ul.links li.valid:before {
background-image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAA6UlEQVQ4T2NkoDJgpLJ5DKMGEhei/Ve1DRj/M/T///01sNDwwQdkXSSHIdSw/YyMDAL//zMcKNC56ki2gWiGffjL8NuhWOfWRbIMJMYwkMEYXoZo/J9foHMtEWZz/2UNB0ZG5vVQb2J1GUwthoETrmidZ2RkNPj3769joe6NA0ALEpgYGOaDNADDDK9hWF3Ye0VNn4WR9cJ/hv8P/v//P5GJkamfWMOwGggSnHhZM5+BiWkCzBvEuAynl2ESQK/PB3o9gRTDcLoQJNFxRomfg4NzAzBpFKAnDXzJn+SETSgvjRpIKIQIywMAWmd1FTm7YC8AAAAASUVORK5CYII=);
top: -3px;
}

.valid {
color: #20d420;
}

.invalid {
color: #db2d2d;
}



</style>
</head>
<body>
<div class="container">

<h1>Directory searched</h1>

<ul class="directories">
{{range .Directories}}
<li><a href="file:///{{absolutePath .}}">{{absolutePath .}}</a></li>
{{end}}
</ul>

<h1>Result of link validation</h1>


{{if .ResultOfValidation}}
<div class="result valid">
All files contain valid links
</div>
{{else}}
<div class="result invalid">
There are some files with invalid links
</div>
{{end}}

<ul class="documents">
{{range .Documents}}
<li class="result">
<h2 class="{{if .IsValid}}valid{{else}}invalid{{end}}"><a href="file:///{{absolutePath .Path}}">{{.Path}}</a></h2>

<ul class="links">
{{range .Hyperlinks}}
<li class="result {{if .IsWorking}}valid{{else}}invalid{{end}}"><a href="{{.Url}}">{{.Url}}</a></li>
{{end}}
</ul>
</li>
{{end}}
</ul>

</div>

<div class="info">
<p class="time">Link validation conducted on {{.Date}}</p>
<p class="author">&copy;&nbsp;2014, Department of Clinical Research, University Hospital Basel</p>
</div>
</body>
</html>
`
