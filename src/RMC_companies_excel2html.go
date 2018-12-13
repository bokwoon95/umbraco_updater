package main

import (
	"errors"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

/*	main() takes in up to 2 optional command line arguments:
 *	os.Args[0]: name of program (always present)
 *	os.Args[1]: input filename (must be .xlsx)
 *	os.Args[2]: output filename (must be .html)
 */
func main() {

	argslist := os.Args
	inputfile, err := searchForInputFile(argslist)
	if err != nil {
		fmt.Println("Unable to determine inputfile")
		return
	}
	outputfile, err := searchForOutputFile(argslist, inputfile)
	if err != nil {
		fmt.Println("Unable to determine outputfile")
		return
	}

	headertemplate := "<p><em>“List updated on %s”</em></p>\n" +
		"<div id=\"templatemo_content_wrapper\">\n" +
		"<div id=\"templatemo_content\">\n" +
		"<div class=\"col_w265 float_l\">\n" +
		"<table border=\"1\" cellspacing=\"0\" cellpadding=\"0\" width=\"100%%\" bordercolor=\"#666\">\n" +
		"<tbody>\n"
	currentDate := time.Now().Format("2 January 2006")
	header := fmt.Sprintf(headertemplate, currentDate)

	footertemplate := `</tbody>
</table>
</div>
</div>
</div>
<p></p>`

	bodytemplate := "<tr bordercolor=\"#666666\">\n" +
		"<td width=\"123\" align=\"middle\">\n" +
		"<div>%s</div>\n" +
		"</td>\n" +
		"<td width=\"226\" align=\"middle\">\n" +
		"<div><strong>%s</strong></div>\n" +
		"</td>\n" +
		"<td width=\"333\" align=\"middle\">\n" +
		"<div>%s</div>\n" +
		"</td>\n" +
		"</tr>\n"

	bigstring := header
	xlFile, err := xlsx.OpenFile(inputfile)
	if err != nil {
		fmt.Println("There was a problem opening the excel file")
	}

	// bigstring = headertemplate + multiple bodytemplates + footertemplate
	// Every excel row becomes a formatted bodytemplate string inside bigstring
	for i, sheet := range xlFile.Sheets {
		for _, row := range sheet.Rows {
			// Read only the first 3 columns of each row and store inside rowvalues
			// rowvalues[0] : Certification Number
			// rowvalues[1] : RMC Producer
			// rowvalues[2] : Plant Address/ Name
			var rowvalues [3]string
			for k, cell := range row.Cells {
				text := cell.String()
				// Read only the first 3 columns and ignore the rest
				if k >= 3 {
					break
				}
				rowvalues[k] = strings.TrimSpace(text)
			}
			// If the first 3 columns are non-empty, insert them into the format specifiers
			// in the `bodytemplate` string
			if rowvalues[0] != "" && rowvalues[1] != "" && rowvalues[2] != "" {
				tempstring := fmt.Sprintf(bodytemplate, rowvalues[0], rowvalues[1], rowvalues[2])
				bigstring = bigstring + tempstring
			}
		}
		// Read only the first sheet and ignore the rest
		if i >= 1 {
			break
		}
	}
	bigstring = bigstring + footertemplate
	err = ioutil.WriteFile(outputfile, []byte(bigstring), 0644)
	if err != nil {
		fmt.Println("There was a problem writing to the html file")
	}
}

func searchForInputFile(argslist []string) (string, error) {

	// If an input file was provided as an argument,
	// check for existence; if input file doesn't exist, throw an error
	if len(argslist) >= 2 {
		inputfile := argslist[1]
		if _, err := os.Stat(inputfile); os.IsNotExist(err) {
			fmt.Printf("The file %s does not exist\n", inputfile)
			return "", errors.New("file does not exist")
		}
		return inputfile, nil
	}

	// If an input file was not provided, we must search for one
	file, err := os.Open(".")
	if err != nil {
		log.Fatalf("failed opening directory: %s", err)
	}
	defer file.Close()

	inputregex := regexp.MustCompile(`^Company holding our certificates rev(?P<revision>\d*)\.(?P<extension>.{1,})`)
	inputfile := ""
	revisionMax := -1
	fileList, _ := file.Readdir(0)

	// Loop through each element in current directory
	// if it matches `inputregex`, tentatively store it as the inputfile
	// There may be multiple files that match the regex, pick the one with the
	// highest revision number (the digits at the end of the filename)
	for _, file := range fileList {
		regexArray := inputregex.FindStringSubmatch(file.Name())
		// regexArray[0]: entire filename
		// regexArray[1]: revision number
		// regexArray[2]: file extension
		if regexArray != nil {
			if regexp.MustCompile(`^xls.?`).FindStringSubmatch(regexArray[2]) != nil {
				revision, err := strconv.Atoi(regexArray[1])
				if err != nil {
					log.Fatalf("unable to convert string to integer")
				}
				if revision > revisionMax {
					revisionMax = revision
					inputfile = regexArray[0]
				}
			}
		}
	}

	if inputfile == "" {
		return "", errors.New("no inputfile matching inputregex")
	}
	return inputfile, nil
}

func searchForOutputFile(argslist []string, inputfile string) (string, error) {

	// If an output file was provided as an argument,
	// check for existence; if output file already exists, throw an error
	if len(argslist) >= 3 {
		outputfile := argslist[2]
		if _, err := os.Stat(outputfile); !os.IsNotExist(err) {
			fmt.Printf("The file %s already exists\n", inputfile)
			return "", errors.New("file already exists")
		}
		return outputfile, nil
	}

	// If an output file was not provided, we must make our own
	// Simply replace the .xlsx extension with .html
	outputfile := ""
	outputregex := regexp.MustCompile(`^(.+)\.[^\.]+$`)
	outputregexArray := outputregex.FindStringSubmatch(inputfile)
	if outputregexArray == nil {
		log.Fatalf("outputregex matching failed on inputfile string")
	}
	outputfileName := outputregexArray[1]
	outputfile = outputfileName + ".html"
	return outputfile, nil
}
