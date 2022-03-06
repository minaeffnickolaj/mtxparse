package main

/*
	XSLX to MXT file parser
	This utilite uses xuri/excelize
	GNU GPL v2.0, 2022
	N.S. Minaev
*/

import (
	"bufio"
	"bytes"
	"io/ioutil"
	"os"
	"text/template"

	"github.com/xuri/excelize/v2"
	"golang.org/x/text/encoding/charmap"
)

// MobaXTerm connection struct
type MobaXTerm struct {
	APCode        string
	RKName        string // Региональная компания
	AptName       string // Имя аптеки в нотации типа 'мскАпт1001'
	ServerAddress string // Непостоянная часть адреса сервера
	Username      string // пользователь системы, efarma по умолчанию
}

func main() {
	print("Let's go parse files...")
	//Opening XLSX file
	XLSXFile, err := excelize.OpenFile("apt.xlsx")
	if err != nil {
		println("Unable to load XLSX file:", err)
	}
	defer func() { // correctly close XLSX
		if err := XLSXFile.Close(); err != nil {
			println(err)
			println("Parsing done!")
		}
	}()
	// struct slice of MobaXTerm connections
	MXTConnectionsSlice := []MobaXTerm{}
	//read XLSX
	Rows, err := XLSXFile.GetRows("Аптеки")
	if err != nil {
		println(err)
		return
	}
	for _, row := range Rows {
		/*
			Excel row in Excelize realized like [][]string, when each
			cell is [i][y] - i is Row, y is Cell
		*/
		MXTConnection := new(MobaXTerm) //create MXT
		MXTConnection.APCode = row[1]
		MXTConnection.RKName = row[2]
		MXTConnection.AptName = row[0]
		MXTConnection.ServerAddress = row[3]
		MXTConnection.Username = "efarma"
		MXTConnectionsSlice = append(MXTConnectionsSlice, *MXTConnection) // pointer to MXTConnection
	}
	//Begin generate MXT file
	//Creating template of MobaXTerm connection
	MXTTemplate := template.New("MobaXTermTemplate")
	MXTTemplateText := "\n\n[Bookmarks_2]\nSubRep={{.RKName}}\\{{.AptName}}\nImgNum=41\n{{.AptName}}({{.APCode}})=#91#4%{{.ServerAddress}}.apt.rigla.ru%10433%[{{.Username}}]%0%-1%-1%-1%-1%0%0%-1%%%%%0%0%%-1%%-1%-1%0%-1%0%-1#MobaFont%10%0%0%-1%15%236,236,236%30,30,30%180,180,192%0%-1%0%%xterm%-1%-1%_Std_Colors_0_%80%24%0%1%-1%<none>%%0%0%-1#0# #-1"
	MXTTemplate.Parse(MXTTemplateText)
	//create file to save output
	ParsedFile, err := os.Create("unprepared.mxtconnections")
	if err != nil {
		println(err)
		println("File not created!")
		return
	}
	for _, MobaXTerm := range MXTConnectionsSlice {
		MXTTemplate.Execute(ParsedFile, MobaXTerm)
	}
	// finally
	ConvertFileToOem855(FileToString())
	println("Job done!")
}

func FileToString() string {
	File, err := os.Open("unprepared.mxtconnections") //open file
	if err != nil {
		println("Error in FileToString func!", err)
	}
	ReturnedString := bufio.NewScanner(File) //scan file
	var BufferString bytes.Buffer
	StringWriter := bufio.NewWriter(&BufferString) // write to []byte
	for ReturnedString.Scan() {
		StringWriter.Write(ReturnedString.Bytes())
		StringWriter.WriteByte('\n')
	}
	if !ReturnedString.Scan() {
		println("Failed to read: %v", ReturnedString.Err())
	}
	Result := string(BufferString.Bytes())
	return Result
}

func ConvertFileToOem855(InputString string) {
	Encoder := charmap.CodePage855.NewEncoder()
	s, err := Encoder.String(InputString)
	if err != nil {
		panic(err)
	}
	ioutil.WriteFile("output.mxtsessions", []byte(s), os.ModePerm)
}
