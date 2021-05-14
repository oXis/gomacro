// +build windows

package main

import (
	"fmt"
	"os"
	"path"

	"github.com/oxis/gomacro/pkg/gomacro"
)

func main() {
	// Initialise the lib
	gomacro.Init()
	defer gomacro.Uninitialize()

	// Open Word and get a hendle to documents
	documents := gomacro.NewDocuments(true)
	defer documents.Close()

	fmt.Printf("Word version is %s\n", documents.Application.Version)
	currpath, err := os.Getwd()
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println(currpath)

	// Add a new document
	document := documents.AddDocument()

	obj := document.InlineShapes.AddOLEObject("", path.Join(currpath, "test.txt"), false, false)
	obj.SetWidth(1)
	obj.Setheight(1)

	document.SaveAs(path.Join(currpath, "Test.doc"))
	documents.Save()
}
