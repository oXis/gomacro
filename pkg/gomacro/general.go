package gomacro

import (
	"fmt"

	"golang.org/x/sys/windows/registry"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Documents represents Word documents
type Documents struct {
	_Documents  *ole.IDispatch
	Application *Application

	Document *Document
}

// Document represents a Word document
type Document struct {
	_Document   *ole.IDispatch
	Application *Application
	VBProject   *VBProject
}

//Application Holds the OLE app
type Application struct {
	_Application *ole.IDispatch
	Options      *Options
	Version      string
}

//Options Holds the OLE app
type Options struct {
	_Options *ole.IDispatch
}

// Init to init OLE binding
func Init() {
	ole.CoInitialize(0)
}

// NewDocuments Create a new document
func NewDocuments(v bool) *Documents {
	doc := &Documents{}
	return doc.NewDocument(v)
}

//Uninitialize ...
func Uninitialize() {
	ole.CoUninitialize()
}

func setupRegistry(version string, i uint32) {

	k, err := registry.OpenKey(registry.CURRENT_USER, fmt.Sprintf("Software\\Microsoft\\Office\\%s\\Word\\Security", version), registry.QUERY_VALUE|registry.SET_VALUE)
	if err != nil {
		panic(err)
	}
	if err := k.SetDWordValue("AccessVBOM", i); err != nil {
		panic(err)
	}

	val, _, _ := k.GetIntegerValue("AccessVBOM")
	fmt.Printf("Software\\Microsoft\\Office\\%s\\Word\\Security to %v\n", version, val)

	if err := k.Close(); err != nil {
		panic(err)
	}
}

////////// DOCUMENTS METHODS //////////

// NewDocument Create a new Word document
func (d *Documents) NewDocument(v bool) *Documents {
	unknown, err := oleutil.CreateObject("Word.Application")
	if err != nil {
		panic("Cannot create Word.Application")
	}

	word, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		panic("Cannot QueryInterface of Word")
	}

	d.Application = &Application{_Application: word,
		Options: &Options{_Options: oleutil.MustGetProperty(word, "Options").ToIDispatch()}}

	oleutil.PutProperty(d.Application._Application, "Visible", v)

	d._Documents = oleutil.MustGetProperty(d.Application._Application, "Documents").ToIDispatch()
	d.Application.Version = oleutil.MustGetProperty(d.Application._Application, "Version").ToString()

	setupRegistry(d.Application.Version, 1)

	return d
}

// AddDocument Add a new page to the document. Also init VBProject and subsequent
func (d *Documents) AddDocument() *Document {
	d.Document = &Document{_Document: oleutil.MustCallMethod(d._Documents, "Add").ToIDispatch()}

	d.Document.Application = d.Application

	d.Document.VBProject = &VBProject{
		_VBProject: oleutil.MustGetProperty(d.Document._Document, "VBProject").ToIDispatch(),
	}

	d.Document.VBProject.GetVBComponentsObject()

	return d.Document
}

// GetDocumentsObject Get documents
func (d *Documents) GetDocumentsObject() *ole.IDispatch {
	if d._Documents == nil {
		panic("No documents present in Documents")
	}
	return d._Documents
}

//Save saves the doc
func (d *Documents) Save() {
	oleutil.MustCallMethod(d._Documents, "Save")
	d.Document.Save()
}

//Close Close doc
func (d *Documents) Close() {

	oleutil.MustCallMethod(d._Documents, "Close", false)
	oleutil.MustCallMethod(d.Application._Application, "Quit")
	d.Application._Application.Release()

	setupRegistry(d.Application.Version, 0)
}

////////// DOCUMENT METHODS //////////

// GetVBProjectObject Get VBProject
func (d *Document) GetVBProjectObject() *ole.IDispatch {
	if d.VBProject._VBProject == nil {
		panic("No VBProject!")
	}
	return d.VBProject._VBProject
}

// GetDocumentObject Get document
func (d *Document) GetDocumentObject() *ole.IDispatch {
	if d._Document == nil {
		panic("No document present in Documents")
	}
	return d._Document
}

//UndoClear ...
func (d *Document) UndoClear() {
	oleutil.MustCallMethod(d._Document, "UndoClear").ToIDispatch()
}

//Repaginate ...
func (d *Document) Repaginate() {
	oleutil.MustCallMethod(d._Document, "Repaginate").ToIDispatch()
}

//RemoveDocumentInformation ...
func (d *Document) RemoveDocumentInformation(param ...interface{}) {
	oleutil.MustCallMethod(d._Document, "RemoveDocumentInformation", param...).ToIDispatch()
}

//SaveAs saves the doc
func (d *Document) SaveAs(path string) {
	oleutil.MustCallMethod(d._Document, "SaveAs2", path, 13, false)
}

//Save saves the doc
func (d *Document) Save() {
	oleutil.MustPutProperty(d._Document, "Saved", true)
	oleutil.MustCallMethod(d._Document, "Save")
}

////////// APPLICATION METHODS //////////

// SetOption Set the option to value
func (a *Application) SetOption(option string, param ...interface{}) {
	oleutil.MustPutProperty(a._Application, option, param...).ToIDispatch()
}

// ScreenRefresh ...
func (a *Application) ScreenRefresh() {
	oleutil.MustCallMethod(a._Application, "ScreenRefresh").ToIDispatch()
}

// GetApplicationObject Get documents
func (a *Application) GetApplicationObject() *ole.IDispatch {
	if a._Application == nil {
		panic("No documents present in Documents")
	}
	return a._Application
}

////////// OPTIONS METHODS //////////

// SetOption Set the option to value
func (o *Options) SetOption(option string, param ...interface{}) {
	oleutil.MustPutProperty(o._Options, option, param...).ToIDispatch()
}
