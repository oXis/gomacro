package gomacro

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

//Application Holds the OLE app
type Application struct {
	_Application *ole.IDispatch
	Version      string
	Options      *Options
	Selection    *Selection
}

//Options Holds the OLE app
type Options struct {
	_Options *ole.IDispatch
}

type Selection struct {
	_Selection *ole.IDispatch
}

////////// APPLICATION METHODS //////////

func (a *Application) Init() {

	a.Version = oleutil.MustGetProperty(a._Application, "Version").ToString()
	a.Selection = &Selection{_Selection: oleutil.MustGetProperty(a._Application, "Selection").ToIDispatch()}
}

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

////////// SELECTION METHODS //////////

func (s *Selection) GetSelection() *ole.IDispatch {
	return s._Selection
}
