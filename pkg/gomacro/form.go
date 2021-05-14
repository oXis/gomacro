package gomacro

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Form From struct
type Form struct {
	VBComponent
	Designer *Designer

	Items map[string]*Element
}

//Designer ...
type Designer struct {
	_Designer *ole.IDispatch
	Controls  *Controls
}

//Controls ...
type Controls struct {
	_Controls *ole.IDispatch
}

//Element ...
type Element struct {
	_Element *ole.IDispatch
}

func (f *Form) getDesigner() {
	f.Designer = &Designer{_Designer: oleutil.MustGetProperty(f._VBComponent, "Designer").ToIDispatch()}
}

func (d *Designer) getControl() {
	d.Controls = &Controls{_Controls: oleutil.MustGetProperty(d._Designer, "Controls").ToIDispatch()}
}

//Add Adds a new item to the form
func (f *Form) Add(element, name string) {
	f.Items[name] = &Element{oleutil.MustCallMethod(f.Designer.Controls._Controls, "Add", element, name).ToIDispatch()}
}

//GetElementObject Gets the element with name
func (f *Form) GetElementObject(name string) *ole.IDispatch {
	c, ok := f.Items[name]
	if !ok {
		return nil
	}

	return c._Element
}

//GetElement Gets the element with name
func (f *Form) GetElement(name string) *Element {
	c, ok := f.Items[name]
	if !ok {
		return nil
	}

	return c
}

//SetProperty to set properties of element
func (e *Element) SetProperty(name string, param ...interface{}) {
	oleutil.MustPutProperty(e._Element, name, param...).ToIDispatch()
}
