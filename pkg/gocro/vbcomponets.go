package gocro

import (
	"errors"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const (
	//MODULE type
	MODULE int = 1
	//CLASS type
	CLASS int = 2
	//FORM type
	FORM int = 3
)

// VBComponents Holds VB conponents
type VBComponents struct {
	_VBComponents *ole.IDispatch
	Components    map[string]*VBComponent
	Forms         map[string]*Form
}

// VBComponent Holds VB conponent
type VBComponent struct {
	_VBComponent *ole.IDispatch
	CodeModule   *CodeModule
}

// FetchItems Get all VBComponents and add them to a map[string]*ole.IDispatch
func (v *VBComponents) fetchItems() {

	//init map
	v.Components = make(map[string]*VBComponent)
	v.Forms = make(map[string]*Form)

	newEnum, _ := v._VBComponents.CallMethod("_NewEnum")
	enum, _ := newEnum.ToIUnknown().IEnumVARIANT(ole.IID_IEnumVariant)

	defer newEnum.Clear()
	defer enum.Release()

	for item, length, _ := enum.Next(1); length > 0; item, length, _ = enum.Next(1) {
		m := oleutil.MustGetProperty(item.ToIDispatch(), "Name")

		v.Components[m.ToString()] = &VBComponent{_VBComponent: item.ToIDispatch(),
			CodeModule: &CodeModule{_CodeModule: oleutil.MustGetProperty(item.ToIDispatch(), "codeModule").ToIDispatch()}}
	}
}

// GetVBComponent Return the VBComponent by name if it exists
func (v *VBComponents) GetVBComponent(name string) (*VBComponent, error) {

	c, ok := v.Components[name]
	if !ok {
		return nil, errors.New("VBComponent not found")
	}

	return c, nil
}

// GetForm Return the Form by name if it exists
func (v *VBComponents) GetForm(name string) (*Form, error) {

	c, ok := v.Forms[name]
	if !ok {
		return nil, errors.New("VBComponent not found")
	}

	return c, nil
}

// GetVBComponentObject Return the VBComponent *ole.IDispatch by name if it exists
func (v *VBComponents) GetVBComponentObject(name string) *ole.IDispatch {

	c, ok := v.Components[name]
	if !ok {
		return nil
	}

	return c._VBComponent
}

// AddVBComponent Add a new Module
func (v *VBComponents) AddVBComponent(name string, cType int) *VBComponent {
	comp := oleutil.MustCallMethod(v._VBComponents, "Add", cType).ToIDispatch()

	comp.PutProperty("Name", name)
	v.Components[name] = &VBComponent{_VBComponent: comp,
		CodeModule: &CodeModule{_CodeModule: oleutil.MustGetProperty(comp, "codeModule").ToIDispatch()}}

	return v.Components[name]
}

// AddNewForm Add a new Module
func (v *VBComponents) AddNewForm(name string) *Form {
	comp := oleutil.MustCallMethod(v._VBComponents, "Add", FORM).ToIDispatch()

	comp.PutProperty("Name", name)
	tmp := &VBComponent{_VBComponent: comp,
		CodeModule: &CodeModule{_CodeModule: oleutil.MustGetProperty(comp, "codeModule").ToIDispatch()}}

	v.Forms[name] = &Form{*tmp, &Designer{}, make(map[string]*Element)}

	v.Forms[name].getDesigner()
	v.Forms[name].Designer.getControl()

	return v.Forms[name]
}

//////// VBCOMPONENT ////////

// GetVBComponentObject Return the VBComponent *ole.IDispatch
func (v *VBComponent) GetVBComponentObject() *ole.IDispatch {
	return v._VBComponent
}

//SetName sets the name of the VBComponent. SetProperty("Name") returns Read Only property....
func (v *VBComponent) SetName(name string) {
	oleutil.MustPutProperty(v._VBComponent, "Name", name)
}

//GetName sets the name of the VBComponent. GetProperty("Name") returns Doc name and not VB name....
func (v *VBComponent) GetName() string {
	return oleutil.MustGetProperty(v._VBComponent, "Name").ToString()
}

//SetProperty to set a property
func (v *VBComponent) SetProperty(name string, param ...interface{}) {

	prop := oleutil.MustGetProperty(v._VBComponent, "Properties").ToIDispatch()

	newEnum, _ := prop.CallMethod("_NewEnum")
	enum, _ := newEnum.ToIUnknown().IEnumVARIANT(ole.IID_IEnumVariant)

	defer newEnum.Clear()
	defer enum.Release()

	for item, length, _ := enum.Next(1); length > 0; item, length, _ = enum.Next(1) {
		m := oleutil.MustGetProperty(item.ToIDispatch(), "Name")
		if m.ToString() == name {
			oleutil.MustPutProperty(item.ToIDispatch(), "Value", param...)
		}
	}
}

//GetProperty to set a property
func (v *VBComponent) GetProperty(name string, param ...interface{}) string {

	prop := oleutil.MustGetProperty(v._VBComponent, "Properties").ToIDispatch()

	newEnum, _ := prop.CallMethod("_NewEnum")
	enum, _ := newEnum.ToIUnknown().IEnumVARIANT(ole.IID_IEnumVariant)

	defer newEnum.Clear()
	defer enum.Release()

	for item, length, _ := enum.Next(1); length > 0; item, length, _ = enum.Next(1) {
		m := oleutil.MustGetProperty(item.ToIDispatch(), "Name")
		if m.ToString() == name {
			return oleutil.MustGetProperty(item.ToIDispatch(), "Value", param...).ToString()
		}
	}

	return ""
}
