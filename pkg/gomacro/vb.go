package gomacro

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// VBProject Holds VB projects
type VBProject struct {
	_VBProject   *ole.IDispatch
	Name         string
	VBComponents VBComponents
}

// PutProperty Set a property
func (v *VBProject) PutProperty(name string, params ...interface{}) {
	oleutil.MustPutProperty(v._VBProject, name, params...)
}

// GetProperty Get VBProject prperty
func (v *VBProject) GetProperty(name string, params ...interface{}) *ole.IDispatch {
	return oleutil.MustGetProperty(v._VBProject, name, params...).ToIDispatch()
}

// SetName Set name of VBProject
func (v *VBProject) SetName(name string) {
	v.PutProperty("Name", name)
	v.Name = name
}

// GetName Get name of VBProject
func (v *VBProject) GetName(name string) string {
	return v.Name
}

func (v *VBProject) initVBComponentsObject() {

	if v.VBComponents._VBComponents == nil {
		v.VBComponents._VBComponents = v.GetProperty("VBComponents")
		v.VBComponents.fetchItems()
	}
}

// GetVBComponentsObject Get GetVBComponents
func (v *VBProject) GetVBComponentsObject() *ole.IDispatch {

	if v.VBComponents._VBComponents == nil {
		v.initVBComponentsObject()
	}

	return v.VBComponents._VBComponents
}
