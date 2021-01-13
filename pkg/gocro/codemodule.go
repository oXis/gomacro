package gocro

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

//CodeModule Represents a codeModule
type CodeModule struct {
	_CodeModule *ole.IDispatch
}

//GetCodeModuleObject Returns CodeModule *ole.IDispatch
func (c *CodeModule) GetCodeModuleObject() *ole.IDispatch {
	return c._CodeModule
}

//AddFromString Add content to the code module
func (c *CodeModule) AddFromString(content string) {
	oleutil.MustCallMethod(c._CodeModule, "AddFromString", content).ToIDispatch()
}
