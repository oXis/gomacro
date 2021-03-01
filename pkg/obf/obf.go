package obf

import (
	"fmt"
	"math/rand"
	"strings"
	"time"

	regexp2 "github.com/dlclark/regexp2"
	"github.com/tjarratt/babble"
)

func init() {
	rand.Seed(time.Now().UTC().UnixNano())
}

const letterBytes = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

var babbler = babble.NewBabbler()
var forbiddenWords = []string{"AddHandler", "AddressOf", "Alias", "And", "AndAlso", "As", "Boolean", "ByRef", "Byte", "ByVal", "Call", "Case", "Catch", "CBool", "CByte", "CChar", "CDate", "CDbl", "CDec", "Char", "CInt", "CObj", "Const", "Continue", "CSByte", "CShort", "CSng", "CStr", "CType", "CUInt", "CULng", "CUShort", "Date", "Decimal", "Declare", "Default", "Delegate", "Dim", "DirectCast", "Do", "Double", "Each", "Else", "ElseIf", "End", "Statement", "End", "EndIf", "Enum", "Erase", "Error", "Event", "Exit", "False", "Finally", "For", "Eac", "Next", "Friend", "Function", "Get", "GetType", "GetXMLNamespace", "Global", "GoSub", "GoTo", "Handles", "If", "Implements", "Imports", "In", "Inherits", "Integer", "Interface", "Is", "IsNot", "Let", "Lib", "Like", "Long", "Loop", "Me", "Mod", "Module", "MustInherit", "MustOverride", "MyBase", "MyClass", "NameOf", "Namespace", "Narrowing", "Constraint", "New", "Not", "Nothing", "NotInheritable", "NotOverridable", "Object", "Of", "On", "Operator", "Option", "Optional", "Or", "OrElse", "Out", "Overloads", "Overridable", "Overrides", "ParamArray", "Partial", "Private", "Property", "Protected", "Public", "RaiseEvent", "ReadOnly", "ReDim", "REM", "RemoveHandler", "Resume", "Return", "SByte", "Select", "Set", "Shadows", "Shared", "Short", "Single", "Static", "Step", "Stop", "String", "Structure", "Sub", "SyncLock", "Then", "Throw", "To", "True", "Try", "TryCast", "TypeOf", "UInteger", "ULong", "UShort", "Using", "Variant", "Wend", "When", "While", "Widening", "With", "WithEvents", "WriteOnly", "Xor"}

// RandStringBytes ...
func RandStringBytes(n int) string {
	b := make([]byte, n)
	for i := range b {
		b[i] = letterBytes[rand.Intn(len(letterBytes))]
	}
	return string(b)
}

// RandWord ...
func RandWord() string {
	// optionally set the number of words you want

	babbler.Count = 1
	word := strings.ToLower(babbler.Babble())

	for _, w := range forbiddenWords {
		if word == strings.ToLower(w) {
			return RandWord()
		}
	}

	forbiddenWords = append(forbiddenWords, word)

	return word
}

func removeDuplicatesFromSlice(s []string) []string {
	m := make(map[string]bool)
	for _, item := range s {
		if _, ok := m[item]; !ok {
			m[item] = true
		}
	}

	var result []string
	for item := range m {
		result = append(result, item)
	}
	return result
}

func regexp2FindAllString(re *regexp2.Regexp, s string) [][]regexp2.Group {

	var matches [][]regexp2.Group

	m, _ := re.FindStringMatch(s)
	for m != nil {
		matches = append(matches, m.Groups())
		m, _ = re.FindNextMatch(m)
	}
	return matches
}

func getFunctions(code string) []string {
	// functions, _ := regexp.MatchString(`(Function|Sub)[ ]+(\w+)\(`, vbcode)
	functions := regexp2.MustCompile(`(Function|Sub)[ ]+(\w+)\(`, 0)
	gprs := regexp2FindAllString(functions, code)

	var ret []string

	for _, g := range gprs {
		// for _, g2 := range g {
		// 	fmt.Printf("%s\n", g2.String())
		// }
		fmt.Printf("Function Name: %s\n", g[2].String())

		ret = append(ret, g[2].String())
	}

	return ret
}

func getFunctionsParameters(code string) []string {
	// parameters, _ := regexp.MatchString(`(?:Function|Sub)[ ]+\w+\(((?:\w+[ ]+As[ ]+\w+(?:, )*)*)\)`, vbcode)
	parameters := regexp2.MustCompile(`(?:Function|Sub)[ ]+\w+\(((?:\w+[ ]+As[ ]+\w+(?:, )*)*)\)`, 0)
	gprs := regexp2FindAllString(parameters, code)

	var ret []string

	for _, g := range gprs {
		// for _, g2 := range g {
		// 	fmt.Printf("%s\n", g2.String())
		// }
		parameterNames := regexp2.MustCompile(`(?:(\w+)[ ]+As[ ]+\w+(?:, )*)`, 0)
		gprs = regexp2FindAllString(parameterNames, g[1].String())
		for _, g2 := range gprs {
			// for _, g2 := range g {
			// 	fmt.Printf("%s\n", g2.String())
			// }
			fmt.Printf("Parameter name: %s\n", g2[1].String())
			ret = append(ret, g2[1].String())
		}
	}

	return ret
}

func getVariables(code string) []string {
	// variables, _ := regexp.MatchString(`^\s*(Dim|Set)[ ]+(\w+)`, vbcode)
	variables := regexp2.MustCompile(`[ ]?(Dim|Set|Public|For)[ ]+(?!Function)(\w+)`, 0)
	gprs := regexp2FindAllString(variables, code)

	var ret []string

	for _, g := range gprs {
		// for _, g2 := range g {
		// 	fmt.Printf("%s\n", g2.String())
		// }
		fmt.Printf("Var name: %s\n", g[2].String())
		ret = append(ret, g[2].String())
	}

	return ret
}

func getStrings(code string) []string {
	str := regexp2.MustCompile(`"(.+?)"`, 0)
	gprs := regexp2FindAllString(str, code)

	var ret []string

	for _, g := range gprs {
		// for _, g2 := range g {
		// 	fmt.Printf("%s\n", g2.String())
		// }
		fmt.Printf("str: %s\n", g[0].String())
		ret = append(ret, g[0].String())
	}

	return ret
}

func makeRange(min, max int) []int {
	a := make([]int, max-min+1)
	for i := range a {
		a[i] = min + i
	}
	return a
}

func shuffleString(str string) (string, string) {

	formatString := ""
	formatStringList := ""
	tmpStringList := []string{}

	str = str[1 : len(str)-1] // remove quotes

	for i := 0; i < len(str); {

		step := rand.Intn(4) + 2
		if (i + step) > len(str) {
			step = len(str) - i
		}

		tmpStringList = append(tmpStringList, str[i:i+step])
		i = i + step
	}

	tmpStringList2 := make([]string, len(tmpStringList))

	// Insure list is always scrambled
	randIndex := rand.Perm(len(tmpStringList))
	for {
		b := false
		for i, j := range makeRange(0, len(tmpStringList)-1) {
			if randIndex[i] != j {
				b = true
				break
			}
		}
		if b {
			break
		}
		randIndex = rand.Perm(len(tmpStringList))
	}

	for _, s := range randIndex {
		formatString += fmt.Sprintf("{%v}", s)
		// formatStringList += fmt.Sprintf(`"%v",`, tmpStringList[s])
	}

	for i, s := range randIndex {
		tmpStringList2[s] = tmpStringList[i]
	}

	for _, s := range tmpStringList2 {
		formatStringList += fmt.Sprintf(`"%v",`, s)
	}

	formatStringList = formatStringList[:len(formatStringList)-1]

	return formatString, formatStringList
}

// ObfuscateVBCode ...
func ObfuscateVBCode(code string, objFunc, objParam, objVar, objString bool) (string, map[string]string, map[string]string, map[string]string, map[string]string) {

	funcMap := make(map[string]string)
	if objFunc {
		functions := removeDuplicatesFromSlice(getFunctions(code))
		for _, s := range functions {
			funcMap[s] = RandWord()
		}
	}

	paramMap := make(map[string]string)
	if objParam {
		parameters := removeDuplicatesFromSlice(getFunctionsParameters(code))
		for _, p := range parameters {
			paramMap[p] = RandWord()
		}
	}

	varMap := make(map[string]string)
	if objVar {
		variables := removeDuplicatesFromSlice(getVariables(code))
		for _, p := range variables {
			varMap[p] = RandWord()
		}
	}

	stringMap := make(map[string]string)
	if objString {
		str := removeDuplicatesFromSlice(getStrings(code))
		for _, p := range str {
			formatString, formatStringList := shuffleString(p)
			stringMap[p] = fmt.Sprintf(`Format("%v", %v)`, formatString, formatStringList)
		}
	}

	return code, funcMap, paramMap, varMap, stringMap
}

// ReplaceAllInCode ...
func ReplaceAllInCode(code string, funcMap, paramMap, varMap, stringMap map[string]string) string {

	for s := range funcMap {
		fmt.Printf("Replacing %s with %s\n", s, funcMap[s])
		code = strings.ReplaceAll(code, s, funcMap[s])
	}

	for p := range paramMap {
		fmt.Printf("Replacing %s with %s\n", p, paramMap[p])
		code = strings.ReplaceAll(code, p, paramMap[p])
	}

	for p := range varMap {
		fmt.Printf("Replacing %s with %s\n", p, varMap[p])
		code = strings.ReplaceAll(code, p, varMap[p])
	}

	for p := range stringMap {
		fmt.Printf("Replacing %s with %s\n", p, stringMap[p])
		code = strings.ReplaceAll(code, p, stringMap[p])
	}

	return code
}
