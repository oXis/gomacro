package obf

import (
	"fmt"
	"math/rand"
	"strings"
	"time"

	regexp2 "github.com/dlclark/regexp2"
)

func init() {
	rand.Seed(time.Now().UTC().UnixNano())
}

const letterBytes = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

// RandStringBytes ...
func RandStringBytes(n int) string {
	b := make([]byte, n)
	for i := range b {
		b[i] = letterBytes[rand.Intn(len(letterBytes))]
	}
	return string(b)
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

// ObfuscateVBCode ...
func ObfuscateVBCode(code string, size int) (string, map[string]string, map[string]string, map[string]string) {

	functions := removeDuplicatesFromSlice(getFunctions(code))
	parameters := removeDuplicatesFromSlice(getFunctionsParameters(code))
	variables := removeDuplicatesFromSlice(getVariables(code))

	funcMap := make(map[string]string)
	for _, s := range functions {
		funcMap[s] = RandStringBytes(size)
		code = strings.ReplaceAll(code, s, funcMap[s])
	}

	paramMap := make(map[string]string)
	for _, p := range parameters {
		paramMap[p] = RandStringBytes(size)
		code = strings.ReplaceAll(code, p, paramMap[p])
	}

	varMap := make(map[string]string)
	for _, p := range variables {
		varMap[p] = RandStringBytes(size)
		code = strings.ReplaceAll(code, p, varMap[p])
	}

	return code, funcMap, paramMap, varMap
}
