# GoMacro

Small utility and lib to create Word Documents with malicious macros.

- `pkg/gomacro` contains the lib code that interface with go-ole.
- `pkg/obf` contains code to obfuscate VB code.
- `cmd/main` contains a complete example on how to use gomacro lib to create a Word Doc with a macro.
- `resources` contains macro VB code

The main file is heavily commented so you can follow the steps to make a doc.

Blogpost [here](https://oxis.github.io/GoMacro,-a-small-utility-to-create-Word-macros-with-Go/)
