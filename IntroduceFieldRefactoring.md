This refactoring moves a variable declaration up to the module level scope. For example, in the below code, we can right-click on the declaration of the local variable `inputs` and select `Rubberduck -> Refactor -> Introduce Field`:

```
Option Explicit

Public Function ParsePascal(prog As String, inputRange As Range) As String
    Dim inputs As Scripting.Dictionary
    Set inputs = New Scripting.Dictionary
    
    Dim r As Long
    For r = 1 To inputRange.Rows().Count
        inputs(inputRange.Cells(r, 1).value) = inputRange.Cells(r, 2).value
    Next r
    
    ParsePascal = ParsePascalDict(prog, inputs)
End Function
```

The declaration will be removed from the current scope and added to the end of the module level declarations section as a new `Private` field. The result of refactoring the above code will be:

```
Option Explicit

Private inputs As Scripting.Dictionary
Public Function ParsePascal(prog As String, inputRange As Range) As String
    Set inputs = New Scripting.Dictionary
    
    Dim r As Long
    For r = 1 To inputRange.Rows().Count
        inputs(inputRange.Cells(r, 1).value) = inputRange.Cells(r, 2).value
    Next r
    
    ParsePascal = ParsePascalDict(prog, inputs)
End Function
```