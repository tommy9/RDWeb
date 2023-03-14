This refactoring moves a variable declaration to be a parameter to the enclosing method. For example, in the below code, we can right-click on the declaration of the local variable `program` and select `Rubberduck -> Refactor -> Introduce Parameter`:

```
Private Sub TestLexer()
    Dim program As String
    program = "IF t > 0 THEN 1 ELSE 2.3"
    
    Dim lex As Lexer
    Set lex = Lexer.Create(program)
    
    Dim t As token
    Do
        Set t = lex.GetNextToken
        Debug.Print t.ToString
    Loop Until t.tokentype = Tokens.tEOF
End Sub
```

The declaration will be removed and a new `ByVal` parameter added to the argument list. The result of refactoring the above code will be:

```
Private Sub TestLexer(ByVal program As String)
    program = "IF t > 0 THEN 1 ELSE 2.3"
    
    Dim lex As Lexer
    Set lex = Lexer.Create(program)
    
    Dim t As token
    Do
        Set t = lex.GetNextToken
        Debug.Print t.ToString
    Loop Until t.tokentype = Tokens.tEOF
End Sub
```

Note that further updates will then be needed to your program:
 - updating all calling functions to pass in the new parameter
 - removing the variable assignment from the modified method so that the passed in value gets used instead
