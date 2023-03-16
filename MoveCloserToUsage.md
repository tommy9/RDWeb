This refactoring moves a variable declaration closer to where it is used. For a variable declared within a procedure, it will be moved immediately above its first use. For fields declared at a module level and only used in one procedure, the declaration will be moved into that procedure, immediately above its first use. 

### Example
The below code has a field `paramFolder` declared at the module level but only used in the procedure `SubmitToProd`, so we can use Rubberduck to first move that declaration inside the procedure.
```
Private paramFolder As String

Sub SubmitToProd()
    Dim timeStamp As String
    Dim managerPath As String
    Dim managerName As String
    Dim response As VbMsgBoxResult
    
    timeStamp = Format(Now, "YYYYMMDD_hhmmss")
    paramFolder = "SubmittedToProd\"
    managerPath = ThisWorkbook.path & "\" & paramFolder
    managerName = CreateTimeStampedFileName(ThisWorkbook.Name, timeStamp)

    SaveParamVersion managerPath, managerName
    
    response = MsgBox("Refresh has been submitted to the production server", vbInformation + vbOKOnly)
End Sub
```

#### Moving into the procedure
Moving into the procedure is somewhat like the pposite of the _Introduce Field_ refactoring. On selecting the refactoring, Rubberduck will ask you if you want the new declaration to be a `Static` or `Dim` declaration. `Static` will allow the variable to persist it's value between calls to the procedure. If this isn't needed, then choose for the declaration to be `Dim`. The result of this first refactoring is:

```
Sub SubmitToProd()
    Dim timeStamp As String
    Dim managerPath As String
    Dim managerName As String
    Dim response As VbMsgBoxResult
    
    timeStamp = Format(Now, "YYYYMMDD_hhmmss")
    Dim paramFolder As String
    paramFolder = "SubmittedToProd\"
    managerPath = ThisWorkbook.path & "\" & paramFolder
    managerName = CreateTimeStampedFileName(ThisWorkbook.Name, timeStamp)

    SaveParamVersion managerPath, managerName
    
    response = MsgBox("Refresh has been submitted to the production server", vbInformation + vbOKOnly)
End Sub
```

#### Moving within a procedure
We notice that in the above refactoring, the declaration of `paramFolder` has not been moved to the set of declarations at the top of the procedure. Instead, it has been moved just before the usage. We can now select each of the declarations at the top of the procedure and move those closer to usage as well.

This is useful for grouping related lines of code to enhance readability and spotting opportunities to extract sections of code to a new method.

```
Sub SubmitToProd()
    
    Dim timeStamp As String
    timeStamp = Format(Now, "YYYYMMDD_hhmmss")
    Dim paramFolder As String
    paramFolder = "SubmittedToProd\"
    Dim managerPath As String
    managerPath = ThisWorkbook.path & "\" & paramFolder
    Dim managerName As String
    managerName = CreateTimeStampedFileName(ThisWorkbook.Name, timeStamp)

    SaveParamVersion managerPath, managerName
    
    Dim response As VbMsgBoxResult
    response = MsgBox("Refresh has been submitted to the production server", vbInformation + vbOKOnly)
End Sub
```

Rubberduck will update the indentation of the declaration to fit it's new location if that is different to the location it was moved from.


