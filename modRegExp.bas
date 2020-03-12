Attribute VB_Name = "modRegExp"
Option Explicit

Private Sub AddInMenuProperties()
    ' Custom function for changing file properties (not used during run time)
    ActiveWorkbook.BuiltinDocumentProperties("Title").Value = "Regular Expression Functions 1.0"
    ActiveWorkbook.BuiltinDocumentProperties("Comments").Value = "Regular expression functions from Google Sheets."
End Sub

Public Function REGEXMATCH(text As String, _
    regular_expression As String) As Boolean
Attribute REGEXMATCH.VB_Description = "Whether a piece of text matches a regular expression."
Attribute REGEXMATCH.VB_ProcData.VB_Invoke_Func = " \n21"

    ' Create regular expression object
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    With objRegExp
        .Global = True
        .MultiLine = True
        .IgnoreCase = False 'Case-sensitive
        .Pattern = regular_expression
    End With
    
    ' Return True/False based on matching
    REGEXMATCH = objRegExp.test(text)

End Function

Public Function REGEXEXTRACT(text As String, _
    regular_expression As String) As String
Attribute REGEXEXTRACT.VB_Description = "Extracts matching substrings according to a regular expression."
Attribute REGEXEXTRACT.VB_ProcData.VB_Invoke_Func = " \n21"

    ' Create regular expression object
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    With objRegExp
        .Global = False ' Only first match will be returned
        .MultiLine = True
        .IgnoreCase = False ' Case-sensitive
        .Pattern = regular_expression
    End With
    
    ' Return extracted string
    REGEXEXTRACT = objRegExp.Execute(text).Item(0).Value

End Function

Public Function REGEXREPLACE(text As String, _
    regular_expression As String, replacement As String) As String
Attribute REGEXREPLACE.VB_Description = "Replaces part of a text string with a different text string using regular expressions."
Attribute REGEXREPLACE.VB_ProcData.VB_Invoke_Func = " \n21"

    ' Create regular expression object
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    With objRegExp
        .Global = True
        .MultiLine = True
        .IgnoreCase = False ' Case-sensitive
        .Pattern = regular_expression
    End With
    
    ' Return replaced string
    REGEXREPLACE = objRegExp.Replace(text, replacement)

End Function

Sub RegExpArg()
    
    Application.MacroOptions "REGEXMATCH", "Whether a piece of text matches a regular expression.", , , , , "RegExp", , , , _
        Array("The text to be tested against the regular expression.", _
        "The regular expression to test the text against.")
    
    Application.MacroOptions "REGEXEXTRACT", "Extracts matching substrings according to a regular expression.", , , , , "RegExp", , , , _
        Array("The input text.", _
        "The first part of text that matches this expression will be returned.")
    
    Application.MacroOptions "REGEXREPLACE", "Replaces part of a text string with a different text string using regular expressions.", , , , , "RegExp", , , , _
        Array("The text, a part of which will be replaced.", _
        "The regular expression. All matching instances in text will be replaced.", _
        "The text which will be inserted into the original text.")
    
End Sub
