Attribute VB_Name = "RegExLib"

' Extracts a text given a RegEx Pattern
'
' String @text = text that will be tested
' String @pattern = regex pattern that will be used to test text

Function RegexExtract(text As String, pattern As String)
Attribute RegexExtract.VB_Description = "Extracts fraction from a text respecting a given RegEx Pattern developed by Felipe Dias Pereira. Meet me at github (@fdiaspp)."
Attribute RegexExtract.VB_ProcData.VB_Invoke_Func = " \n9"
        
    Dim regex As RegExp, matchObject As MatchCollection, result As String
    Set regex = New RegExp
    
    ' .Global property gives the ability for object to test all possibilites
    regex.Global = True
    
    ' .patter define the pattern to be used
    regex.pattern = pattern
    
    ' Executing the pattern against text. Returning an MatchCollection Object
    Set matchObject = regex.Execute(text)
    
    ' Initializing the result variable
    result = ""
    
    ' Realize the concatenation of all results retrived by RegExp.Execute() method
    For i = 0 To matchObject.Count - 1
        result = result & matchObject.Item(i)
    Next i
    
    ' return
    RegexExtract = result
    
End Function

Sub UDF()
Dim description As String, arg(2) As String

    description = "Extracts fraction from a text respecting a given RegEx Pattern developed by Felipe Dias Pereira. Meet me at github (@fdiaspp)."
    arg(0) = "Insert the text that will be tested against the RegEx Pattern"
    arg(1) = "The RegEx Pattern"
    

    Application.MacroOptions Macro:="RegexExtract", _
                         description:=description, _
                         StatusBar:="Please wait ...", _
                         ArgumentDescriptions:=arg
End Sub

