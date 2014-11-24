Option Explicit

Function GetExchange(Currency1 As String, Currency2 As String) As Variant
Dim Search As String
Dim fetchStr As String
Dim p As String

If Len(Currency1) = 3 And Len(Currency2) = 3 Then
    Search = Currency1 & Currency2
    fetchStr = Fetcher(Search)
    p = CLng(Price(fetchStr))
Else
    p = "Error"
End If

GetExchange = p

End Function

Function Fetcher(Search As String) As String
Dim IE As New InternetExplorer
Dim URL As String
Dim Doc As HTMLDocument
Dim str As String
Dim Price As Variant

URL = "http://finance.yahoo.com/webservice/v1/symbols/" & Search & "=X/quote?format=json"
 
IE.navigate URL
 
Do
   DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

Set Doc = IE.document
str = Doc.body.innerText
 
Fetcher = str
End Function

Function jsonDecodePossible(jsonString As Variant) As Boolean
Dim sc As Object
Dim jsonDecode
Dim b As Boolean
 
b = False
On Error GoTo weiter:
Set sc = CreateObject("ScriptControl"): sc.Language = "JScript"
Set jsonDecode = sc.Eval("(" + jsonString + ")")
b = True

weiter:
jsonDecodePossible = b
End Function

Function jsonDecode(jsonString As Variant)
Dim sc As Object
Dim b As Boolean
 
Set sc = CreateObject("ScriptControl"): sc.Language = "JScript"
Set jsonDecode = sc.Eval("(" + jsonString + ")")

End Function

Function Price(jsonStr As Variant)
Dim json

If jsonDecodePossible(jsonStr) = True Then
    Set json = jsonDecode(jsonStr)
    Price = json.Price
Else
    Price = "Error"
End If
End Function
