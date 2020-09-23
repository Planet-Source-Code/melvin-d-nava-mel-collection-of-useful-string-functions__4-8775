<div align="center">

## MEL :: Collection of Useful String Functions


</div>

### Description

Collection of functions to deal with strings, including to capitalize, format, compare, parse links, check arrays and others. This is an added value to any existing string library. If you have comments or suggestions please do. Let me know this has been useful to you by voting for me or linking back to my website.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Melvin D\. Nava](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/melvin-d-nava.md)
**Level**          |Intermediate
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__4-35.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/melvin-d-nava-mel-collection-of-useful-string-functions__4-8775/archive/master.zip)





### Source Code

```
<%
'
' client user agent (navigator)
' & operating system variables
' use it in your apps
ua = Request.ServerVariables("HTTP_USER_AGENT")
os = Request.ServerVariables("HTTP_UA_OS")
'
' short replacement for
' Response.Write just use:
' Write("Hello World!")
	Public Function Write(ByVal str)
	Response.Write(str & vbCrLf)
	End Function
'
' uppercase first words of a string
	Public Function Capitalize(ByVal str)
	DIM arrTemp, strTemp, i
	arrTemp = Split(str, " ")
	For i = 0 to Ubound(arrTemp)
		strTemp = strTemp & " " & UCase(Left(arrTemp(i),1)) & LCase(Mid(arrTemp(i),2))
	Next
	Capitalize = strTemp
	End Function
'
' uppercase first letter of a string
	Public function CapFirstWord(str)
	CapFirstWord = UCase(Left(str,1)) & Lcase(Mid(str,2))
	End Function
'
' evals an expression an returns
' true or false. almost just like
' Visual Basic. Use only variables
' in trueR and falseR
	Public Function IIf(expr,trueR,falseR)
	On Error Resume Next
	If Eval(expr) Then IIf = trueR Else IIf = falseR
	End Function
'
' returns true if the variable
' has any part of the string. i.e.:
' IsExpr("in god name", "in name")
' returns True
	Function IsExpr(patrn, strng)
	Dim regEx, retVal				' Create variable.
	Set regEx = New RegExp			' Create regular expression.
	regEx.Pattern = patrn			' Set pattern.
	regEx.IgnoreCase = True			' Set case sensitivity.
	retVal = regEx.Test(strng)		' Execute the search test.
	If retVal Then IsExpr = True Else IsExpr = False
	End Function
'
' it bolds any strings[strToBold]
' found on another string[strText].
' Useful when implementing a
' search engine
	Public Function BoldFoundString(ByVal strText, ByVal strToBold)
	On Error Resume Next
	DIM strTemp
	strTemp = Replace(strText,strToBold,"<b>" & strToBold & "</b>")
	BoldFoundString = strTemp
	End Function
'
' clear quotes from strings
	Public Function ClearQuotes(ByVal str)
	If Not Isnull(str) And str <> "" Then
		ClearQuotes = Replace(str,"'"," ")
	Else
		ClearQuotes = str
	End If
	End Function
'
' repeats a string(strC) a number
' of times(intT). e.i:
' RepeatString("a", 4) returns aaaa
	Public Function RepeatString(strC, intT)
	DIM i, strTemp
	strTemp = strC
	If intT > 1 Then
		For i = 1 To intT+1
			strTemp = strTemp & strC
		Next
	End If
	RepeatString = strTemp
	End Function
'
' check that every item in an
' array is not empty (true if OK)
	Public Function IsArrayNotEmpty(ByVal arr())
	DIM i
	For Each i In arr
		If Trim(Len(i)) < 1 Then
			IsArrayNotEmpty = False
			Exit Function
		Else
			IsArrayNotEmpty = True
		End If
	Next
	End Function
'
' two global variables needed
' by next func FormVariables2Arrays()
DIM arrColumns(1), arrValues(1)
'
' redimentions 2 global arrays
' and puts request.form variables
' into them (arrColumns-arrValues)
' this one is necesary to use some
' of my asp custom functions like
' DatabaseUpdate() and DatabaseInsert()
	Public Sub FormVariables2Arrays()
	DIM v, x, i
	v = 0
	For Each x In Request.Form
		v = v + 1
	Next
	v = v - 1
	REDIM arrColumns(v), arrValues(v)
	i = 0
	For Each x In Request.Form
		arrColumns(i)	= x
		arrValues(i)	= Request.Form(x)
		i = i + 1
	Next
	End Sub
'
' find http urls in a string and
' returns the same the string with
' ready links
' TODO: make it parse emails
	Public Function ParseStringLinks(strInput)
	DIM iCurrentLocation, iLinkStart, iLinkEnd, strOutput, strLinkText
	iCurrentLocation = 1
	Do While InStr(iCurrentLocation, strInput, "http://", 1) <> 0
		iLinkStart = InStr(iCurrentLocation, strInput, "http://", 1)
		iLinkEnd = InStr(iLinkStart, strInput, " ", 1)
		If iLinkEnd = 0 Then iLinkEnd = Len(strInput) + 1
		Select Case Mid(strInput, iLinkEnd - 1, 1)
			Case ".", "!", "?"
			iLinkEnd = iLinkEnd - 1
		End Select
		strOutput = strOutput & Mid(strInput, iCurrentLocation, iLinkStart - iCurrentLocation)
		strLinkText = Mid(strInput, iLinkStart, iLinkEnd - iLinkStart)
		strOutput = strOutput & Replace("<¿xax? href="""&strLinkText&""">"&strLinkText&"</¿xax?>","¿xax?","a")
		iCurrentLocation = iLinkEnd
	Loop
	strOutput = strOutput & Mid(strInput, iCurrentLocation)
	ParseStringLinks = strOutput
	End Function
%>
```

