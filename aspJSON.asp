<%
'July 2012 - Version 1.0 by Gerrit van Kuipers
'Updated by Francesco Rombecchi - Siti e Grafica - http://www.sitiegrafica.it
Class aspJSON
	Public data
	Private p_JSONstring
	Private p_datatype

	Private Sub Class_Initialize()
		Set data = Collection()
		p_datatype = "{}"
	End Sub

	Private Sub Class_Terminate()
		Set data = Nothing
	End Sub

	Public Function loadJSON(strInput)
		dim lines, currentlevel, line, currentkey, currentvalue, in_string, in_escape, i_tmp, char
		
		p_JSONstring = CleanUpJSONstring(Trim(strInput))
		lines = Split(p_JSONstring, vbCrLf)

		Dim level(99)
		currentlevel = 1
		Set level(currentlevel) = data
		For Each line In lines
			currentkey = ""
			currentvalue = ""
			If Instr(line, ":") > 0 Then
				'"created":"2010-04-30 09:20:09"

				in_string = False
				in_escape = False
				For i_tmp = 1 To Len(line)
					If in_escape Then
						in_escape = False
					Else
						char = Mid(line, i_tmp, 1)
						Select Case char
							Case """"
								in_string = Not in_string
							Case ":"
								If Not in_escape Then
									currentkey = Left(line, i_tmp - 1)
									currentvalue = Mid(line, i_tmp + 1)
									Exit For
								End If
							Case "\"
								in_escape = True
						End Select
					End If
				Next
				currentkey = Strip(JSONDecode(currentkey), """")
				If Not level(currentlevel).exists(currentkey) Then level(currentlevel).Add currentkey, ""
			End If
			If Instr(line,"{") > 0 Or Instr(line,"[") > 0 Then
				If Len(currentkey) = 0 Then currentkey = level(currentlevel).Count
				Set level(currentlevel).Item(currentkey) = Collection()
				Set level(currentlevel + 1) = level(currentlevel).Item(currentkey)
				currentlevel = currentlevel + 1
				currentkey = ""
			ElseIf Instr(line,"}") > 0 Or Instr(line,"]") > 0 Then
				currentlevel = currentlevel - 1
			ElseIf Len(Trim(line)) > 0 Then
				if Len(currentvalue) = 0 Then currentvalue = getJSONValue(line)
				currentvalue = getJSONValue(currentvalue)

				If Len(currentkey) = 0 Then currentkey = level(currentlevel).Count
				level(currentlevel).Item(currentkey) = currentvalue
			End If
		Next
	End Function
	
	
	Public Function convertCSS(CSSString)
		dim App, Elements, Element, Pos, Risultato, Params, Param
		Risultato = ""
		if not IsNull(CSSString) then
			App = Replace(CSSString, VbCrLf, "")
			Risultato = "{"
			Elements = split(App, "}")
			for each Element in Elements
				Pos = InStr(Element, "{")
				if Pos > 0 then
					Risultato = Risultato & """" & Trim(Left(Element, Pos-1))  & """: {"
					Element = Right(Element, Len(Element)-Pos)								
					Params = split(Element, ";")
					App = ""
					for each Param in Params
						if Trim(Param) <> "" then
							Pos = InStr(Param, ":")						
							if Pos > 0 then
								App = App & """" & Trim(Left(Param, Pos-1))  & """: """ & Trim(Right(Param, Len(Param)-Pos)) & """, "
							end if
						end if					
					next
					if Len(App) > 1 then Risultato = Risultato & Left(App, Len(App)-2)
					Risultato = Risultato & "}, "
				end if
			next
			if Len(Risultato) > 1 then Risultato = Left(Risultato, Len(Risultato)-2)
			Risultato = Risultato & "}"
		end if		
		oJSON.loadJSON(Risultato)
		convertCSS = Risultato
	end Function
	
	
	Public Function FormattaJSON(JSONString)
		dim App, CR
		App = JSONString
		if Trim(App) <> "" then		
			App = Replace(App, "{", "  {" & "<blockquote>")		
			App = Replace(App, "}", "</blockquote>}")
			App = Replace(App, ",", ",<br>")
		end if
		FormattaJSON = App
	end Function
	
	
	Public Function CSSoutput(formattato)
		dim App, cr, cs
		cr = ""
		cs = ""
		if formattato then 
			cr = vbCrLf
			cs = "  "
		end if
		App = JSONoutput()		
		App = Replace(App, vbCrLf, "")
		App = Replace(App, """", "")
		App = Replace(App, "  ", "")
		App = Mid(App, 2, Len(App)-2)
		App = Replace(App, "},", cr & "}" & vbCrLf & cr)
		App = Replace(App, ": {", " {" & cr & cs)
		App = Replace(App, ",", "; " & cr & cs)
		if formattato then App = Left(App, Len(App)-1) & cr & "}"
		CSSoutput = App
	end Function



	Public Function Collection()
		set Collection = Server.CreateObject("Scripting.Dictionary")
	End Function

	Public Function AddToCollection(dictobj)
		dim newlabel
		if TypeName(dictobj) <> "Dictionary" then Err.Raise 1, "AddToCollection Error", "Not a collection."
		newlabel = dictobj.Count
		dictobj.Add newlabel, Collection()
		set AddToCollection = dictobj.item(newlabel)
	end function
	
	Public Function AddElement(element, overwrite)
		dim elemexists, subItem, test
		if (element <> "") then
			elemexists = data.Exists(element)
			'Response.Write ("<h3>Esiste "& element & " ? "& elemexists &"</h3>")
			if (overwrite) or (not elemexists) then
				Set data(element) = Collection()
			end if
		end if		
	end function
	
	Public Function AddItem(element, param, value)
		if (element <> "") and (param <> "") and (value <> "") then
			data.item(element).item(param) = value
		end if		
	end function
	
	Public Function GetItemValue(element, param)
		GetItemValue = ""
		if (element <> "") and (param <> "") then		
			if data.exists(element) then GetItemValue = data.item(element).item(param)
		end if		
	end function

	Private Function CleanUpJSONstring(originalstring)
		dim in_string, in_escape, i_tmp, char_tmp, s_tmp, line_tmp
		
		originalstring = Replace(originalstring,vbCrLf, "")

		p_datatype = Left(originalstring, 1) & Right(originalstring, 1)
		originalstring = Mid(originalstring, 2, Len(originalstring) - 2)
		in_string = False : in_escape = False
		For i_tmp = 1 To Len(originalstring)
			If in_escape Then
				in_escape = False
			Else
				char_tmp = Mid(originalstring, i_tmp, 1)
				Select Case char_tmp
					Case "\" : in_escape = True
					Case """" : s_tmp = s_tmp & char_tmp : in_string = Not in_string
					Case "{", "["
						s_tmp = s_tmp & char_tmp & InlineIf(in_string, "", vbCrLf)
					Case "}", "]"
						s_tmp = s_tmp & InlineIf(in_string, "", vbCrLf) & char_tmp
					Case "," : s_tmp = s_tmp & char_tmp & InlineIf(in_string, "", vbCrLf)
					Case Else : s_tmp = s_tmp & char_tmp
				End Select
			End If
		Next

		CleanUpJSONstring = ""
		s_tmp = split(s_tmp, vbCrLf)
		For Each line_tmp In s_tmp
			CleanUpJSONstring = CleanUpJSONstring & Trim(line_tmp) & vbCrLf
		Next
	End Function

	Private Function getJSONValue(ByVal val)
		val = Trim(val)
		If Left(val,1) = ":"  Then val = Mid(val, 2)
		If Right(val,1) = "," Then val = Left(val, Len(val) - 1)
		val = Trim(val)

		Select Case val
			Case "true"  : getJSONValue = True
			Case "false" : getJSONValue = False
			Case "null" : getJSONValue = Null
			Case Else
				If (Instr(val, """") = 0) Then
					If IsNumeric(val) Then
						getJSONValue = CDbl(val)
					Else
						getJSONValue = val
					End If
				Else
					If Left(val,1) = """" Then val = Mid(val, 2)
					If Right(val,1) = """" Then val = Left(val, Len(val) - 1)
					getJSONValue = JSONDecode(Trim(val))
				End If
		End Select
	End Function

	Private JSONoutput_level
	Public Function JSONoutput()
		JSONoutput_level = 1
		JSONoutput = Left(p_datatype, 1) & vbCrLf & GetDict(data) & Right(p_datatype, 1)
	End Function

	Private Function GetDict(objDict)
		dim item, dicttype, label, keyvals
		For Each item In objDict
			Select Case TypeName(objDict.Item(item))
				Case "Dictionary"
					GetDict = GetDict & Space(JSONoutput_level * 4)
					
					dicttype = "[]"
					For Each label In objDict.Item(item).Keys
						 If Not IsInt(label) Then dicttype = "{}"
					Next

					If IsInt(item) Then
						GetDict = GetDict & Left(dicttype,1) & vbCrLf
					Else
						GetDict = GetDict & """" & JSONEncode(item) & """" & ": " & Left(dicttype,1) & vbCrLf
					End If
					JSONoutput_level = JSONoutput_level + 1
					
					keyvals =  objDict.Keys
					GetDict = GetDict & GetSubDict(objDict.Item(item)) & Space(JSONoutput_level * 4) & Right(dicttype,1) & InlineIf(item = keyvals(objDict.Count - 1),"" , ",") & vbCrLf
				Case Else
					keyvals =  objDict.Keys
					GetDict = GetDict & Space(JSONoutput_level * 4) & InlineIf(IsInt(item), "", """" & JSONEncode(item) & """: ") & WriteValue(objDict.Item(item)) & InlineIf(item = keyvals(objDict.Count - 1),"" , ",") & vbCrLf
			End Select
		Next
	End Function

	Private Function IsInt(val)
		IsInt = (TypeName(val) = "Integer" Or TypeName(val) = "Long")
	End Function

	Private Function GetSubDict(objSubDict)
		GetSubDict = GetDict(objSubDict)
		JSONoutput_level= JSONoutput_level -1
	End Function

	Private Function WriteValue(ByVal val)
		Select Case TypeName(val)
			Case "Double", "Integer", "Long": WriteValue = val
			Case "Null"						: WriteValue = "null"
			Case "Boolean"					: WriteValue = InlineIf(val, "true", "false")
			Case Else		: WriteValue = """" & JSONEncode(val) & """"
		End Select
	End Function

	Private Function JSONEncode(ByVal val)
		val = Replace(val, "\", "\\")
		val = Replace(val, """", "\""")
		'val = Replace(val, "/", "\/")
		val = Replace(val, Chr(8), "\b")
		val = Replace(val, Chr(12), "\f")
		val = Replace(val, Chr(10), "\n")
		val = Replace(val, Chr(13), "\r")
		val = Replace(val, Chr(9), "\t")
		JSONEncode = Trim(val)
	End Function

	Private Function JSONDecode(ByVal val)
		val = Replace(val, "\""", """")
		val = Replace(val, "\\", "\")
		val = Replace(val, "\/", "/")
		val = Replace(val, "\b", Chr(8))
		val = Replace(val, "\f", Chr(12))
		val = Replace(val, "\n", Chr(10))
		val = Replace(val, "\r", Chr(13))
		val = Replace(val, "\t", Chr(9))
		JSONDecode = Trim(val)
	End Function

	Private Function InlineIf(condition, returntrue, returnfalse)
		If condition Then InlineIf = returntrue Else InlineIf = returnfalse
	End Function

	Private Function Strip(ByVal val, stripper)
		If Left(val, 1) = stripper Then val = Mid(val, 2)
		If Right(val, 1) = stripper Then val = Left(val, Len(val) - 1)
		Strip = val
	End Function
End Class
%>
