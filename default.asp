<%@ Language=VBScript %>
<% Option Explicit %>
<!--#INCLUDE file="aspJSON.asp"-->

<html>

<head>
<title>CSS To JSON</title>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
<meta content="it" http-equiv="Content-Language">
<style type="text/css">
  blockquote {margin:0 0 0 20px}
	textarea {width:100%; height:300px;}
</style>

</head>

<body>
<%
dim CSSDef, CSS, JSON, NomeDef, Nome

function WriteJSON(Elem)
	if oJSON.data.exists(Elem) then
		Response.Write "<code>"
		For Each subItem In oJSON.data.item(Elem)
		    Response.Write subItem & ": " & oJSON.data.item(Elem).item(subItem) & "<br>"
		Next
		Response.Write "</code>"
	end if
end function


CSSDef = "table { "& VbCrLf & _
	 	 " border-collapse:collapse;"& VbCrLf & _
		 " border-spacing:0;"& VbCrLf & _
		 "}"& VbCrLf & _
		 "fieldset { "& VbCrLf & _
		 " border:0;"& VbCrLf & _
		 "}"& VbCrLf & _
		 "    "& VbCrLf & _
		 "div    {background-image: url('images/image.jpg'); background-position: left center; background-color: red}"& VbCrLf & _
		 "    "& VbCrLf & _
		 "input{"& VbCrLf & _
		 " border:1px solid #b0b0b0;"& VbCrLf & _
		 " padding: 2px 5px;"& VbCrLf & _
		 " color:#979797;"& VbCrLf & _
		 "}"& VbCrLf & VbCrLf
		 
NomeDef = "div"

CSS = Request("CSS")
Nome = Request("Nome")
if Trim(CSS) = "" then CSS = CSSDef
if Trim(Nome) = "" then Nome = NomeDef
		

%>

<h1>CSS To JSON</h1>
<h5>Made by <a href="http://www.sitiegrafica.it" target="_blank">Siti e Grafica</a></h5>

Questa utility permette di convertire un CSS in formato JSON e viceversa.<br/>
Viene utilizzata la libreria esterna (modificata) ASPJSON - <a href="http://www.aspjson.com/">http://www.aspjson.com/</a> 

<div style="clear:both; height:20px"></div>


<div style="float:left; width:500px">
	<h3>CSS iniziale</h3>
	<form action="<%=Request.ServerVariables("URL")%>" method="post">
		<textarea name="CSS"><%=CSS%></textarea><br/>
		Elemento: <input name="Nome" value="<%=Nome%>" placeholder="Elemento da analizzare" style="width:300px">
		<input type="submit" value="Analizza" />
	</form>
</div>
<% if CSS = CSSDef then %>
<div style="float:left; width:500px; margin-left:20px">
	<h3>JSON ASPETTATO</h3>
	<textarea>
{
  "table": {
    "border-collapse": "collapse",
    "border-spacing": "0"
  },
  "fieldset": {
    "border": "0"
  },
  "div": {
    "background-image": "url('images/image.jpg')",
    "background-position": "leftcenter",
    "background-color": "red"
  },
  "input": {
    "border": "1px solid #b0b0b0",
    "padding": "2px 5px",
    "color": "#979797"
  }
}
	</textarea>
</div>
<% end if %>

<%
	Dim oJSON, subItem
	Set oJSON = New aspJSON
	JSON = oJSON.convertCSS(CSS)   ' Converte il CSS in una stringa JSON e carica l'oggetto JSON
%>
<div style="clear:both; height:20px"></div>

<h3>JSON finale</h3>
<code>
<%=oJSON.FormattaJSON(JSON)%>
</code>
<div style="clear:both; height:20px"></div>


<h3>Lettura dell'elemento "<%=Nome%>"</h3>
<%=WriteJSON(Nome)%>
<div style="clear:both; height:20px"></div>


<h3>Lettura dell'item "background-image" di "<%=Nome%>"</h3>
<code>
"<%=oJSON.GetItemValue(Nome, "background-image")%>"
</code>
<div style="clear:both; height:20px"></div>


<h3>Lettura dell'item "non-esiste" di "<%=Nome%>"</h3>
<code>
"<%=oJSON.GetItemValue(Nome, "non-esiste")%>"
</code>
<div style="clear:both; height:20px"></div>


<h3>Lettura dell'item "non-esiste" di "non-esiste"</h3>
<code>
"<%=oJSON.GetItemValue("non-esiste", "non-esiste")%>"
</code>
<div style="clear:both; height:20px"></div>


<hr>



<%
	dim NameToAdd, Param1ToAdd, Param2ToAdd, Param3ToAdd, Value1ToAdd, Value2ToAdd, Value3ToAdd
	dim newitem
	NameToAdd = "a:hoover"
%>


<%
	Param1ToAdd = "border"
	Value1ToAdd = "1px #F00 solid"
	Param2ToAdd = "background-color"
	Value2ToAdd = "aqua"
	Param3ToAdd = "margin-top"
	Value3ToAdd = "20px"
%>
<h3>Aggiunta dell'elemento "<%=NameToAdd%>"</h3>
<%	
	oJSON.AddElement NameToAdd, true
	' Aggiunge gli item
	oJSON.AddItem NameToAdd, Param1ToAdd, Value1ToAdd
	oJSON.AddItem NameToAdd, Param2ToAdd, Value2ToAdd
	oJSON.AddItem NameToAdd, Param3ToAdd, Value3ToAdd

		
%>
<code>
<%=oJSON.FormattaJSON(oJSON.JSONoutput())%>
</code>
<div style="clear:both; height:20px"></div>




<hr>





<%
	Param1ToAdd = "border"
	Value1ToAdd = "0"
	Param2ToAdd = "font-size"
	Value2ToAdd = "10px"
	Param3ToAdd = "margin-bottom"
	Value3ToAdd = "none"
%>
<h3>Reset totale dell'elemento "<%=NameToAdd%>"</h3>
<%
	
	oJSON.AddElement NameToAdd, true   ' Crea/Resetta(true) o Modifica(false) l'elemento	
	' Aggiunge gli item
	oJSON.AddItem NameToAdd, Param1ToAdd, Value1ToAdd
	oJSON.AddItem NameToAdd, Param2ToAdd, Value2ToAdd
	oJSON.AddItem NameToAdd, Param3ToAdd, Value3ToAdd
	
	
%>
<code>
<%=oJSON.FormattaJSON(oJSON.JSONoutput())%>
</code>
<div style="clear:both; height:20px"></div>




<hr>





<%
	Param1ToAdd = "border"
	Value1ToAdd = ""
	Param2ToAdd = "background-color"
	Value2ToAdd = "aqua"
	Param3ToAdd = "margin-top"
	Value3ToAdd = "20px"
%>
<h3>Update dell'elemento "<%=NameToAdd%>"</h3>
<%
	
	oJSON.AddElement NameToAdd, false   ' Crea/Resetta(true) o Modifica(false) l'elemento	
	' Aggiunge gli item
	oJSON.AddItem NameToAdd, Param1ToAdd, Value1ToAdd
	oJSON.AddItem NameToAdd, Param2ToAdd, Value2ToAdd
	oJSON.AddItem NameToAdd, Param3ToAdd, Value3ToAdd
	
	
%>
<code>
<%=oJSON.FormattaJSON(oJSON.JSONoutput())%>
</code>
<div style="clear:both; height:20px"></div>




<hr>




<h3>Lettura dell'elemento "<%=NameToAdd%>"</h3>
<%=WriteJSON(NameToAdd)%>




<hr>




<div style="float:left; width:500px">
	<h3>CSS formattato</h3>
	<textarea><%=oJSON.CSSoutput(true)%></textarea>
</div>

<div style="float:left; width:500px; margin-left:20px">
	<h3>CSS non formattato</h3>
	<textarea><%=oJSON.CSSoutput(false)%></textarea>
</div>
<div style="clear:both; height:20px"></div>



<h3>Referenze</h3>
<ul>
	<li><a href="http://www.aspjson.com/">http://www.aspjson.com/</a>: Leggere e scrivere JSON</li>
	<li><a href="https://code.google.com/p/aspjson/">https://code.google.com/p/aspjson/</a></li>
	<li><a href="http://jsonlint.com/">http://jsonlint.com/</a>: Validatore JSON</li>
</ul>

<% Set oJSON = nothing %>

</body>
</html>
