ASP CSSToJSON 
=============

This utility Convert CSS to JSON and vice-versa
It is based on ASPJSON - http://www.aspjson.com/
Thanks. http://bit.ly/sitiegrafica

USAGE:
-----------
See default.asp demo.

' Create object JSON

Dim oJSON

Set oJSON = New aspJSON



' Convert CSS in JSON (also loads the oJSON object)

JSON = oJSON.convertCSS(CSS) 



' Read the item value

myvar = oJSON.GetItemValue("div", "background-image")



' Add element with overwrite (will delete all the items)

NameToAdd = "h1"

oJSON.AddElement NameToAdd, true



' Add/update element without overwrite it if exists (without deleting the items)

NameToAdd = "h1"

oJSON.AddElement NameToAdd, false



' Add item to element

Param1ToAdd = "border-bottom"

Value1ToAdd = "1px #F00 solid"

oJSON.AddItem NameToAdd, Param1ToAdd, Value1ToAdd



' Exports JSON to CSS (with format)

oJSON.CSSoutput(true)



' Exports JSON to CSS (without format)

oJSON.CSSoutput(false)
