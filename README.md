# KLH Revit Developer Tools
[![Build status](https://dev.azure.com/klhengineers/KLH/_apis/build/status/DeveloperTools-CI)](https://dev.azure.com/klhengineers/KLH/_build/latest?definitionId=18)
## KLHSnoop

Look at the object properties of elements. 

### Usage
Create a wrapping plugin and call the form like
```vb
Dim form As New DeveloperTools.KLHSnoop(commandData.Application.ActiveUIDocument)
form.Show()
```
or
```cs
var form = new DeveloperTools.KLHSnoop(commandData.Application.ActiveUIDocument);
form.Show();
```

### Features
![Screenshot](Resources/screenshot.png)

* **Property Grouping**: Groups properties by thier derived class.
* **Selection Parity**: If the UIDocument has elements in its selection when the tool is run, those elements are shown in the tool. Select elements in the tool to select them in the document. 
* **Lazy Loading**: Double click an object in the form to load its data.
* **Copy Value**: Right click an object to copy its value as shown, or all of its value as a json object.
* **Add Elements By Ids**: Right click an empty part of the form to add elements to the tree using their comma separated ElementIds
* **Zoom**: Ctrl+Scroll to zoom in and out.
* **Type Highlighting**: <span style="color:blue">Structs</span>, <span style="color:cyan">classes</span>, <span style="color:green">elements</span>, <span style="color:darkkhaki">enums</span>, and even <span style="color:red">exceptions</span> are highlighted differently.
* **Send to new window**: Right click any object in the tree to open it in a new modeless KLHSnoop window.
* **Pass in arguments**: Properties with ArgumentExceptions can be given an argument. Right click and Edit Parameters to pass in primitive types, ElementIds, Elements, or just a default object from a class' default constructor.
