Imports System

Imports Autodesk.Revit.UI
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.Attributes
<Transaction(TransactionMode.Manual)>
Public Class DeveloperToolsCmd
    Implements IExternalCommand
    Public Function Execute(commandData As ExternalCommandData, ByRef message As String, elements As ElementSet) As Result Implements IExternalCommand.Execute
        Dim form As New DeveloperTools.KLHSnoop(commandData.Application.ActiveUIDocument)
        form.Show()
    End Function
End Class
