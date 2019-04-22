' Copyright � 2019 Kohrs Lonnemann Heil Engineers, PSC (KLH Engineers)
' 1538 Alexandria Pike, Suite 11, Fort Thomas Kentucky, 41075
' www.klhengrs.com
'
' This software the proprietary and confidential property of KLH Engineers.
' 
' No party may possess or use this source code in any form for any reason
' without the express written consent of KLH Engineers.

Imports System.Windows
Imports System.Windows.Input
Imports Autodesk.Revit.DB

Public Class ElementSelector
    Public Property ElementIds As IEnumerable(Of ElementId)
    Public Property Filters As IList(Of ElementFilter)
    Public Property InputString As String
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        DataContext = Me
        ElementIds = {}
    End Sub

    Private Sub TextBox_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        Dim i = 0
        e.Handled = Not e.Text.Split(","c).All(Function(s) Integer.TryParse(s, i))
    End Sub

    Private Sub AddElements(sender As Object, e As RoutedEventArgs)
        For Each eid In InputString.Split(","c).Select(Function(s) New ElementId(Integer.Parse(s)))
            ElementIds = ElementIds.Append(eid)
        Next

        DialogResult = True
        Close()
    End Sub
End Class
