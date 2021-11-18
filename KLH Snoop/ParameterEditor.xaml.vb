Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Windows
Imports System.Windows.Media
Imports Autodesk.Revit.DB
Imports MS.Win32

Public Class ParameterEditor
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    ''' <summary>
    ''' Raises the PropertyChanged event.
    ''' </summary>
    ''' <param name="propertyName">Optional. Name of the property that has changed.</param>
    ''' <param name="newValue">Optional. New property value.</param>
    ''' <param name="oldValue">Optional. Old property value.</param>
    <System.Diagnostics.DebuggerStepThrough()>
    Protected Sub KLHOnPropertyChanged(<CallerMemberName> Optional ByVal propertyName As String = Nothing,
                                           Optional newValue As Object = Nothing,
                                           Optional oldValue As Object = Nothing)

        'Raise the property changed event.
        RaiseEvent PropertyChanged(sender:=Me, e:=(New PropertyChangedEventArgs(propertyName:=propertyName)))
    End Sub

    Private _parameterValues As New List(Of ParameterValue)
    Private ReadOnly _doc As Document
    Public Property ParameterValues As List(Of ParameterValue)
        Get
            Return _parameterValues
        End Get
        Set
            _parameterValues = Value
            KLHOnPropertyChanged()
        End Set
    End Property


    Public Sub New(params As IEnumerable(Of ParameterInfo), doc As Document)

        ' This call is required by the designer.
        InitializeComponent()


        ' Add any initialization after the InitializeComponent() call.
        _doc = doc
        ParameterValues = New List(Of ParameterValue)
        For Each parameterInfo As ParameterInfo In params
            ParameterValues.Add(New ParameterValue(parameterInfo, _doc))
        Next
        DataContext = Me
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        For Each p In ParameterValues
            If Not p.FromString(p.Result) Then
                Exit Sub
            End If
        Next
        DialogResult = True
        Close()
    End Sub
End Class

Public Class ParameterValue
    Implements INotifyPropertyChanged
    Public Property Result As Object

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    ''' <summary>
    ''' Raises the PropertyChanged event.
    ''' </summary>
    ''' <param name="propertyName">Optional. Name of the property that has changed.</param>
    ''' <param name="newValue">Optional. New property value.</param>
    ''' <param name="oldValue">Optional. Old property value.</param>
    <System.Diagnostics.DebuggerStepThrough()>
    Protected Sub KLHOnPropertyChanged(<CallerMemberName> Optional ByVal propertyName As String = Nothing,
                                           Optional newValue As Object = Nothing,
                                           Optional oldValue As Object = Nothing)

        'Raise the property changed event.
        RaiseEvent PropertyChanged(sender:=Me, e:=(New PropertyChangedEventArgs(propertyName:=propertyName)))
    End Sub

    Public Property Parameter As ParameterInfo
    Private _argumentString As String
    Private ReadOnly _doc As Document

    Public Sub New(p As ParameterInfo, doc As Document)
        _doc = doc
        Parameter = p
        ArgumentString = String.Empty
    End Sub

    Public Property ArgumentString As String
        Get
            Return _argumentString
        End Get
        Set(value As String)
            _argumentString = value
            'PropertyChanged("ArgumentString")
            KLHOnPropertyChanged(NameOf(ValidationColor))
        End Set
    End Property

    Public ReadOnly Property ValidationColor As SolidColorBrush
        Get
            If FromString(Nothing) Then
                Return Brushes.Green
            Else
                Return Brushes.Firebrick
            End If
        End Get
    End Property

    Public Property Editable As Boolean = True

    Public Function FromString(ByRef out As Object) As Boolean
        Select Case Parameter.ParameterType
            Case GetType(String)
                out = ArgumentString
                Return True
            Case GetType(Boolean)
                Return Boolean.TryParse(ArgumentString, out)
            Case GetType(Integer)
                Return Integer.TryParse(ArgumentString, out)
            Case GetType(Double)
                Return Double.TryParse(ArgumentString, out)
            Case GetType(Guid)
                Return Guid.TryParse(ArgumentString, out)
            Case GetType(ElementId)
                Dim int = 0
                If Integer.TryParse(ArgumentString, int) Then
                    out = New ElementId(int)
                    Return True
                Else
                    Return False
                End If
            Case Else
                If Parameter.ParameterType.IsEnum Then
                    Try
                        out = [Enum].Parse(Parameter.ParameterType, ArgumentString)
                        Return True
                    Catch e As ArgumentException
                        Return False
                    End Try
                ElseIf Parameter.ParameterType.IsSubclassOf(GetType(Element)) Then
                    Dim int = 0
                    If Integer.TryParse(ArgumentString, int) Then
                        out = _doc.GetElement(New ElementId(int))
                        Return Parameter.ParameterType.IsInstanceOfType(out)
                    Else
                        Return False
                    End If
                ElseIf Parameter.ParameterType.GetConstructor(New Type() {}) IsNot Nothing Then
                    out = Parameter.ParameterType.GetConstructor(New Type() {}).Invoke(New Object() {})
                    _argumentString = "{Default " & Parameter.ParameterType.Name & " Object}"
                    Return True
                Else
                    _argumentString = "{Unable to enter " & Parameter.ParameterType.Name & " parameter into KLHSnoop}"
                    Editable = False
                    Return False
                End If
        End Select
    End Function
End Class
