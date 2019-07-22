Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Media
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.DB.ExtensibleStorage
Imports Autodesk.Revit.UI
Imports Newtonsoft

''' <summary>
'''     The KLHSnoop Window object.
''' </summary>
Public Class KLHSnoop
    Implements INotifyPropertyChanged

    ''' <summary>
    '''     Creates a new KLHSnoop window populated with the elements in the selection
    '''     of the UIDocument
    ''' </summary>
    Public Sub New(uiDoc As UIDocument)
        DataContext = Me
        InitializeComponent()

        Dim eIds = uiDoc.Selection.GetElementIds

        Roots = New ObservableCollection(Of TreeMember)(
            eIds.Select(Function(eid) New TreeMember(
                uiDoc.Document.GetElement(eid).Name,
                uiDoc.Document.GetElement(eid),
                Me,
                True)))
        Me.UIDoc = uiDoc
    End Sub

    ''' <summary>
    '''     Creates a new KLHSnoop window populated with a given root TreeMember nodes.
    ''' </summary>
    Public Sub New(uiDoc As UIDocument, ParamArray rootNodes As TreeMember())
        DataContext = Me
        InitializeComponent()

        Roots = New ObservableCollection(Of TreeMember)(rootNodes)
        For Each root In Roots
            root.AlwaysShow = True
        Next
        Me.UIDoc = uiDoc
    End Sub

    ''' <summary>
    '''     The UIDocument that the KLHSnoop interacts with for selection and population.
    ''' </summary>
    Public Property UIDoc As UIDocument

    ''' <summary>
    '''     The root nodes of the object display
    ''' </summary>
    Public Property Roots As ObservableCollection(Of TreeMember)
        Get
            Return _roots
        End Get
        Set
            _roots = Value
            KLHOnPropertyChanged()
        End Set
    End Property

    Private _roots As New ObservableCollection(Of TreeMember)

    Private _loadProgress As Integer = 0

    Public Property LoadProgress As Integer
        Get
            Return _loadProgress
        End Get
        Set
            _loadProgress = Value
            KLHOnPropertyChanged()
        End Set
    End Property

    Public Property SearchString As String
        Get
            Return _searchString
        End Get
        Set(value As String)
            _searchString = value
            KLHOnPropertyChanged()
        End Set
    End Property
    Private _searchString As String

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

    Private Sub TextBox_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Key <> Key.Enter Then Exit Sub
    End Sub

    Private Sub RefreshData(sender As Object, e As RoutedEventArgs)
        If TypeOf CType(sender, Controls.Control).DataContext IsNot TreeGroup Then CType(CType(sender, Controls.Control).DataContext, TreeMember).LoadMembers(LoadProgress)
    End Sub

    Private Sub TreeView_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object))
        If TypeOf sender IsNot TreeMember Then Exit Sub
        Dim el As Element = CType(DirectCast(sender, TreeView).SelectedItem.Value, Element)
        If TypeOf el Is Element Then
            UIDoc.Selection.SetElementIds({el.Id})
        End If
    End Sub

    Private Sub NewWindow(sender As Object, e As RoutedEventArgs)
        Dim w = New KLHSnoop(UIDoc, DirectCast(sender.DataContext, TreeMember))
        w.Show()
    End Sub

    Private Sub StackPanel_MouseDown(sender As Object, e As MouseButtonEventArgs)
        If e.ClickCount = 2 AndAlso sender.DataContext IsNot Nothing AndAlso TypeOf sender.DataContext IsNot TreeGroup Then
            sender.DataContext.LoadMembers(LoadProgress)
            e.Handled = True
        End If
    End Sub

    Private Sub AddElements(sender As Object, e As RoutedEventArgs)
        Dim ew As New ElementSelector
        Dim result = ew.ShowDialog
        If result IsNot Nothing AndAlso result Then
            For Each eid In ew.ElementIds
                Roots.Add(New TreeMember(UIDoc.Document.GetElement(eid).Name, UIDoc.Document.GetElement(eid), Me, True))
            Next
        End If
    End Sub

    Private Sub EditParameters(sender As Object, e As RoutedEventArgs)
        Dim typedSender = DirectCast(sender.DataContext, TreeMember)
        If typedSender.Parameters IsNot Nothing AndAlso typedSender.Parameters.Any Then
            Dim pe = New ParameterEditor(typedSender.Parameters, UIDoc.Document)
            Dim result = pe.ShowDialog
            If result IsNot Nothing AndAlso result Then
                typedSender.Value = typedSender.Evaluate(pe.ParameterValues.Select(Function(t) t.Result).ToArray)
            End If
        End If
    End Sub

    Private Sub CopyValue(sender As Object, e As RoutedEventArgs)
        Dim typedSender = DirectCast(sender.DataContext, TreeMember)
        If typedSender.Parameters IsNot Nothing Then
            My.Computer.Clipboard.SetText(typedSender.ValueString)
        End If
    End Sub

    Private Sub CopyJSON(sender As Object, e As RoutedEventArgs)
        Dim typedSender = DirectCast(sender.DataContext, TreeMember)
        If typedSender.Parameters IsNot Nothing Then
            My.Computer.Clipboard.SetText(Json.JsonConvert.SerializeObject(typedSender.Value))
        End If
    End Sub

    Private Sub AddDocument(sender As Object, e As RoutedEventArgs)
        Roots.Add(New TreeMember("Document", UIDoc.Document, Me, True))
    End Sub

    Private Sub AddApplication(sender As Object, e As RoutedEventArgs)
        Roots.Add(New TreeMember("Application", UIDoc.Application, Me, True))
    End Sub

    Private Sub AddActiveView(sender As Object, e As RoutedEventArgs)
        Roots.Add(New TreeMember("Application", UIDoc.Document.ActiveView, Me, True))
    End Sub

    Private Sub SnoopWindow_MouseWheel(sender As Object, e As MouseWheelEventArgs)
        If My.Computer.Keyboard.CtrlKeyDown Then
            Me.FontSize = Math.Max(Me.FontSize + (e.Delta / 120), 1)
            e.Handled = True
        End If
    End Sub
End Class

Public Interface ITreeItem

    ''' <summary>
    '''     The displayed name of the node.
    ''' </summary>
    Property Name As String

    ''' <summary>
    '''     The displayed members of the node.
    ''' </summary>
    Property Members As ObservableCollection(Of ITreeItem)
    ReadOnly Property IconPath As String
    Property AlwaysShow As Boolean
End Interface

''' <summary>
'''     A node on the KLHSnoop browser containing information about an Object.
''' </summary>
Public Class TreeMember : Implements ITreeItem
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

    Private ReadOnly _parent As Object = Nothing
    Private ReadOnly _type As MemberTypes = 0

    Public Property AlwaysShow As Boolean Implements ITreeItem.AlwaysShow

    Public Property Name As String Implements ITreeItem.Name

    ''' <summary>
    '''     The displayed members of the node, grouped into TreeGroup objects.
    ''' </summary>
    Public Property Members As ObservableCollection(Of ITreeItem) Implements ITreeItem.Members
        Get
            Return _members
        End Get
        Set
            _members = Value
            KLHOnPropertyChanged()
        End Set
    End Property

    Private _members As ObservableCollection(Of ITreeItem)

    ''' <summary>
    '''     The name of the type of the value represented by this TreeNode
    ''' </summary>
    Public ReadOnly Property TypeName As String
        Get
            Return If(Value?.GetType.Name, "Nothing")
        End Get
    End Property

    ''' <summary>
    '''     The Object represented by this TreeNode
    ''' </summary>
    Public Property Value As Object
        Get
            Return _value
        End Get
        Set
            _value = Value
            KLHOnPropertyChanged()
            KLHOnPropertyChanged(NameOf(ValueString))
            KLHOnPropertyChanged(NameOf(TypeName))
        End Set
    End Property
    Private _value As Object

    ''' <summary>
    '''     A formatted String displayed to represent the value of the Object represented
    '''     by this TreeNode
    ''' </summary>
    Public ReadOnly Property ValueString As String
        Get
            If Value Is Nothing Then Return Nothing

            Dim vs = If(Value.ToString.Length > 50, Value.ToString.Substring(0, 50) & "...", Value.ToString)
            If Value.GetType.IsEnum Then
                vs = vs & " (" & Convert.ToInt32(Value) & ")"
            End If
            Return vs
        End Get
    End Property

    ''' <summary>
    '''     The definitions of what parameters are passed to the function called to
    '''     retrieve the value of the node.
    ''' </summary>
    Public Property Parameters As IEnumerable(Of ParameterInfo)
    Public ReadOnly Property AnyParameters As Boolean
        Get
            Return If(Parameters?.Any(), False)
        End Get
    End Property


    ''' <summary>
    '''     The path to which icon represents the node depending on what type of member
    '''     the node represents in its parent.
    ''' </summary>
    Public ReadOnly Property IconPath As String Implements ITreeItem.IconPath
        Get
            Select Case _type
                Case MemberTypes.Property
                    Return "/KLH.Revit.DeveloperTools;component/Resources/property.png"
                Case MemberTypes.Method
                    Return "/KLH.Revit.DeveloperTools;component/Resources/method.png"
                Case Else
                    Return "/KLH.Revit.DeveloperTools;component/Resources/collection.png"
            End Select
        End Get
    End Property

    ''' <summary>
    '''     The window that this node belongs to
    ''' </summary>
    Public Property Window As KLHSnoop

    ''' <summary>
    '''     Create a new TreeNode with a predefined name and value.
    ''' </summary>
    Public Sub New(name As String, value As Object, ByRef window As KLHSnoop, alwaysShow As Boolean)
        Me.Name = name
        Me.Value = value
        Me.Window = window
        Me.AlwaysShow = alwaysShow
    End Sub

    ''' <summary>
    '''     Create a new TreeNode using a parent and a Member. The name will be the
    '''     node's name as a member of the parent and the value is retrieved
    '''     dynamically.
    ''' </summary>
    Public Sub New(m As MemberInfo, parent As Object, ByRef window As KLHSnoop)

        _parent = parent
        _type = m.MemberType
        AlwaysShow = False
        Name = m.Name
        Me.Window = window
        If TypeOf m Is PropertyInfo Then
            Parameters = If(DirectCast(m, PropertyInfo).GetMethod?.GetParameters, {})
        ElseIf TypeOf m Is MethodInfo Then
            Parameters = DirectCast(m, MethodInfo).GetParameters
        End If
        If Parameters IsNot Nothing AndAlso Parameters.Any Then
            Name = Name & "(" & String.Join(", ", Parameters.Select(Function(pi) pi.ParameterType.Name).ToArray) & ")"
        End If
        Value = Evaluate({})
    End Sub

    ''' <summary>
    '''     Retrieve the value of a TreeNode that was constructed using a parent and
    '''     MemberInfo.
    ''' </summary>
    ''' <param name="params">Parameters passed to the invocation</param>
    Public Function Evaluate(params As Object()) As Object
        Try
            If params.Count = Parameters.Count Then
                Return _parent.GetType.InvokeMember(
                    New String(Name.TakeWhile(Function(c) c <> "(").ToArray()),
                    BindingFlags.GetProperty Or BindingFlags.InvokeMethod,
                    Nothing,
                    _parent,
                    params)
            Else
                Return New ArgumentException()
            End If
        Catch e As TargetInvocationException
            Return e.InnerException
        Catch e As Exception
            Return e
        End Try
    End Function

    ''' <summary>
    '''     Creates a child node for each member of the Object that this TreeNode
    '''     represents.
    ''' </summary>
    Public Sub LoadMembers(ByRef progress As Integer)

        Dim loaded = 0

        If Value Is Nothing Then Exit Sub

        Dim rawMembers = Value.GetType _
                .GetMembers(BindingFlags.Public _
                            Or BindingFlags.NonPublic _
                            Or BindingFlags.Instance) _
                .Where(Function(m) (TypeOf m Is PropertyInfo) _
                                   Or (TypeOf m Is MethodInfo _
                                       AndAlso Not DirectCast(m, MethodInfo).ReturnType Is GetType(Void)))

        Dim noDuplicates = rawMembers _
                .Where(Function(m) Not rawMembers _
                                       .Any(
                                           Function(p) _
                                               TypeOf p Is PropertyInfo AndAlso
                                               DirectCast(p, PropertyInfo).GetMethod Is m))

        Dim loadTotal = noDuplicates.Count

        Dim groupedRawMembers = noDuplicates.GroupBy(Function(m) m.DeclaringType)

        Members = New ObservableCollection(Of ITreeItem)(
            groupedRawMembers.Select(Function(g) New TreeGroup(
                       g.Key.Name,
                       New ObservableCollection(Of ITreeItem)(g.Select(Function(mi)
                                                                           loaded = loaded + 100
                                                                           Window.LoadProgress = Math.Floor(loaded / loadTotal)
                                                                           Return New TreeMember(mi, Value, Window)
                                                                       End Function) _
                                                                  .OrderBy(Function(tm) tm.Name) _
                                                                  .ToList)
                       )
                 )
            )

        Dim oc As New ObservableCollection(Of ITreeItem)
        If TypeOf Value Is ParameterSet Or TypeOf Value Is ParameterMap Then
            For Each p As Parameter In Value
                oc.Add(New TreeMember(p.Definition.Name, p, Window, False))
            Next
            oc = New ObservableCollection(Of ITreeItem)(oc.OrderBy(Function(ti) ti.Name))
        ElseIf TypeOf Value Is IDictionary Then
            For Each key In DirectCast(Value, IDictionary).Keys
                oc.Add(New TreeMember(key.ToString, Value(key), Window, False))
            Next
        ElseIf TypeOf Value Is IEnumerable Then
            Dim i = 0
            For Each o In Value
                oc.Add(New TreeMember("Results(" & i & ")", o, Window, False))
                i = i + 1
            Next
        End If
        If oc.Any Then Members.Add(New TreeGroup("Results", oc))
        Dim v = TryCast(Value, Element)
        If v IsNot Nothing Then
            Dim schemas As New ObservableCollection(Of ITreeItem)
            Dim slist = v.GetEntitySchemaGuids.Select(Function(guid) Schema.ListSchemas().Single(Function(sch) sch.GUID.Equals(guid)))
            For Each s In slist
                Dim entity = v.GetEntity(s)
                Dim fields = s.ListFields
                Dim schemaValues As New ObservableCollection(Of ITreeItem)
                For Each field In fields
                    Dim method As MethodInfo = entity.GetType.GetMethod("Get",
                                                                        BindingFlags.Public Or BindingFlags.Instance,
                                                                        Nothing,
                                                                        {GetType(String)},
                                                                        Nothing)
                    Select Case field.ContainerType
                        Case ContainerType.Simple
                            method = method.MakeGenericMethod(field.ValueType)
                        Case ContainerType.Array
                            method = method.MakeGenericMethod(GetType(IList(Of )).MakeGenericType(field.ValueType))
                        Case ContainerType.Map
                            method = method.MakeGenericMethod(GetType(IDictionary(Of ,)).MakeGenericType(field.KeyType, field.ValueType))
                    End Select
                    Dim val = method.Invoke(entity, {field.FieldName})
                    schemaValues.Add(New TreeMember(field.FieldName, val, Window, False))
                Next
                schemas.Add(New TreeGroup(s.SchemaName, schemaValues))
            Next
            If schemas.Any Then Members.Add(New TreeGroup("Extensible Storage Schemas", schemas))
        End If
    End Sub
End Class

''' <summary>
'''     A labeled group of TreeNodes.
''' </summary>
Public Class TreeGroup : Implements ITreeItem
    Public Property Name As String Implements ITreeItem.Name
    Public Property Members As ObservableCollection(Of ITreeItem) Implements ITreeItem.Members
    Public Property AlwaysShow As Boolean Implements ITreeItem.AlwaysShow
        Get
            Return True
        End Get
        Set(value As Boolean)

        End Set
    End Property

    Public ReadOnly Property IconPath As String Implements ITreeItem.IconPath
        Get
            Return "/KLH.Revit.DeveloperTools;component/Resources/class.png"
        End Get
    End Property

    Public Sub New(n As String, m As ObservableCollection(Of ITreeItem))
        Name = n
        Members = m
    End Sub
End Class

Public Class ObjectToBrushConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object _
        Implements IValueConverter.Convert

        If value Is Nothing Then
            Return Brushes.DarkBlue
        ElseIf TypeOf value Is Exception Then
            Return Brushes.Red
        ElseIf value.GetType.IsEnum Then
            Return Brushes.DarkKhaki
        ElseIf TypeOf value Is Element Then
            Return Brushes.DarkGreen
        ElseIf value.GetType.IsClass Then
            Return Brushes.Cyan
        Else
            Return Brushes.DarkBlue
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) _
        As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class

Public Class ObjectToBrushLabelConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object _
        Implements IValueConverter.Convert
        If value Is Nothing Then
            Return "Nothing"
        ElseIf TypeOf value Is Exception Then
            Return "Exception"
        ElseIf value.GetType.IsEnum Then
            Return "Enum"
        ElseIf TypeOf value Is Element Then
            Return "Element, select to select in Revit"
        ElseIf value.GetType.IsClass Then
            Return "Class"
        Else
            Return "Value Type"
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) _
        As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class

Public Class BooleanToVisibilityConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        If value Then
            Return Windows.Visibility.Visible
        Else
            Return Windows.Visibility.Collapsed
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class

Public Class SearchConverter : Implements IMultiValueConverter

    Public Function Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
        If values.Length <> 3 Then Return False
        Dim valueStrings = values.Select(Function(v) If(v IsNot Nothing, DirectCast(v, String).ToUpper, Nothing))
        Return values(2) _
                OrElse valueStrings(0) Is Nothing _
                OrElse valueStrings(0) = String.Empty _
                OrElse valueStrings(1).Contains(valueStrings(0).ToUpper)
    End Function

    Public Function ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class