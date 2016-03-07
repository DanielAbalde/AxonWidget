Imports System.Collections.Generic
Imports System.Linq
Imports System.Xml.Linq
Imports System.Drawing
Imports System.Windows.Forms
Imports Grasshopper
Imports Grasshopper.Kernel
Imports Grasshopper.GUI
Imports Grasshopper.GUI.Canvas

Public Class SetupWidget
    Inherits Grasshopper.Kernel.GH_AssemblyPriority

    Public Overrides Function PriorityLoad() As GH_LoadingInstruction

        AddHandler Instances.CanvasCreated, AddressOf SetupAxon

        Return GH_LoadingInstruction.Proceed
    End Function

    Private Sub SetupAxon(canvas As GH_Canvas)

        RemoveHandler Instances.CanvasCreated, AddressOf SetupAxon

        For Each w As Grasshopper.GUI.Widgets.IGH_Widget In Instances.ActiveCanvas.Widgets
            If w.Name.Equals("Axon") Then Exit Sub
        Next
        Instances.ActiveCanvas.Widgets.Add(New Axon)
    End Sub

End Class

Public Class Axon
    Inherits Grasshopper.GUI.Widgets.GH_Widget

    Private Enabled As Boolean
    Private _IsActive As Boolean
    Private _Mode As AxonMode
    Private Elements As List(Of AxonElement)
    Private _Level As Integer
    Private _Selected As Integer
    Private ParamOfWire As IGH_Param
    Private TabSaved As GUI.Ribbon.GH_LayoutTab
    Private CenterControl As Drawing.PointF
    Private WithEvents WECanvas As GUI.Canvas.GH_Canvas
    Private XML As AxonXML
    Private _SizeIcon As Integer
    Private _RadiusInner As Integer
    Private _Max As Integer
    Private Radius As Integer
    Private Boundary As Integer
    Private sorted As Boolean

    Sub New()
        Enabled = Grasshopper.Instances.Settings.GetValue("Widget.Axon.Show", True)
        _Mode = Grasshopper.Instances.Settings.GetValue("Widget.Axon.Mode", 1)
        _SizeIcon = Grasshopper.Instances.Settings.GetValue("Widget.Axon.Attributes.SizeIcon", 24)
        _RadiusInner = Grasshopper.Instances.Settings.GetValue("Widget.Axon.Attributes.RadiusInner", 20)
        _Max = Grasshopper.Instances.Settings.GetValue("Widget.Axon.Attributes.Maximum", 43)
        Radius = _RadiusInner + _SizeIcon
        Boundary = RadiusInner + _SizeIcon * 1.6
        If _Mode = AxonMode.None Then _Mode = AxonMode.ByCategory
        Owner = Grasshopper.Instances.ActiveCanvas
        WECanvas = Grasshopper.Instances.ActiveCanvas

        Elements = New List(Of AxonElement)
        _Level = 0
        _Selected = -1
        ParamOfWire = Nothing
        CenterControl = Nothing
        _IsActive = False
    End Sub

#Region "Mouse"
    Public Overrides Function RespondToMouseDown(sender As GH_Canvas, e As GH_CanvasMouseEvent) As GH_ObjectResponse
        If Not Enabled Then Return GH_ObjectResponse.Ignore
        If sender.Document Is Nothing Then Return GH_ObjectResponse.Ignore

        If Not IsActive Then
            If (e.Button = MouseButtons.Right) Then
                If (IsWireInteraction()) Then
                    CenterControl = Me.ConstrainCursor(e.ControlLocation, 80)
                    Me.Owner.ActiveInteraction.Destroy()
                    IsActive = True
                    Return GH_ObjectResponse.Capture
                Else
                    If (Control.ModifierKeys = Keys.Alt) Then
                        CenterControl = Me.ConstrainCursor(e.ControlLocation, 80)
                        IsActive = True
                        Return GH_ObjectResponse.Capture
                    Else
                        Return GH_ObjectResponse.Ignore
                    End If
                End If
            Else
                Return GH_ObjectResponse.Ignore
            End If
        Else
            If (InDisk(e.ControlLocation)) Then
                If (InCircle(e.ControlLocation)) Then
                    If (e.Button = MouseButtons.Right) Then Return GH_ObjectResponse.Handled
                    If (e.Button = MouseButtons.Left) AndAlso Level > 0 Then
                        Dim loc As Drawing.PointF = CenterControl
                        Dim param As IGH_Param = ParamOfWire
                        IsActive = False
                        CenterControl = loc
                        ParamOfWire = param
                        IsActive = True
                        Return GH_ObjectResponse.Capture
                    End If
                    Return GH_ObjectResponse.Ignore
                Else
                    If Selected <> -1 Then
                        Level += 1
                    Else
                        Return GH_ObjectResponse.Ignore
                    End If
                    Select Case Mode
                        Case AxonMode.ByCategory
                            Select Case Level
                                Case 0
                                    SetElements()
                                Case 1
                                    SetElements()
                                Case 2
                                    If ParamOfWire Is Nothing Then
                                        InstantiateObject()
                                    Else
                                        InstantiateObject(ParamOfWire)
                                    End If
                                    IsActive = False
                                    Return GH_ObjectResponse.Release
                                Case Else
                                    IsActive = False
                                    Return GH_ObjectResponse.Release
                            End Select

                        Case Else
                            If ParamOfWire Is Nothing Then
                                InstantiateObject()
                            Else
                                InstantiateObject(ParamOfWire)
                            End If
                            IsActive = False
                            Return GH_ObjectResponse.Release
                    End Select
                    Return GH_ObjectResponse.Ignore
                End If

            Else
                If (e.Button = MouseButtons.Right) Then
                    IsActive = False
                    CenterControl = Me.ConstrainCursor(e.ControlLocation, 80)
                    IsActive = True
                    Return GH_ObjectResponse.Capture
                Else
                    IsActive = False
                    Return GH_ObjectResponse.Release
                End If
            End If
        End If

    End Function

    Public Overrides Function RespondToMouseUp(sender As GH_Canvas, e As GH_CanvasMouseEvent) As GH_ObjectResponse
        If sender.Document Is Nothing Then Return GH_ObjectResponse.Ignore
        If InDisk(e.ControlLocation) Then
            If InCircle(e.ControlLocation) Then
                Return MyBase.RespondToMouseUp(sender, e)
            Else
                Return GH_ObjectResponse.Ignore
            End If
        Else
            Return GH_ObjectResponse.Release
        End If
    End Function

    Public Overrides Function RespondToMouseMove(sender As GH_Canvas, e As GH_CanvasMouseEvent) As GH_ObjectResponse
        If sender.Document Is Nothing Then Return GH_ObjectResponse.Ignore
        If IsActive Then
            OnButton(e.ControlLocation)
            ' Rhino.RhinoApp.WriteLine(String.Format("IsActive: {0}, Selected: {1}, Elements.Count: {2}, Mode: {3}, Level: {4}, Boundary: {5}", IsActive, Selected, Elements.Count, Mode.ToString, Level, Boundary))

            Return MyBase.RespondToMouseMove(sender, e)
        Else
            Return GH_ObjectResponse.Ignore
        End If
    End Function

    Public Overrides Function RespondToKeyDown(sender As GH_Canvas, e As System.Windows.Forms.KeyEventArgs) As GH_ObjectResponse
        If e.KeyCode = System.Windows.Forms.Keys.Escape OrElse e.KeyCode = System.Windows.Forms.Keys.Cancel Then
            Instances.CursorServer.ResetCursor(sender)
            IsActive = False
            Return GH_ObjectResponse.Release
        End If
        Return GH_ObjectResponse.Ignore
    End Function

    Private Sub OnMouseDown(sender As Object, e As MouseEventArgs) Handles WECanvas.MouseDown
        If Not Enabled Then Exit Sub
        If sender.Document Is Nothing Then Exit Sub
        If (e.Button = MouseButtons.Right) Then
            If (IsWireInteraction()) Then
                CenterControl = Me.ConstrainCursor(Owner.CursorControlPosition, 80)
                Me.Owner.ActiveInteraction = Nothing

                IsActive = True
            End If
        Else
            Exit Sub
        End If
    End Sub

    Private Sub DocumentChanged(sender As GH_Canvas, e As GH_CanvasDocumentChangedEventArgs) Handles WECanvas.DocumentChanged
        If Enabled AndAlso e.NewDocument IsNot Nothing Then
            XML = New AxonXML(e.NewDocument)
            If Not sorted Then
                XML.XSortComponentsByFrecuency(True)
                sorted = True
            End If
        End If
    End Sub
#End Region

#Region "Property"
    Public Overrides ReadOnly Property Description As String
        Get
            Return "Electric Avenue"
        End Get
    End Property

    Public Overrides ReadOnly Property Icon_24x24 As Bitmap
        Get
            Return My.Resources.AxonIcon4
        End Get
    End Property

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Axon"
        End Get
    End Property

    Public Overrides Property Visible As Boolean
        Get
            Return Enabled
        End Get
        Set(value As Boolean)
            Enabled = value
        End Set
    End Property

    Public Property Mode As AxonMode
        Get
            Return _Mode
        End Get
        Set(value As AxonMode)
            If (value = _Mode) Then Exit Property
            Dim loc As Drawing.PointF = CenterControl
            Dim param As IGH_Param = ParamOfWire
            IsActive = False
            CenterControl = loc
            ParamOfWire = param
            IsActive = True
            Owner.Refresh()
            _Mode = value
            Grasshopper.Instances.Settings.SetValue("Widget.Axon.Mode", CInt(value))
            SetElements()
        End Set
    End Property

    Public Property IsActive As Boolean
        Get
            Return _IsActive
        End Get
        Set(value As Boolean)
            _IsActive = value
            If value Then
                Owner.ActiveWidget = Me
                SetElements()
            Else
                ClearAll()
                Me.Owner.Invalidate()
            End If
        End Set
    End Property

    Private Property Level As Integer
        Get
            Return _Level
        End Get
        Set(value As Integer)
            _Level = value
            If value = 0 Then Boundary = Radius * 1.5
        End Set
    End Property

    Public Property Selected As Integer
        Get
            Return _Selected
        End Get
        Set(value As Integer)
            _Selected = value
            Me.Owner.Invalidate()
        End Set
    End Property

    'Public ReadOnly Property ConnectivityDatabase As AxonConnectivityDatabase
    ' Get
    'Return _ConnectivityDatabase
    'End Get
    ' End Property

    Public ReadOnly Property AxonXML As AxonXML
        Get
            Return XML
        End Get
    End Property

    Public Property SizeIcon As Integer
        Get
            Return _SizeIcon
        End Get
        Set(value As Integer)
            _SizeIcon = value
            Radius = RadiusInner + _SizeIcon
            Grasshopper.Instances.Settings.SetValue("Widget.Axon.Attributes.SizeIcon", value)
            SetElements()
        End Set
    End Property

    Public Property RadiusInner As Integer
        Get
            Return _RadiusInner
        End Get
        Set(value As Integer)
            _RadiusInner = value
            Radius = RadiusInner + _SizeIcon
            Grasshopper.Instances.Settings.SetValue("Widget.Axon.Attributes.RadiusInner", value)
            SetElements()
        End Set
    End Property

    Public Property Maximun As Integer
        Get
            Return _Max
        End Get
        Set(value As Integer)
            _Max = value
            Grasshopper.Instances.Settings.SetValue("Widget.Axon.Attributes.Maximum", value)
            SetElements()
        End Set
    End Property
#End Region

#Region "Function"
    Private Function IsWireInteraction() As Boolean
        Dim it As Grasshopper.GUI.Canvas.Interaction.IGH_MouseInteraction = Me.Owner.ActiveInteraction
        If (it IsNot Nothing) AndAlso (TypeOf (it) Is Grasshopper.GUI.Canvas.Interaction.GH_WireInteraction) Then

            Dim wi As Grasshopper.GUI.Canvas.Interaction.GH_WireInteraction = DirectCast(it, Grasshopper.GUI.Canvas.Interaction.GH_WireInteraction)

            Dim att As IGH_Attributes = Me.Owner.Document.FindAttributeByGrip(wi.CanvasPointDown, False, 20)
            If (att IsNot Nothing) AndAlso (TypeOf (att.DocObject) Is IGH_Param) Then
                ParamOfWire = DirectCast(att.DocObject, IGH_Param)

                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If

    End Function

    Public Overrides Function Contains(pt_control As Drawing.Point, pt_canvas As PointF) As Boolean
        Return InDisk(pt_control)
    End Function

    Public Function OnButton(Location As Drawing.Point) As Boolean
        If Not IsActive Then Return False
        If Elements.Count = 0 Then Return False

        Dim i As New Integer
        While True
            If i >= Elements.Count Then
                Selected = -1
                Exit While
            End If
            If Elements(i).Contains(Location) Then
                Selected = i
                Exit While
            End If
            i += 1
        End While

        Return Selected <> -1
    End Function

    Public Function InDisk(Location As Drawing.PointF) As Boolean
        Return Grasshopper.GUI.GH_GraphicsUtil.Distance(CenterControl, Location) <= Boundary
    End Function

    Public Function InCircle(Location As Drawing.PointF) As Boolean
        Return Grasshopper.GUI.GH_GraphicsUtil.Distance(CenterControl, Location) <= RadiusInner
    End Function

    Public Function ConstrainCursor(Location As System.Drawing.Point, width As Integer) As Drawing.Point
        Dim frustrum As System.Drawing.Rectangle = Me.Owner.ClientRectangle
        frustrum.Inflate(-(width + 5), -(width + 5))
        Location.X = System.Math.Max(Location.X, frustrum.Left)
        Location.X = System.Math.Min(Location.X, frustrum.Right)
        Location.Y = System.Math.Max(Location.Y, frustrum.Top)
        Location.Y = System.Math.Min(Location.Y, frustrum.Bottom)
        Return Location
    End Function

    Public Function IsInputParameter(Param As IGH_Param) As Boolean
        Dim isInput As New Boolean
        Select Case Param.Kind
            Case GH_ParamKind.input
                isInput = True
            Case GH_ParamKind.output
                isInput = False
            Case GH_ParamKind.floating
                If Not ParamOfWire.Attributes.HasInputGrip Then
                    isInput = False
                ElseIf Not ParamOfWire.Attributes.HasOutputGrip Then
                    isInput = True
                ElseIf (System.Math.Abs(Param.Attributes.OutputGrip.X - Owner.Viewport.UnprojectPoint(CenterControl).X) < System.Math.Abs(Param.Attributes.InputGrip.X - Owner.Viewport.UnprojectPoint(CenterControl).X)) Then
                    isInput = False
                Else
                    isInput = True
                End If
        End Select
        Return isInput
    End Function
#End Region

#Region "Sub"
    Private Sub SetElements()
        Boundary = Radius * 1.5
        Select Case Mode
            Case AxonMode.ByCategory
                SetElementsByCategory()
            Case AxonMode.ByFrequency
                SetElementsByFrequency()
            Case AxonMode.ByParameter
                SetElementsByParameter()
            Case AxonMode.ByRandom
                SetElementsByRandom()
        End Select
    End Sub

    Private Sub SetElementsByCategory()

        Select Case Level
            Case 0

                Dim server As GH_ComponentServer = Instances.ComponentServer
                Elements.Clear()
                TabSaved = Nothing

                Dim sizeIcon2 As Integer = 16
                Dim ang As New Double
                Dim radius2 As Double = RadiusInner + sizeIcon2 * 1.2
                Dim maxPerCicle As Integer = Math.Floor(radius2 * 2 * Math.PI / Math.Sqrt(sizeIcon2 ^ 2 * 2))
                Dim paso As New Integer

                For i As Int32 = 0 To server.CompleteRibbonLayout.Tabs.Count - 1

                    Dim tab As Grasshopper.GUI.Ribbon.GH_LayoutTab = server.CompleteRibbonLayout.Tabs(i)

                    If paso >= maxPerCicle Then
                        paso = 0
                        radius2 += sizeIcon2 * 1.5
                        ang = 0
                        maxPerCicle = Math.Floor((radius2 * 2 * Math.PI - sizeIcon2 * 0.5) / Math.Sqrt(sizeIcon2 ^ 2 * 2))
                    End If

                    Dim pt As Drawing.Point = GH_Convert.ToPoint(CenterControl)
                    pt.X += Math.Sin(ang) * radius2
                    pt.Y += Math.Cos(ang) * radius2

                    Dim rec As New Rectangle(pt.X - sizeIcon2 * 0.5F, pt.Y - sizeIcon2 * 0.5F, sizeIcon2, sizeIcon2)
                    Elements.Add(New AxonElement(tab.Name, New Bitmap(server.GetCategoryIcon(tab.Name), New Size(sizeIcon2, sizeIcon2)), rec, tab))

                    ang += (2 * Math.PI) / Math.Floor((radius2 * 2 * Math.PI / Math.Sqrt(sizeIcon2 ^ 2 * 2)))
                    paso += 1

                Next
                If radius2 > Boundary Then Boundary = radius2 + sizeIcon2 * 1.5

            Case 1
                If Selected = -1 Then Exit Sub

                Dim tab As GUI.Ribbon.GH_LayoutTab = DirectCast(Elements(Selected).Element, GUI.Ribbon.GH_LayoutTab)
                TabSaved = tab
                Elements.Clear()

                Dim ang As Double = 0
                Dim Radius2 As Double = Radius
                Dim paso As Integer = 0
                Dim Sides As Integer = tab.Panels.Count
                Dim maxPerCicle As Integer = Math.Floor((Radius * 2 * Math.PI / Sides) / Math.Sqrt(SizeIcon ^ 2 * 2))
                If 2 * Math.PI * Radius / Sides < Math.Sqrt(SizeIcon ^ 2 * 2) Then Radius2 += SizeIcon * 1.3

                For j As Int32 = 0 To Sides - 1

                    Dim panel As GUI.Ribbon.GH_LayoutPanel = TabSaved.Panels(j)
                    If 2 * Math.PI * Radius / Sides < Math.Sqrt(SizeIcon ^ 2 * 2) Then
                        Radius2 = Radius + SizeIcon * 1.3
                    Else
                        Radius2 = Radius
                    End If
                    paso = 0
                    ang = 2 * Math.PI * (j / Sides)
                    maxPerCicle = Math.Max(1, Math.Floor((Radius * 2 * Math.PI / Sides) / Math.Sqrt(SizeIcon ^ 2 * 2)))

                    For i As Int32 = 0 To panel.Items.Count - 1

                        If paso >= maxPerCicle - 1 Then
                            paso = 0
                            If i <> 0 Then Radius2 += SizeIcon * 1.3
                            ang = 2 * Math.PI * (j / Sides)

                            maxPerCicle = Math.Max(1, Math.Floor(((Radius2 * 2 * Math.PI / Sides)) / Math.Sqrt(SizeIcon ^ 2 * 2)))
                        End If

                        If maxPerCicle = 1 Then
                            ang += ((2 * Math.PI / Sides) / Math.Max(1, Math.Floor(((Radius2 * 2 * Math.PI / Sides) / Math.Sqrt(SizeIcon ^ 2 * 2)))) * 0.5)
                        Else
                            ang += (2 * Math.PI / Sides) / Math.Max(1, Math.Floor(((Radius2 * 2 * Math.PI / Sides) / Math.Sqrt(SizeIcon ^ 2 * 2))))
                        End If

                        Dim pt As Drawing.Point = GH_Convert.ToPoint(CenterControl)
                        pt.X += Math.Sin(ang) * Radius2
                        pt.Y += Math.Cos(ang) * Radius2

                        Dim rec As New Rectangle(pt.X - SizeIcon * 0.5, pt.Y - SizeIcon * 0.5, SizeIcon, SizeIcon)
                        paso += 1

                        Dim proxy As IGH_ObjectProxy = Instances.ComponentServer.EmitObjectProxy(panel.Items(i).Id)
                        Elements.Add(New AxonElement(proxy.Desc.Name, New Bitmap(proxy.Icon, New Size(SizeIcon, SizeIcon)), rec, proxy))

                    Next
                    If Radius2 > Boundary Then Boundary = Radius2 + SizeIcon * 1.5

                Next

        End Select

    End Sub

    Private Sub SetElementsByFrequency()

        Elements.Clear()

        Try

            Dim connections As New List(Of IGH_ObjectProxy)
            Dim ParamIndex As New List(Of Integer)

            If ParamOfWire IsNot Nothing Then

                Dim TopLevelOfParamWire As IGH_DocumentObject = ParamOfWire.Attributes.GetTopLevel.DocObject
                Dim isInput As Boolean = IsInputParameter(ParamOfWire)
                Dim index As Integer = XML.GetIndexOfParameter(TopLevelOfParamWire, ParamOfWire, isInput)

                If AxonXML.XComponentExists(TopLevelOfParamWire.ComponentGuid) Then
                    Dim connected As List(Of Tuple(Of IGH_ObjectProxy, Integer)) = XML.GetConnectedComponents(_Max, TopLevelOfParamWire.ComponentGuid, index, isInput)
                    If connected IsNot Nothing AndAlso connected.Count <> 0 Then
                        connections = (From c In connected Select c.Item1).ToList()
                        ParamIndex = (From c In connected Select c.Item2).ToList()
                    End If
                Else
                    connections = XML.GetTopComponents(_Max)
                End If

                If connections Is Nothing Or connections.Count = 0 Then
                    connections = XML.GetTopComponents(_Max)
                End If

            Else
                    connections = XML.GetTopComponents(_Max)
            End If

            If connections.Count = 0 Then Exit Sub

            Dim maxPerCicle As Integer = Math.Floor(Radius * 2 * Math.PI / Math.Sqrt(SizeIcon ^ 2 * 2))
            Dim ang As New Double
            Dim radius2 As Double = Radius
            Dim paso As New Integer

            For i As Int32 = 0 To connections.Count - 1

                Dim proxy As IGH_ObjectProxy = connections(i)
                If proxy Is Nothing Then Continue For
                Dim index As New Integer
                If i < ParamIndex.Count Then
                    index = ParamIndex(i)
                Else
                    index = 0
                End If

                If index = Nothing Then index = 0

                If paso >= maxPerCicle Then
                    paso = 0
                    radius2 += SizeIcon * 1.5
                    ang = 0
                    maxPerCicle = Math.Floor((radius2 * 2 * Math.PI - SizeIcon * 0.5) / Math.Sqrt(SizeIcon ^ 2 * 2))
                End If

                Dim pt As Drawing.Point = GH_Convert.ToPoint(CenterControl)
                pt.X += Math.Sin(ang) * radius2
                pt.Y += Math.Cos(ang) * radius2

                Dim rec As New Rectangle(pt.X - SizeIcon * 0.5F, pt.Y - SizeIcon * 0.5F, SizeIcon, SizeIcon)

                Elements.Add(New AxonElement(proxy.Desc.Name, New Bitmap(proxy.Icon, New Size(SizeIcon, SizeIcon)), rec, proxy, index))

                ang += (2 * Math.PI) / Math.Floor((radius2 * 2 * Math.PI / Math.Sqrt(SizeIcon ^ 2 * 2)))
                paso += 1

            Next
            If radius2 > Boundary Then Boundary = radius2 + SizeIcon * 1.5

        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub SetElementsByParameter()
        Elements.Clear()

        Try

            Dim connections As New List(Of IGH_ObjectProxy)
            Dim ParamIndex As New List(Of Integer)

            If ParamOfWire IsNot Nothing Then

                Dim isInput As Boolean = IsInputParameter(ParamOfWire)

                If AxonXML.XComponentExists(ParamOfWire.ComponentGuid) Then
                    Dim connected As List(Of Tuple(Of IGH_ObjectProxy, Integer)) = XML.GetConnectedComponents(_Max, ParamOfWire.ComponentGuid, 0, isInput)
                    If connected IsNot Nothing AndAlso connected.Count <> 0 Then
                        connections = (From c In connected Select c.Item1).ToList()
                        ParamIndex = (From c In connected Select c.Item2).ToList()
                    End If
                Else
                    connections = XML.GetTopComponents(_Max)
                End If

                If connections.Count = 0 Then
                    connections = XML.GetTopComponents(_Max)
                End If

            Else
                connections = XML.GetTopComponents(_Max)
            End If


            Dim maxPerCicle As Integer = Math.Floor(Radius * 2 * Math.PI / Math.Sqrt(SizeIcon ^ 2 * 2))
            Dim ang As New Double
            Dim radius2 As Double = Radius
            Dim paso As New Integer

            For i As Int32 = 0 To connections.Count - 1

                Dim proxy As IGH_ObjectProxy = connections(i)
                If proxy Is Nothing Then Continue For
                Dim index As New Integer
                If i < ParamIndex.Count Then
                    index = ParamIndex(i)
                Else
                    index = 0
                End If

                If index = Nothing Then index = 0

                If paso >= maxPerCicle Then
                    paso = 0
                    radius2 += SizeIcon * 1.5
                    ang = 0
                    maxPerCicle = Math.Floor((radius2 * 2 * Math.PI - SizeIcon * 0.5) / Math.Sqrt(SizeIcon ^ 2 * 2))
                End If

                Dim pt As Drawing.Point = GH_Convert.ToPoint(CenterControl)
                pt.X += Math.Sin(ang) * radius2
                pt.Y += Math.Cos(ang) * radius2

                Dim rec As New Rectangle(pt.X - SizeIcon * 0.5F, pt.Y - SizeIcon * 0.5F, SizeIcon, SizeIcon)

                Elements.Add(New AxonElement(proxy.Desc.Name, New Bitmap(proxy.Icon, New Size(SizeIcon, SizeIcon)), rec, proxy, index))

                ang += (2 * Math.PI) / Math.Floor((radius2 * 2 * Math.PI / Math.Sqrt(SizeIcon ^ 2 * 2)))
                paso += 1

            Next
            If radius2 > Boundary Then Boundary = radius2 + SizeIcon * 1.5

        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub SetElementsByRandom()
        Elements.Clear()

        Try

            Dim connections As New List(Of IGH_ObjectProxy)
            Dim ParamIndex As New List(Of Integer)

            If ParamOfWire IsNot Nothing Then

                Dim TopLevelOfParamWire As IGH_DocumentObject = ParamOfWire.Attributes.GetTopLevel.DocObject
                Dim isInput As Boolean = IsInputParameter(ParamOfWire)
                Dim index As Integer = XML.GetIndexOfParameter(TopLevelOfParamWire, ParamOfWire, isInput)

                If AxonXML.XComponentExists(TopLevelOfParamWire.ComponentGuid) Then
                    Dim connected As List(Of Tuple(Of IGH_ObjectProxy, Integer)) = XML.GetConnectedRandomComponents(_Max, TopLevelOfParamWire.ComponentGuid, index, isInput)
                    If connected IsNot Nothing AndAlso connected.Count <> 0 Then
                        connections = (From c In connected Select c.Item1).ToList()
                        ParamIndex = (From c In connected Select c.Item2).ToList()
                    End If
                Else
                    connections = XML.GetRandomComponents(_Max)
                End If

                If connections.Count = 0 Then
                    connections = XML.GetRandomComponents(_Max)
                End If

            Else
                connections = XML.GetRandomComponents(_Max)
            End If


            Dim maxPerCicle As Integer = Math.Floor(Radius * 2 * Math.PI / Math.Sqrt(SizeIcon ^ 2 * 2))
            Dim ang As New Double
            Dim radius2 As Double = Radius
            Dim paso As New Integer

            For i As Int32 = 0 To connections.Count - 1

                Dim proxy As IGH_ObjectProxy = connections(i)
                If proxy Is Nothing Then Continue For
                Dim index As New Integer
                If i < ParamIndex.Count Then
                    index = ParamIndex(i)
                Else
                    index = 0
                End If

                If index = Nothing Then index = 0

                If paso >= maxPerCicle Then
                    paso = 0
                    radius2 += SizeIcon * 1.5
                    ang = 0
                    maxPerCicle = Math.Floor((radius2 * 2 * Math.PI - SizeIcon * 0.5) / Math.Sqrt(SizeIcon ^ 2 * 2))
                End If

                Dim pt As Drawing.Point = GH_Convert.ToPoint(CenterControl)
                pt.X += Math.Sin(ang) * radius2
                pt.Y += Math.Cos(ang) * radius2

                Dim rec As New Rectangle(pt.X - SizeIcon * 0.5F, pt.Y - SizeIcon * 0.5F, SizeIcon, SizeIcon)

                Elements.Add(New AxonElement(proxy.Desc.Name, New Bitmap(proxy.Icon, New Size(SizeIcon, SizeIcon)), rec, proxy, index))

                ang += (2 * Math.PI) / Math.Floor((radius2 * 2 * Math.PI / Math.Sqrt(SizeIcon ^ 2 * 2)))
                paso += 1

            Next
            If radius2 > Boundary Then Boundary = radius2 + SizeIcon * 1.5

        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(ex.ToString)
        End Try
    End Sub

    Public Sub ClearAll()
        Elements.Clear()
        Level = 0
        _Selected = -1
        ParamOfWire = Nothing
        CenterControl = Nothing
    End Sub

    Public Sub InstantiateObject()
        If (Me.Owner Is Nothing) Or (Selected = -1) Then
            Exit Sub
        End If
        Try
            Dim proxy As IGH_ObjectProxy = DirectCast(Elements(Selected).Element, IGH_ObjectProxy)
            If proxy Is Nothing Then Exit Sub
            Dim comp As IGH_DocumentObject = proxy.CreateInstance()
            comp.NewInstanceGuid()
            If comp.Attributes Is Nothing Then comp.CreateAttributes()
            comp.Attributes.Pivot = Owner.Viewport.UnprojectPoint(CenterControl)
            '  comp.Attributes.ExpireLayout()
            ' comp.Attributes.PerformLayout()
            Owner.Document.AutoSave(GH_AutoSaveTrigger.object_added)
            Owner.Document.UndoUtil.RecordAddObjectEvent("Add " + comp.Name, comp)
            Owner.Document.AddObject(comp, False, Owner.Document.ObjectCount)
            comp.ExpireSolution(True)

        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(ex.ToString)
        End Try
    End Sub

    Public Sub InstantiateObject(Param As IGH_Param)
        If (Param Is Nothing) Or (Owner.Document Is Nothing) Or (Selected = -1) Then
            Exit Sub
        End If
        Try
            Dim proxy As IGH_ObjectProxy = DirectCast(Elements(Selected).Element, IGH_ObjectProxy)
            Dim index As Integer = CInt(Elements(Selected).Tag)

            If proxy Is Nothing Then
                Rhino.RhinoApp.WriteLine("Cannot convert to IGH_ObjectProxy an object of type " & Elements(Selected).Element.GetType().ToString)
                Exit Sub
            End If

            Dim obj As IGH_DocumentObject = proxy.CreateInstance()
            If obj Is Nothing Then Exit Sub
            obj.NewInstanceGuid()
            If obj.Attributes Is Nothing Then obj.CreateAttributes()
            Dim Pivot As Drawing.PointF = Owner.Viewport.UnprojectPoint(CenterControl)
            obj.Attributes.Pivot = Pivot
            ' obj.Attributes.ExpireLayout()
            Owner.Document.AutoSave(GH_AutoSaveTrigger.object_added)
            Owner.Document.UndoUtil.RecordAddObjectEvent("Add " + obj.Name, obj)
            Owner.Document.AddObject(obj, False, Owner.Document.ObjectCount)
            '    Owner.Refresh()

            'AddSource.
            Dim IsInput As Boolean = IsInputParameter(Param)

            If TypeOf obj Is GH_Component Then
                Dim C As GH_Component = TryCast(obj, GH_Component)
                If C Is Nothing Then Exit Sub
                Dim input As IGH_Param = Nothing, output As IGH_Param = Nothing
                Dim TopLevel As IGH_DocumentObject = Param.Attributes.GetTopLevel.DocObject

                If IsInput Then
                    input = Param
                    output = C.Params.Output(index)
                Else
                    input = C.Params.Input(index)
                    output = Param
                End If
                If Not Owner.Validator.CanCreateWire(input, output) Then
                    Rhino.RhinoApp.WriteLine("No se puede crear cable")
                    Exit Sub
                End If

                Owner.Document.UndoUtil.RecordWireEvent("New wire", input)
                input.AddSource(output)

            Else
                If TypeOf obj Is IGH_Param Then
                    Dim C As IGH_Param = TryCast(obj, IGH_Param)
                    If C Is Nothing Then Exit Sub
                    Dim input As IGH_Param = Nothing, output As IGH_Param = Nothing
                    If IsInput Then
                        input = Param
                        output = C
                    Else
                        input = C
                        output = Param
                    End If

                    If Not Owner.Validator.CanCreateWire(input, output) Then
                        Rhino.RhinoApp.WriteLine("No se puede crear cable")
                        Exit Sub
                    End If

                    Owner.Document.UndoUtil.RecordWireEvent("New wire", input)
                    input.AddSource(output)

                Else
                    Rhino.RhinoApp.WriteLine("Ni comp ni param, " & obj.GetType().ToString)
                End If
            End If
            ' Owner.Refresh()
            Owner.Document.NewSolution(True)

        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(ex.ToString)
        End Try
    End Sub

#End Region

#Region "Render"
    Public Overrides Sub Render(Canvas As GH_Canvas)
        If Not IsActive Then Exit Sub

        Canvas.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality
        Dim xform As System.Drawing.Drawing2D.Matrix = Canvas.Graphics.Transform
        Canvas.Graphics.ResetTransform()
        RenderBackGroundDisk(Canvas.Graphics)
        RenderIcons(Canvas.Graphics)

        Canvas.Graphics.Transform = xform

        '  sender.Viewport.ApplyProjection(sender.Graphics)
    End Sub

    Private Sub RenderBackGroundDisk(G As System.Drawing.Graphics)
        Try

            Dim RectInnerEdge As New Drawing.Rectangle(CenterControl.X - RadiusInner, CenterControl.Y - RadiusInner, RadiusInner * 2, RadiusInner * 2)
            RectInnerEdge.Inflate(-1, -1)

            'Cross.
            Dim PenCross As New Drawing.Pen(Color.DarkGray, 1.0F)
            G.DrawLine(PenCross, CenterControl.X + 2, CenterControl.Y + 2, CenterControl.X - 2, CenterControl.Y - 2)
            G.DrawLine(PenCross, CenterControl.X - 2, CenterControl.Y + 2, CenterControl.X + 2, CenterControl.Y - 2)

            'Background dick.
            Dim bpath As New Drawing.Drawing2D.GraphicsPath()
            bpath.AddEllipse(RectInnerEdge)
            Dim RectOuterEdge As Drawing.Rectangle = RectInnerEdge
            RectOuterEdge.Inflate(Boundary - RadiusInner, Boundary - RadiusInner)
            bpath.AddEllipse(RectOuterEdge)

            Dim grad As New Drawing.Drawing2D.PathGradientBrush(bpath)
            grad.CenterPoint = New Drawing.PointF(CenterControl.X, CenterControl.Y)
            grad.CenterColor = System.Drawing.Color.FromArgb(200, 50, 50, 50)
            grad.SurroundColors = New Color(0) {System.Drawing.Color.FromArgb(0, 50, 50, 50)}
            Dim blend As New Drawing2D.ColorBlend()
            blend.Colors = New Color(2) {
                  System.Drawing.Color.FromArgb(0, 50, 50, 50),
                  System.Drawing.Color.FromArgb(160, 50, 50, 50),
                  System.Drawing.Color.FromArgb(255, 50, 50, 50)}
            blend.Positions = New Single(2) {0.0F, 1 - RadiusInner / Boundary, 1.0F}
            grad.InterpolationColors = blend
            G.FillPath(grad, bpath)

            'Separators.
            If Mode = AxonMode.ByCategory AndAlso Level > 0 AndAlso TabSaved IsNot Nothing Then
                Dim Sides As Integer = TabSaved.Panels.Count

                Dim ang As New Double
                Dim ang2 As Double = 360 / Sides

                For j As Int32 = 0 To Sides - 1

                    ang = 2 * Math.PI * (j / Sides)

                    Dim ptl As Drawing.Point = GH_Convert.ToPoint(CenterControl)
                    ptl.X += (RadiusInner + 1) * Math.Sin(ang)
                    ptl.Y += (RadiusInner + 1) * Math.Cos(ang)
                    Dim ptl1 As Drawing.Point = ptl
                    ptl1.X += (Boundary - RadiusInner + 1) * Math.Sin(ang)
                    ptl1.Y += (Boundary - RadiusInner + 1) * Math.Cos(ang)

                    Dim pgrad As New Drawing2D.LinearGradientBrush(ptl, ptl1, Color.FromArgb(160, 50, 50, 50), Color.FromArgb(0, 50, 50, 50))
                    Dim pl As New Pen(pgrad)
                    pl.DashPattern = New Single(1) {3, 3}
                    G.DrawLine(pl, ptl, ptl1)
                    pl.Dispose()
                    pgrad.Dispose()


                    Dim textToDisplay As String = TabSaved.Panels(j).Name
                    Dim TextFont As Font = Grasshopper.Kernel.GH_FontServer.Standard
                    Dim TextSize As Size = TextRenderer.MeasureText(textToDisplay, TextFont)

                    Dim ptl2 As Drawing.Point = GH_Convert.ToPoint(CenterControl)
                    Dim offset As Integer = (Boundary - RadiusInner) / 2 + RadiusInner
                    If offset - TextSize.Width * 0.5 < RadiusInner + 3 Then
                        offset = RadiusInner + 3 + TextSize.Width * 0.5
                    End If
                    ptl2.X += offset * Math.Sin(ang + Math.Sin((TextSize.Height * 0.6) / offset))
                    ptl2.Y += offset * Math.Cos(ang + Math.Sin((TextSize.Height * 0.6) / offset))


                    Dim TextRect As New RectangleF(ptl2.X - TextSize.Width * 0.5, ptl2.Y - TextSize.Height * 0.5, TextSize.Width, TextSize.Height)

                    Dim m As New Drawing.Drawing2D.Matrix()

                    If j <= Sides \ 2 Then
                        m.RotateAt(-j * (360 / Sides) + 90, ptl2)
                    Else
                        m.RotateAt(-j * (360 / Sides) - 90, ptl2)
                    End If

                    G.Transform = m
                    G.DrawString(textToDisplay, TextFont, Brushes.DimGray, TextRect)
                    G.ResetTransform()
                Next
            End If

            'Circle edge.
            Dim discPath As New System.Drawing.Drawing2D.GraphicsPath()
            discPath.AddEllipse(RectInnerEdge)
            Dim discInnerFill As New System.Drawing.Drawing2D.LinearGradientBrush(RectInnerEdge,
                  System.Drawing.Color.FromArgb(150, System.Drawing.Color.White), System.Drawing.Color.FromArgb(0, System.Drawing.Color.White), System.Drawing.Drawing2D.LinearGradientMode.Vertical)
            discInnerFill.WrapMode = System.Drawing.Drawing2D.WrapMode.TileFlipXY
            Dim discInnerEdge As New System.Drawing.Pen(discInnerFill, 2.0F)
            Dim discInnerBorder As System.Drawing.Rectangle = RectInnerEdge
            discInnerBorder.Inflate(1, 1)
            G.DrawEllipse(discInnerEdge, discInnerBorder)

            Dim discEdge As New System.Drawing.Pen(System.Drawing.Color.FromArgb(160, System.Drawing.Color.Black), 1.0F)
            G.DrawEllipse(discEdge, RectInnerEdge)


            PenCross.Dispose()

            bpath.Dispose()
            grad.Dispose()

            discPath.Dispose()
            discInnerEdge.Dispose()
            discInnerFill.Dispose()
            discEdge.Dispose()

        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub RenderIcons(G As System.Drawing.Graphics)
        For i As Int32 = 0 To Elements.Count - 1
            Dim rec As Rectangle = Elements(i).Rectangle
            If (_Selected = i) Then
                Dim selrec As Rectangle = rec
                selrec.Inflate(2, 2)
                selrec.Width -= 1
                selrec.Height -= 1
                Grasshopper.GUI.GH_GraphicsUtil.RenderHighlightBox(G, selrec, 2)
            End If
            Grasshopper.GUI.GH_GraphicsUtil.RenderCenteredIcon(G, rec, Elements(i).Icon)
        Next
    End Sub
#End Region

#Region "ToolTip"
    Public Overrides Sub SetupTooltip(canvasPoint As PointF, e As GH_TooltipDisplayEventArgs)
        If Selected <> -1 Then

            Dim p As IGH_ObjectProxy = TryCast(Elements(Selected).Element, IGH_ObjectProxy)
            If p IsNot Nothing Then
                If p.Desc.NickName IsNot Nothing AndAlso p.Desc.NickName <> String.Empty Then
                    e.Title = String.Format("{0} ({1})", p.Desc.Name, p.Desc.NickName)
                Else
                    e.Title = p.Desc.Name
                End If
                e.Description = p.Desc.Description
                e.Icon = p.Icon
            Else
                e.Title = Elements(Selected).Name
                If Elements(Selected).Icon IsNot Nothing Then e.Icon = Elements(Selected).Icon
            End If
        Else
            MyBase.SetupTooltip(canvasPoint, e)
        End If

    End Sub

    Public Overrides Function IsTooltipRegion(canvas_coordinate As PointF) As Boolean
        Return Selected <> -1
    End Function
#End Region

#Region "Menu"
    Public Overrides Sub AppendToMenu(menu As System.Windows.Forms.ToolStripDropDownMenu)
        MyBase.AppendToMenu(menu)
        GH_DocumentObject.Menu_AppendSeparator(menu)
        GH_DocumentObject.Menu_AppendItem(menu, "By category", New EventHandler(AddressOf Menu_ByCategory), True, _Mode = AxonMode.ByCategory)
        GH_DocumentObject.Menu_AppendItem(menu, "By frequency", New EventHandler(AddressOf Menu_ByFrequency), True, _Mode = AxonMode.ByFrequency)
        GH_DocumentObject.Menu_AppendItem(menu, "By parameter", New EventHandler(AddressOf Menu_ByParameter), True, _Mode = AxonMode.ByParameter)
        GH_DocumentObject.Menu_AppendItem(menu, "By random", New EventHandler(AddressOf menu_ByRandom), True, _Mode = AxonMode.ByRandom)
        GH_DocumentObject.Menu_AppendSeparator(menu)
        GH_DocumentObject.Menu_AppendItem(menu, "Attributes", New EventHandler(AddressOf Menu_Attributes))
    End Sub

    Private Sub Menu_ByCategory(sender As Object, e As EventArgs)
        If (Me.Mode = AxonMode.ByCategory) Then
            Exit Sub
        Else
            Me.Mode = AxonMode.ByCategory
        End If
    End Sub

    Private Sub Menu_ByFrequency(sender As Object, e As EventArgs)
        If (Me.Mode = AxonMode.ByFrequency) Then
            Exit Sub
        Else
            If Not XML.Loaded Then
                Me.Mode = AxonMode.ByCategory
                Exit Sub
            End If
            Me.Mode = AxonMode.ByFrequency
        End If
    End Sub

    Private Sub Menu_ByParameter(sender As Object, e As EventArgs)
        If (Me.Mode = AxonMode.ByParameter) Then
            Exit Sub
        Else
            If Not XML.Loaded Then
                Me.Mode = AxonMode.ByCategory
                Exit Sub
            End If
            Me.Mode = AxonMode.ByParameter
        End If
    End Sub

    Private Sub Menu_ByRandom(sender As Object, e As EventArgs)
        If (Me.Mode = AxonMode.ByRandom) Then
            Exit Sub
        Else
            If Not XML.Loaded Then
                Me.Mode = AxonMode.ByCategory
                Exit Sub
            End If
            Me.Mode = AxonMode.ByRandom
        End If
    End Sub

    Public AttForm As AttributesForm

    Private Sub Menu_Attributes(sender As Object, e As EventArgs)
        If AttForm Is Nothing Then
            AttForm = New AttributesForm(Me)
        Else
            AttForm.Dispose()
            AttForm = Nothing
        End If
    End Sub
#End Region

End Class

Public Class AxonElement

    Public Name As String
    Public Icon As Bitmap
    Public Rectangle As Rectangle
    Public Element As Object
    Public Tag As Object

    Sub New(Nombre As String, Icono As Bitmap, Rect As Rectangle, Elemento As Object)
        Name = Nombre
        Icon = Icono
        Rectangle = Rect
        Element = Elemento
    End Sub

    Sub New(Nombre As String, Icono As Bitmap, Rect As Rectangle, Elemento As Object, Etiqueta As Object)
        Name = Nombre
        Icon = Icono
        Rectangle = Rect
        Element = Elemento
        Tag = Etiqueta
    End Sub

    Public Function Contains(ControlPoint As Drawing.Point) As Boolean
        If (Rectangle = Nothing) Then Return False
        Return Rectangle.Contains(ControlPoint)
    End Function

End Class

Public Enum AxonMode
    None = 0
    ByCategory = 1
    ByFrequency = 2
    ByParameter = 3
    ByRandom = 4
End Enum

Public Class AxonXML

    Private WithEvents ghDoc As GH_Document
    Private XDoc As System.Xml.Linq.XDocument
    Private Server As Grasshopper.Kernel.GH_ComponentServer
    Private _File As String
    Private OnUse As Boolean
    Private AddedCounter As New Integer

    Sub New(Doc As GH_Document)
        If Doc Is Nothing Then Exit Sub
        ghDoc = Doc
        Server = Instances.ComponentServer
        _File = My.Application.Info.DirectoryPath & "\AxonConnectivityDatabase.xml"

        Try
            If Not IO.File.Exists(_File) Then
                Dim ofd As New System.Windows.Forms.OpenFileDialog()
                ofd.CheckFileExists = True
                ofd.CheckPathExists = True
                ofd.DefaultExt = "*.xml"
                ofd.Filter = "*.xml|*.xml"
                ofd.InitialDirectory = My.Application.Info.DirectoryPath
                ofd.Multiselect = False
                ofd.Title = "Axon Connectivity Database not found. Open a valid .xml file"

                If ofd.ShowDialog(Grasshopper.Instances.DocumentEditor) = DialogResult.OK Then
                    _File = ofd.FileName
                Else
                    OnUse = False
                    Exit Sub
                End If
            End If

            XDoc = XDocument.Load(_File)

            If XDoc Is Nothing Then
                Rhino.RhinoApp.WriteLine("Axon database could not be loaded")
                OnUse = False
                Exit Sub
            End If

            If XDoc.Root Is Nothing Or Not XDoc.Root.Name.LocalName.Equals("axondatabase") Then
                ResetRoot()
            End If


            OnUse = True

            Create()

        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(ex.ToString)
        End Try
    End Sub

#Region "Property"
    Public ReadOnly Property GhDocument As GH_Document
        Get
            Return ghDoc
        End Get
    End Property

    Public ReadOnly Property XmlDocument As XDocument
        Get
            Return XDoc
        End Get
    End Property

    Public Property FilePath As String
        Get
            Return _File
        End Get
        Set(value As String)
            _File = value
            XDoc = XDocument.Load(value)
        End Set
    End Property

    Public ReadOnly Property Loaded As Boolean
        Get
            Return OnUse
        End Get
    End Property

    Public ReadOnly Property XComponents As XElement
        Get
            Return XDoc.Root.Element("components")
        End Get
    End Property

    Public ReadOnly Property XComponentNodes As IEnumerable(Of XElement)
        Get
            Return XDoc.Root.Element("components").Elements("component")
        End Get
    End Property

    Public ReadOnly Property XDocuments As XElement
        Get
            Return XDoc.Root.Element("documents")
        End Get
    End Property

    Public ReadOnly Property XDocumentNodes As IEnumerable(Of XElement)
        Get
            Return XDoc.Root.Element("documents").Elements("document")
        End Get
    End Property

    Public Property XComponentsCount As Integer
        Get
            Return CInt(XDoc.Root.Element("components").Attribute("count").Value)
        End Get
        Set(value As Integer)
            XDoc.Root.Element("components").Attribute("count").Value = value
        End Set
    End Property

    Public Property XDocumentsCount As Integer
        Get
            Return CInt(XDoc.Root.Element("documents").Attribute("count").Value)
        End Get
        Set(value As Integer)
            XDoc.Root.Element("documents").Attribute("count").Value = value
        End Set
    End Property

#End Region

#Region "Events"
    Private Sub Create()

        AddHandler ghDoc.ObjectsAdded, AddressOf AddedObjects

        For Each obj As IGH_DocumentObject In ghDoc.Objects
            If TypeOf obj Is IGH_Param Then
                AddHandler obj.ObjectChanged, AddressOf WireEvent

            ElseIf TypeOf obj Is IGH_Component Then
                Dim C As IGH_Component = DirectCast(obj, IGH_Component)
                For Each p As IGH_Param In C.Params
                    AddHandler p.ObjectChanged, AddressOf WireEvent
                Next
            End If
        Next
    End Sub

    Public Sub Destroy()

        RemoveHandler ghDoc.ObjectsAdded, AddressOf AddedObjects

        For Each obj As IGH_DocumentObject In ghDoc.Objects
            If TypeOf obj Is IGH_Param Then
                RemoveHandler obj.ObjectChanged, AddressOf WireEvent
            ElseIf TypeOf obj Is IGH_Component Then
                For Each p As IGH_Param In DirectCast(obj, IGH_Component).Params
                    RemoveHandler p.ObjectChanged, AddressOf WireEvent
                Next
            End If
        Next

    End Sub

    Private Sub AddedObjects(sender As Object, e As Kernel.GH_DocObjectEventArgs)
        Try

            For Each obj As IGH_DocumentObject In e.Objects
                If obj Is Nothing Then Continue For

                If TypeOf obj Is IGH_Param Then
                    AddHandler obj.ObjectChanged, AddressOf WireEvent

                ElseIf TypeOf obj Is IGH_Component Then
                    Dim Comp As IGH_Component = DirectCast(obj, IGH_Component)
                    For Each p As IGH_Param In Comp.Params
                        AddHandler p.ObjectChanged, AddressOf WireEvent
                    Next
                Else
                    Exit Sub
                End If

                Dim proxy As IGH_ObjectProxy = GetObjectProxy(obj.Name)
                If proxy Is Nothing Then Continue For
                Dim C As XElement = XGetComponent(proxy.Guid)
                If C Is Nothing Then
                    AddNewComponent(proxy.Guid)
                Else
                    C.Attribute("frequency").Value = CInt(C.Attribute("frequency").Value) + 1
                End If

                AddedCounter += 1
                If AddedCounter >= 30 Then
                    XSortComponentsByFrecuency(True)
                    AddedCounter = 0
                End If
            Next
            Save()

        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub WireEvent(sender As IGH_DocumentObject, e As GH_ObjectChangedEventArgs)
        Try
            If e.Type = GH_ObjectEventType.Sources Then
                If TypeOf sender Is IGH_Param Then
                    Dim targetParam As IGH_Param = DirectCast(sender, IGH_Param)
                    If targetParam.SourceCount < 1 Then Exit Sub
                    Dim sourceParam As IGH_Param = targetParam.Sources.Last()
                    Dim source As IGH_DocumentObject = sourceParam.Attributes.GetTopLevel.DocObject
                    Dim target As IGH_DocumentObject = targetParam.Attributes.GetTopLevel.DocObject
                    Dim proxysource As IGH_ObjectProxy = GetObjectProxy(source.Name)
                    If proxysource Is Nothing Then Exit Sub
                    Dim proxytarget As IGH_ObjectProxy = GetObjectProxy(target.Name)
                    If proxytarget Is Nothing Then Exit Sub
                    Dim iOutput As Integer = Nothing
                    Dim iInput As Integer = Nothing

                    If TypeOf source Is IGH_Component Then
                        Dim sComp As IGH_Component = DirectCast(source, IGH_Component)
                        For i As Int32 = 0 To sComp.Params.Output.Count - 1
                            If sComp.Params.Output(i).Name.Equals(sourceParam.Name) Then
                                iOutput = i
                                Exit For
                            End If
                        Next
                    ElseIf TypeOf source Is IGH_Param Then
                        iOutput = 0
                    End If

                    If TypeOf target Is IGH_Component Then
                        Dim tComp As IGH_Component = DirectCast(target, IGH_Component)
                        For i As Int32 = 0 To tComp.Params.Input.Count - 1
                            If tComp.Params.Input(i).Name.Equals(targetParam.Name) Then
                                iInput = i
                                Exit For
                            End If
                        Next
                    ElseIf TypeOf target Is IGH_Param Then
                        iInput = 0
                    End If

                    SerializeConnection(proxysource.Guid, iOutput, iInput, proxytarget.Guid)
                End If

            End If
        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Serialize XML"
    Public Sub SerializeConnection(SourceID As Guid, outputIndex As Integer, inputIndex As Integer, targetID As Guid)
        Try
            Dim XsourceComp As XElement = XGetComponent(SourceID)
            If XsourceComp Is Nothing Then XsourceComp = AddNewComponent(SourceID)
            If XsourceComp.Elements("output").Any() Then
                Dim Xoutput As XElement = (From p In XsourceComp.Elements("output") Where CInt(p.Attribute("index").Value).Equals(outputIndex) Select p).SingleOrDefault()
                If Xoutput IsNot Nothing Then
                    If Xoutput.HasElements Then
                        Dim Xtarget As XElement = (From p In Xoutput.Elements("target") Where Guid.Parse(CStr(p.Attribute("id").Value)).Equals(targetID) AndAlso CInt(p.Attribute("input").Value) = inputIndex Select p).SingleOrDefault()
                        If Xtarget Is Nothing Then
                            Xoutput.Add(New XElement("target", New Object(2) {New XAttribute("id", targetID), New XAttribute("input", inputIndex), 1}))
                        Else
                            Xtarget.Value = CInt(Xtarget.Value) + 1
                        End If
                    Else
                        Xoutput.Add(New XElement("target", New Object(2) {New XAttribute("id", targetID), New XAttribute("input", inputIndex), 1}))
                    End If
                End If
            End If
            Dim XTargetComp As XElement = XGetComponent(targetID)
            If XTargetComp Is Nothing Then XTargetComp = AddNewComponent(targetID)
            If XTargetComp.Elements("input").Any() Then
                Dim Xinput As XElement = (From p In XTargetComp.Elements("input") Where CInt(p.Attribute("index").Value).Equals(inputIndex) Select p).SingleOrDefault()
                If Xinput IsNot Nothing Then
                    If Xinput.HasElements Then
                        Dim Xsource As XElement = (From p In Xinput.Elements("source") Where Guid.Parse(CStr(p.Attribute("id").Value)).Equals(SourceID) AndAlso CInt(p.Attribute("output").Value) = outputIndex Select p).SingleOrDefault()
                        If Xsource Is Nothing Then
                            Xinput.Add(New XElement("source", New Object(2) {New XAttribute("id", SourceID), New XAttribute("output", outputIndex), 1}))
                        Else
                            Xsource.Value = CInt(Xsource.Value) + 1
                        End If
                    Else
                        Xinput.Add(New XElement("source", New Object(2) {New XAttribute("id", SourceID), New XAttribute("output", outputIndex), 1}))
                    End If
                End If
            End If
            Save()

        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(vbCrLf & "AxonXML.SerializeConnection() exception at " & Date.Now() & vbCrLf & ex.ToString)
        End Try
    End Sub

    Public Sub WriteFromFile(GhFile As String)
        If GhFile Is Nothing Or GhFile Is String.Empty Then Exit Sub
        If Not IO.File.Exists(GhFile) Then Exit Sub

        Try
            Dim archive As New GH_IO.Serialization.GH_Archive()
            If Not archive.ReadFromFile(GhFile) Then
                Rhino.RhinoApp.WriteLine("The file could Not be read" & vbCrLf & GhFile)
                Exit Sub
            End If
            Dim D As New GH_Document()
            archive.ClearMessages()
            archive.ExtractObject(D, "Definition")
            archive.ClearMessages()
            D.DestroyProxySources()

            If XDocumentExists(D.DocumentID) Then Exit Sub

            XDocuments.Add(New XElement("document", D.DocumentID))
            XDocumentsCount += 1

            For Each obj As IGH_DocumentObject In D.Objects
                If obj.Obsolete Then Continue For
                If Not (Server.IsObjectCached(obj.ComponentGuid)) Then Continue For
                Dim proxy As IGH_ObjectProxy = Server.FindObjectByName(obj.Name, False, False)
                If proxy Is Nothing Then Continue For

                If TypeOf obj Is IGH_Component Then

                    Dim Xcomp As XElement = XGetComponent(proxy.Guid)
                    If Xcomp IsNot Nothing Then
                        Xcomp.Attribute("frequency").Value = CInt(Xcomp.Attribute("frequency").Value) + 1
                    Else
                        Xcomp = AddNewComponent(proxy.Guid)
                    End If
                    If Xcomp Is Nothing Then Continue For

                    Dim C As IGH_Component = DirectCast(obj, IGH_Component)
                    If C.Params.Input.Count = 0 Then Continue For
                    Dim proxytarget As IGH_ObjectProxy = Server.FindObjectByName(C.Name, False, False)
                    If proxytarget Is Nothing Then Continue For


                    For i As Int32 = 0 To C.Params.Input.Count - 1
                        If C.Params.Input(i).SourceCount = 0 Then Continue For
                        For Each sourceParam As IGH_Param In C.Params.Input(i).Sources
                            Dim source As IGH_DocumentObject = sourceParam.Attributes.GetTopLevel.DocObject
                            Dim proxysource As IGH_ObjectProxy = Server.FindObjectByName(source.Name, False, False)
                            If proxysource Is Nothing Then Continue For

                            Dim output As Integer = Nothing
                            If TypeOf source Is IGH_Component Then
                                Dim sourceComp As IGH_Component = DirectCast(source, IGH_Component)
                                For j As Int32 = 0 To sourceComp.Params.Output.Count - 1
                                    If sourceComp.Params.Output(j).Name.Equals(sourceParam.Name) Then
                                        output = j
                                        Exit For
                                    End If
                                Next
                            ElseIf TypeOf source Is IGH_Param Then
                                output = 0
                            End If

                            SerializeConnection(proxysource.Guid, output, i, proxytarget.Guid)
                        Next
                    Next

                ElseIf TypeOf obj Is IGH_Param Then

                    Dim Xcomp As XElement = XGetComponent(proxy.Guid)
                    If Xcomp IsNot Nothing Then
                        Xcomp.Attribute("frequency").Value = CInt(Xcomp.Attribute("frequency").Value) + 1
                    Else
                        Xcomp = AddNewComponent(proxy.Guid)
                    End If
                    If Xcomp Is Nothing Then Continue For

                    Dim P As IGH_Param = DirectCast(obj, IGH_Param)
                    If P.SourceCount = 0 Then Continue For
                    Dim proxytarget As IGH_ObjectProxy = Server.FindObjectByName(P.Name, False, False)
                    If proxytarget Is Nothing Then Continue For

                    For Each sourceParam As IGH_Param In P.Sources
                        Dim source As IGH_DocumentObject = sourceParam.Attributes.GetTopLevel.DocObject
                        Dim proxysource As IGH_ObjectProxy = Server.FindObjectByName(source.Name, False, False)
                        If proxysource Is Nothing Then Continue For

                        Dim output As Integer = Nothing
                        If TypeOf source Is IGH_Component Then
                            Dim sourceComp As IGH_Component = DirectCast(source, IGH_Component)
                            For j As Int32 = 0 To sourceComp.Params.Output.Count - 1
                                If sourceComp.Params.Output(j).Name.Equals(sourceParam.Name) Then
                                    output = j
                                    Exit For
                                End If
                            Next
                        ElseIf TypeOf source Is IGH_Param Then
                            output = 0
                        End If

                        SerializeConnection(proxysource.Guid, output, 0, proxytarget.Guid)
                    Next
                End If
            Next

        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(vbCrLf & "AxonXML.WriteFromFile() exception at " & Date.Now() & vbCrLf & ex.ToString)
        End Try

    End Sub

    Public Function AddNewComponent(ID As Guid) As XElement
        Try
            Dim XComp As New XElement("component", New Object(1) {New XAttribute("id", ID), New XAttribute("frequency", 1)})
            Dim obj As IGH_DocumentObject = Server.EmitObject(ID)
            If obj Is Nothing Then Return Nothing

            If TypeOf obj Is IGH_Component Then
                Dim comp As IGH_Component = DirectCast(obj, IGH_Component)
                For i As Int32 = 0 To comp.Params.Input.Count - 1
                    XComp.Add(New XElement("input", New XAttribute("index", i)))
                Next
                For i As Int32 = 0 To comp.Params.Output.Count - 1
                    XComp.Add(New XElement("output", New XAttribute("index", i)))
                Next

            ElseIf TypeOf obj Is IGH_Param Then
                Dim param As IGH_Param = DirectCast(obj, IGH_Param)
                param.CreateAttributes()
                If param.Attributes.HasInputGrip Then XComp.Add(New XElement("input", New XAttribute("index", 0)))
                If param.Attributes.HasOutputGrip Then XComp.Add(New XElement("output", New XAttribute("index", 0)))
            End If

            XComponents.Add(XComp)
            XComponentsCount += 1

            Return XComp
        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(vbCrLf & "AxonXML.AddNewComponent() exception at " & Date.Now() & vbCrLf & ex.ToString)
            Return Nothing
        End Try

    End Function

    Public Function XComponentExists(ID As Guid) As Boolean
        Return (From d In XComponentNodes Where Guid.Parse(CStr(d.Attribute("id").Value)).Equals(ID) Select d).Any()
    End Function

    Public Function XDocumentExists(ID As Guid) As Boolean
        Return (From d In XDocumentNodes Where Guid.Parse(CStr(d.Value)).Equals(ID) Select d).Any()
    End Function

    Public Function XGetComponent(ID As Guid) As XElement
        Return (From c In XComponentNodes Where Guid.Parse(CStr(c.Attribute("id").Value)).Equals(ID) Select c).SingleOrDefault
    End Function

    Public Sub XSortComponentsByFrecuency(SortSources As Boolean)
        Dim sorted As IEnumerable(Of XElement) = From c In XComponentNodes Let frequency = CInt(c.Attribute("frequency"))
                                                 Order By frequency Descending Select c
        If SortSources Then
            For Each comp As XElement In sorted
                For Each input As XElement In comp.Elements("input")
                    Dim sortedinput As IEnumerable(Of XElement) =
                        From i In input.Elements("source") Let frequency = CInt(i.Value)
                        Order By frequency Descending Select i
                    input.ReplaceNodes(sortedinput)
                Next
                For Each output As XElement In comp.Elements("output")
                    Dim sortedoutput As IEnumerable(Of XElement) =
                        From o In output.Elements("target") Let frequency = CInt(o.Value)
                        Order By frequency Descending Select o
                    output.ReplaceNodes(sortedoutput)
                Next
            Next
        End If
        XComponents.ReplaceNodes(sorted)
        Save()
    End Sub

    Public Sub Save()
        XDoc.Save(_File)
    End Sub

    Public Sub ResetRoot()
        Try
            Dim result As System.Windows.Forms.DialogResult = MessageBox.Show(
 "Are you sure you want to reset the content of the current Axon XML Database?" & vbCrLf & "This process Is irreversible.",
 "Reset database", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If result = DialogResult.Yes Then
                XDoc.Root.Remove()
                XDoc.Add(New XElement("axondatabase"))
                XDoc.Root.Add(New XElement("components", New XAttribute("count", 0)))
                XDoc.Root.Add(New XElement("documents", New XAttribute("count", 0)))
                Save()
                MessageBox.Show("Axon Database reset correctly.")
            End If
        Catch ex As Exception
            Rhino.RhinoApp.WriteLine("AxonXML.ResetRoot() exception at " & Date.Now() & vbCrLf & ex.ToString)
        End Try

    End Sub

#End Region

#Region "Deserialize XML"
    Public Function GetTopComponents(Max As Integer) As List(Of IGH_ObjectProxy)
        Return (From c In XComponentNodes.Take(Max) Select GetObjectProxy(Guid.Parse(CStr(c.Attribute("id").Value)))).ToList()
    End Function

    Public Function GetTopComponentsFrequency(Max As Integer) As List(Of Tuple(Of IGH_ObjectProxy, Integer))
        Return (From c In XComponentNodes.Take(Max) Select New Tuple(Of IGH_ObjectProxy, Integer)(GetObjectProxy(Guid.Parse(CStr(c.Attribute("id").Value))), CInt(c.Attribute("frequency")))).ToList()
    End Function

    Public Function GetTopComponentsLeak(MaxComponents As Integer, MaxConnected As Integer) As List(Of Tuple(Of IGH_ObjectProxy, Integer, Boolean(), Integer(), IGH_ObjectProxy(), Integer(), Integer()))

        Dim Leak As New List(Of Tuple(Of IGH_ObjectProxy, Integer, Boolean(), Integer(), IGH_ObjectProxy(), Integer(), Integer()))

        For i As Int32 = 0 To XComponentNodes.Count - 1
            If i >= MaxComponents Then Exit For

            Dim Xcomp As XElement = XComponentNodes.ElementAt(i)
            Dim proxy As IGH_ObjectProxy = GetObjectProxy(Guid.Parse(CStr(Xcomp.Attribute("id").Value)))
            If proxy Is Nothing Then Continue For
            Dim freq As Integer = CInt(Xcomp.Attribute("frequency").Value)
            Dim IsInput As New List(Of Boolean)
            Dim IndexParam As New List(Of Integer)
            Dim Connected As New List(Of IGH_ObjectProxy)
            Dim ConnectedIndexParam As New List(Of Integer)
            Dim ConnectedFrequency As New List(Of Integer)

            If Xcomp.Elements("input").Any() Then
                For j As Int32 = 0 To Xcomp.Elements("input").Count - 1
                    Dim Xinput As XElement = Xcomp.Elements("input").ElementAt(j)
                    For k As Int32 = 0 To Xinput.Elements("source").Count - 1
                        If k >= MaxConnected Then Exit For
                        Dim Xsource As XElement = Xinput.Elements("source").ElementAt(k)
                        IsInput.Add(True)
                        IndexParam.Add(CInt(Xinput.Attribute("index").Value))
                        Connected.Add(GetObjectProxy(Guid.Parse(CStr(Xsource.Attribute("id").Value))))
                        ConnectedIndexParam.Add(CInt(Xsource.Attribute("output").Value))
                        ConnectedFrequency.Add(CInt(Xsource.Value))
                    Next
                Next
            End If
            If Xcomp.Elements("output").Any() Then
                For j As Int32 = 0 To Xcomp.Elements("output").Count - 1
                    Dim Xoutput As XElement = Xcomp.Elements("output").ElementAt(j)
                    For k As Int32 = 0 To Xoutput.Elements("target").Count - 1
                        If k >= MaxConnected Then Exit For
                        Dim Xtarget As XElement = Xoutput.Elements("target").ElementAt(k)
                        IsInput.Add(False)
                        IndexParam.Add(CInt(Xoutput.Attribute("index").Value))
                        Connected.Add(GetObjectProxy(Guid.Parse(CStr(Xtarget.Attribute("id").Value))))
                        ConnectedIndexParam.Add(CInt(Xtarget.Attribute("input").Value))
                        ConnectedFrequency.Add(CInt(Xtarget.Value))
                    Next
                Next
            End If
            Leak.Add(New Tuple(Of IGH_ObjectProxy, Integer, Boolean(), Integer(), IGH_ObjectProxy(), Integer(), Integer()) _
               (proxy, freq, IsInput.ToArray(), IndexParam.ToArray(), Connected.ToArray(), ConnectedIndexParam.ToArray(), ConnectedFrequency.ToArray()))
        Next
        Return Leak
    End Function

    Public Function GetRandomComponents(Max As Integer) As List(Of IGH_ObjectProxy)
        Randomize(Date.Now().Second + Date.Now.Minute)
        Dim top As Integer = XComponentsCount - 1
        Dim obj As New List(Of IGH_ObjectProxy)
        Dim idz As New List(Of Integer)
        Dim nope As New Integer

        Do While obj.Count < Max
            Dim index As Integer = Rnd() * top
            If idz.Contains(index) Then
                If nope > 100 Then Exit Do
                nope += 1
            Else
                idz.Add(index)
                nope = 0
                obj.Add(GetObjectProxy(Guid.Parse(CStr(XComponentNodes.ElementAt(index).Attribute("id")))))
            End If
        Loop
        'For i As Int32 = 0 To Max - 1
        '    Dim index As Integer = Rnd(i) * top
        '    If idz.Contains(index) Then
        '        If nope > 10 Then Exit For
        '        i = -1
        '        nope += 1
        '        Continue For
        '    Else
        '        idz.Add(index)
        '        nope = 0
        '    End If
        '    obj.Add(GetObjectProxy(Guid.Parse(CStr(XComponentNodes.ElementAt(index).Attribute("id")))))
        'Next
        Return obj
    End Function

    Public Function GetConnectedComponents(Max As Integer, CompID As Guid, ParameterIndex As Integer, OfInput As Boolean) As List(Of Tuple(Of IGH_ObjectProxy, Integer))
        If Max = Nothing Then Max = Integer.MaxValue
        Try
            Dim XComp As XElement = (From c In XComponentNodes Where Guid.Parse(CStr(c.Attribute("id").Value)).Equals(CompID) Select c).SingleOrDefault()
            If XComp Is Nothing Then Return Nothing

            Dim connected As New List(Of Tuple(Of IGH_ObjectProxy, Integer))
            Dim ids As New List(Of Guid)

            If OfInput Then
                Dim Xparam As XElement = (From p In XComp.Elements("input") Where CInt(p.Attribute("index").Value) = ParameterIndex Select p).SingleOrDefault
                If Xparam Is Nothing Then Return Nothing

                Dim sources As List(Of Tuple(Of IGH_ObjectProxy, Integer)) =
                    (From i In Xparam.Elements("source").Take(Max) Select New Tuple(Of IGH_ObjectProxy, Integer)(GetObjectProxy(Guid.Parse(CStr(i.Attribute("id").Value))), CInt(i.Attribute("output").Value))).ToList()
                For Each t As Tuple(Of IGH_ObjectProxy, Integer) In sources
                    If t.Item1 Is Nothing Then Continue For
                    If ids.Contains(t.Item1.Guid) Then Continue For
                    ids.Add(t.Item1.Guid)
                    connected.Add(t)
                Next
                Return connected

            Else
                Dim Xparam As XElement = (From p In XComp.Elements("output") Where CInt(p.Attribute("index").Value) = ParameterIndex Select p).SingleOrDefault()
                If Xparam Is Nothing Then Return Nothing

                Dim targets As List(Of Tuple(Of IGH_ObjectProxy, Integer)) =
                    (From i In Xparam.Elements("target").Take(Max) Select New Tuple(Of IGH_ObjectProxy, Integer)(GetObjectProxy(Guid.Parse(CStr(i.Attribute("id").Value))), CInt(i.Attribute("input").Value))).ToList()
                For Each t As Tuple(Of IGH_ObjectProxy, Integer) In targets
                    If t.Item1 Is Nothing Then Continue For
                    If ids.Contains(t.Item1.Guid) Then Continue For
                    ids.Add(t.Item1.Guid)
                    connected.Add(t)
                Next
                Return connected
            End If

        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(vbCrLf & "AxonXML.GetConnectedComponents() exception at " & Date.Now() & vbCrLf & ex.ToString)
            Return Nothing
        End Try
    End Function

    Public Function GetConnectedRandomComponents(Max As Integer, CompID As Guid, ParameterIndex As Integer, OfInput As Boolean) As List(Of Tuple(Of IGH_ObjectProxy, Integer))
        If Max = Nothing Then Max = Integer.MaxValue
        Try

            Dim XComp As XElement = (From c In XComponentNodes Where Guid.Parse(CStr(c.Attribute("id").Value)).Equals(CompID) Select c).SingleOrDefault()
            If XComp Is Nothing Then Return Nothing
            Dim connected As New List(Of Tuple(Of IGH_ObjectProxy, Integer))
            Dim ids As New List(Of Guid)

            If OfInput Then
                Dim Xparam As XElement = (From p In XComp.Elements("input") Where CInt(p.Attribute("index").Value) = ParameterIndex Select p).SingleOrDefault
                Dim sources As New List(Of Tuple(Of IGH_ObjectProxy, Integer))
                Randomize(Date.Now().Second + Date.Now.Minute)
                Dim top As Integer = Xparam.Elements("source").Count - 1

                For i As Int32 = 0 To top
                    If i >= Max Then Exit For
                    Dim s As XElement = Xparam.Elements("source").ElementAtOrDefault(CInt(Rnd(i) * top))
                    If s Is Nothing Then Continue For
                    sources.Add(New Tuple(Of IGH_ObjectProxy, Integer)(GetObjectProxy(Guid.Parse(CStr(s.Attribute("id").Value))), CInt(s.Attribute("output").Value)))
                Next

                For Each t As Tuple(Of IGH_ObjectProxy, Integer) In sources
                    If t.Item1 Is Nothing Then Continue For
                    If ids.Contains(t.Item1.Guid) Then Continue For
                    ids.Add(t.Item1.Guid)
                    connected.Add(t)
                Next
                Return connected
            Else
                Dim Xparam As XElement = (From p In XComp.Elements("output") Where CInt(p.Attribute("index").Value) = ParameterIndex Select p).SingleOrDefault()
                Dim targets As New List(Of Tuple(Of IGH_ObjectProxy, Integer))
                Randomize(Date.Now().Second + Date.Now.Minute)
                Dim top As Integer = Xparam.Elements("target").Count - 1

                For i As Int32 = 0 To top
                    If i >= Max Then Exit For
                    Dim s As XElement = Xparam.Elements("target").ElementAtOrDefault(CInt(Rnd(i) * top))
                    If s Is Nothing Then Continue For
                    targets.Add(New Tuple(Of IGH_ObjectProxy, Integer)(GetObjectProxy(Guid.Parse(CStr(s.Attribute("id").Value))), CInt(s.Attribute("input").Value)))
                Next
                For Each t As Tuple(Of IGH_ObjectProxy, Integer) In targets
                    If t.Item1 Is Nothing Then Continue For
                    If ids.Contains(t.Item1.Guid) Then Continue For
                    ids.Add(t.Item1.Guid)
                    connected.Add(t)
                Next
                Return connected
            End If
        Catch ex As Exception
            Rhino.RhinoApp.WriteLine(vbCrLf & "AxonXML.GetConnectedRandomComponents() exception at " & Date.Now() & vbCrLf & ex.ToString)
            Return Nothing
        End Try
    End Function

    Public Function GetIndexOfParameter(Component As IGH_DocumentObject, Parameter As IGH_Param, IsInput As Boolean) As Integer

        If TypeOf Component Is IGH_Component Then
            Dim comp As IGH_Component = DirectCast(Component, IGH_Component)
            If IsInput Then
                For i As Int32 = 0 To comp.Params.Input.Count - 1
                    If comp.Params.Input(i).Name.Equals(Parameter.Name) Then
                        Return i
                        Exit For
                    End If
                Next
                Return Nothing
            Else
                For i As Int32 = 0 To comp.Params.Output.Count - 1
                    If comp.Params.Output(i).Name.Equals(Parameter.Name) Then
                        Return i
                        Exit For
                    End If
                Next
                Return Nothing
            End If
        ElseIf TypeOf Component Is IGH_Param
            Return 0
        Else
            Return Nothing
        End If
    End Function

    Public Function GetObjectProxy(ID As Guid) As IGH_ObjectProxy
        Return Server.EmitObjectProxy(ID)
    End Function

    Public Function GetObjectProxy(Name As String) As IGH_ObjectProxy
        Return Server.FindObjectByName(Name, False, False)
    End Function

    Public Overrides Function ToString() As String
        Return XDoc.ToString()
    End Function
#End Region
End Class

Public Class AttributesForm
    Inherits System.Windows.Forms.Form

    Private AxonW As Axon
    Private go As New Boolean

    Sub New(A As Axon)
        AxonW = A
        Me.InitializeComponent()
        NumericUpDownIcons.Value = AxonW.SizeIcon
        NumericUpDownRadius.Value = AxonW.RadiusInner
        NumericUpDownMax.Value = AxonW.Maximun
        go = True
        Me.Show()
    End Sub

    Private Sub NumericUpDownIcons_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownIcons.ValueChanged
        If Not go Then Exit Sub
        AxonW.SizeIcon = NumericUpDownIcons.Value
        AxonW.Owner.Refresh()
    End Sub

    Private Sub NumericUpDownRadius_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownRadius.ValueChanged
        If Not go Then Exit Sub
        AxonW.RadiusInner = NumericUpDownRadius.Value
        AxonW.Owner.Refresh()
    End Sub

    Private Sub NumericUpDownMax_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownMax.ValueChanged
        If Not go Then Exit Sub
        AxonW.Maximun = NumericUpDownMax.Value
        AxonW.Owner.Refresh()
    End Sub

    Private Sub ButtonDefault_MouseDown(sender As Object, e As MouseEventArgs) Handles ButtonDefault.MouseDown
        NumericUpDownRadius.Value = CDec(20)
        NumericUpDownIcons.Value = CDec(24)
        NumericUpDownMax.Value = CDec(43)
        AxonW.SizeIcon = NumericUpDownIcons.Value
        AxonW.RadiusInner = NumericUpDownRadius.Value
        AxonW.Maximun = NumericUpDownMax.Value
        AxonW.Owner.Refresh()
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        AxonW.AttForm = Nothing
    End Sub

#Region "Design"

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.NumericUpDownIcons = New System.Windows.Forms.NumericUpDown()
        Me.LabelIcon = New System.Windows.Forms.Label()
        Me.LabelRadiusInner = New System.Windows.Forms.Label()
        Me.NumericUpDownRadius = New System.Windows.Forms.NumericUpDown()
        Me.LabelMax = New System.Windows.Forms.Label()
        Me.NumericUpDownMax = New System.Windows.Forms.NumericUpDown()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ButtonDefault = New System.Windows.Forms.Button()
        CType(Me.NumericUpDownIcons, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDownRadius, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDownMax, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NumericUpDownIcons
        '
        Me.NumericUpDownIcons.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDownIcons.Location = New System.Drawing.Point(67, 9)
        Me.NumericUpDownIcons.Minimum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.NumericUpDownIcons.Name = "NumericUpDownIcons"
        Me.NumericUpDownIcons.Size = New System.Drawing.Size(67, 20)
        Me.NumericUpDownIcons.TabIndex = 0
        Me.NumericUpDownIcons.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDownIcons.Value = New Decimal(New Integer() {16, 0, 0, 0})
        '
        'LabelIcon
        '
        Me.LabelIcon.AutoSize = True
        Me.LabelIcon.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelIcon.Location = New System.Drawing.Point(12, 9)
        Me.LabelIcon.Name = "LabelIcon"
        Me.LabelIcon.Size = New System.Drawing.Size(42, 15)
        Me.LabelIcon.TabIndex = 1
        Me.LabelIcon.Text = "Icons"
        Me.ToolTip1.SetToolTip(Me.LabelIcon, "Size of icons (in pixels)")
        '
        'LabelRadiusInner
        '
        Me.LabelRadiusInner.AutoSize = True
        Me.LabelRadiusInner.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelRadiusInner.Location = New System.Drawing.Point(12, 35)
        Me.LabelRadiusInner.Name = "LabelRadiusInner"
        Me.LabelRadiusInner.Size = New System.Drawing.Size(49, 15)
        Me.LabelRadiusInner.TabIndex = 3
        Me.LabelRadiusInner.Text = "Radius"
        Me.ToolTip1.SetToolTip(Me.LabelRadiusInner, "Radius inner")
        '
        'NumericUpDownRadius
        '
        Me.NumericUpDownRadius.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDownRadius.Location = New System.Drawing.Point(67, 35)
        Me.NumericUpDownRadius.Name = "NumericUpDownRadius"
        Me.NumericUpDownRadius.Size = New System.Drawing.Size(67, 20)
        Me.NumericUpDownRadius.TabIndex = 2
        Me.NumericUpDownRadius.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.NumericUpDownRadius.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDownRadius.Value = New Decimal(New Integer() {20, 0, 0, 0})
        '
        'LabelMax
        '
        Me.LabelMax.AutoSize = True
        Me.LabelMax.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMax.Location = New System.Drawing.Point(12, 61)
        Me.LabelMax.Name = "LabelMax"
        Me.LabelMax.Size = New System.Drawing.Size(42, 15)
        Me.LabelMax.TabIndex = 5
        Me.LabelMax.Text = "Count"
        Me.ToolTip1.SetToolTip(Me.LabelMax, "Maximum number of components (for ByFrequency and ByDataType)")
        '
        'NumericUpDownMax
        '
        Me.NumericUpDownMax.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDownMax.Location = New System.Drawing.Point(67, 61)
        Me.NumericUpDownMax.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.NumericUpDownMax.Minimum = New Decimal(New Integer() {6, 0, 0, 0})
        Me.NumericUpDownMax.Name = "NumericUpDownMax"
        Me.NumericUpDownMax.Size = New System.Drawing.Size(67, 20)
        Me.NumericUpDownMax.TabIndex = 4
        Me.NumericUpDownMax.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDownMax.Value = New Decimal(New Integer() {25, 0, 0, 0})
        '
        'ButtonDefault
        '
        Me.ButtonDefault.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ButtonDefault.Location = New System.Drawing.Point(0, 90)
        Me.ButtonDefault.Name = "ButtonDefault"
        Me.ButtonDefault.Size = New System.Drawing.Size(146, 23)
        Me.ButtonDefault.TabIndex = 6
        Me.ButtonDefault.Text = "Set default"
        Me.ButtonDefault.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(146, 113)
        Me.Controls.Add(Me.ButtonDefault)
        Me.Controls.Add(Me.LabelMax)
        Me.Controls.Add(Me.NumericUpDownMax)
        Me.Controls.Add(Me.LabelRadiusInner)
        Me.Controls.Add(Me.NumericUpDownRadius)
        Me.Controls.Add(Me.LabelIcon)
        Me.Controls.Add(Me.NumericUpDownIcons)
        Me.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Form1"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Owner = Instances.DocumentEditor
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Axon attributes"
        CType(Me.NumericUpDownIcons, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDownRadius, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDownMax, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents NumericUpDownIcons As NumericUpDown
    Friend WithEvents LabelIcon As Label
    Friend WithEvents LabelRadiusInner As Label
    Friend WithEvents NumericUpDownRadius As NumericUpDown
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents LabelMax As Label
    Friend WithEvents NumericUpDownMax As NumericUpDown
    Friend WithEvents ButtonDefault As Button

#End Region

End Class
