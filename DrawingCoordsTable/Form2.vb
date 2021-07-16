Imports System
Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Imports Inventor
Imports System.Threading.Thread
Imports System.Collections.Generic
Imports System.Linq
Imports System.Windows.Forms

Public Class Form2
    Dim _invApp As Inventor.Application
    Dim _started As Boolean = False
    Dim oOrientationne As String

    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheetExcl As Object

    Dim ThreadSec As System.Threading.Thread

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Try
            _invApp = Marshal.GetActiveObject("Inventor.Application")

        Catch ex As Exception
            Try
                Dim invAppType As Type =
                  GetTypeFromProgID("Inventor.Application")

                _invApp = CreateInstance(invAppType)
                _invApp.Visible = True

                'Note: if you shut down the Inventor session that was started
                'this(way) there is still an Inventor.exe running. We will use
                'this Boolean to test whether or not the Inventor App  will
                'need to be shut down.
                _started = True

            Catch ex2 As Exception
                MsgBox(ex2.ToString())
                MsgBox("Unable to get or start Inventor")
            End Try
        End Try
    End Sub


    Dim oDrwDoc As DrawingDocument


    Dim oView As DrawingView
    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Check() Then
            MsgBox("André et Jésus vous remercient! Amen")
        End If
        If _invApp.Documents.Count = 0 Then
            MsgBox("Need to open a drawing document")
            Return

        ElseIf _invApp.ActiveDocument.DocumentType <>
      DocumentTypeEnum.kDrawingDocumentObject Then
            MsgBox("Need to have an Assembly document active")
            Return
        End If

        Dim TypePointString As String
        Dim TypePoint As New List(Of String)
        TypePointString = "oPoint,Point,_Point,_point,point"
        TypePoint = TypePointString.Split(",").ToList()

        oDrwDoc = _invApp.ActiveDocument

        Dim oSheet As Sheet
        oSheet = oDrwDoc.ActiveSheet

        Dim selSet As SelectSet
        selSet = oDrwDoc.SelectSet


        'Prompts the user to select a drawing view
        Dim ViewTrue As Boolean = False
        Dim OriginTrue As Boolean = False
        Dim v As Integer
        For Each ojb As Object In selSet
            v = v + 1
            If selSet(v).Type = ObjectTypeEnum.kDrawingViewObject Then
                ViewTrue = True
            ElseIf selSet(v).Type = ObjectTypeEnum.kOriginIndicatorObject Then
                OriginTrue = True
            End If
        Next

        Try
            Dim TryGo1 As Integer
            Dim PromptView As Integer

            If ViewTrue = False Then
                PromptView = MsgBox("Select a view for your coordinate system", MsgBoxStyle.OkOnly)
                If PromptView <> 1 Then
                    Return
                Else
                    If selSet.Count <> 1 And TryGo1 = 0 Then
                        MsgBox("Please select only one view")
                        TryGo1 = TryGo1 + 1
                        Return
                    End If
                End If
            End If
            Try
                Dim i As Integer
                For Each obj As Object In selSet
                    i = i + 1
                    If selSet(i).Type = ObjectTypeEnum.kDrawingViewObject Then 'Doit trouver un alternative à RETURN parce que ça suce
                        oView = selSet(i)
                        Exit For
                    Else
                    End If
                Next

            Catch ex As Exception
                MsgBox("We couldn't get the type of the selected object, or its somehow not a DrawingView")
                Return
            End Try

        Catch ex As Exception
            MsgBox("You need to select something")
            Return
        End Try


        'Prompts The User to Select an Origin Indicator
        Dim oOrigin As OriginIndicator
        Dim PromptOrigin As Integer

        'Checks to see if drawing has Origin Indicator
        Try
            If Not oView.HasOriginIndicator Then 'If not displays message
                MsgBox("Please add an Origin Indicator to your drawings (Annotate\Hole)")
                Return
            End If
            'Prompts the user to select an Origin indicator
            If OriginTrue = False Or ViewTrue = False Then
                PromptOrigin = MsgBox("Select the origin indicator for your coordinate system, then press ok", MsgBoxStyle.MsgBoxSetForeground)
                If PromptOrigin <> 1 Then
                    Return
                End If

                If selSet.Count = 0 And PromptOrigin <> 1 Then
                    MsgBox("Need to choose/select the origin")
                    Return
                End If
            End If
            Try
                Dim i As Integer
                For Each obj As Object In selSet
                    i = i + 1
                    If selSet(i).Type = ObjectTypeEnum.kOriginIndicatorObject Then
                        oOrigin = selSet(i)
                    End If
                Next

            Catch ex As Exception
                MsgBox("Unable to get an origin indicator from your selection")
                Return
            End Try

            If Not oOrigin.Type = ObjectTypeEnum.kOriginIndicatorObject Then
                Debug.Print(oOrigin.Type.ToString)
                MsgBox("Please choose an origin indicator")
                Return
            End If
        Catch ex As Exception
            MsgBox("You need to select something")
            Return
        End Try

        'At this point the origin is set
        'The user will have to mark the points he wants to get the coords of
        'with a sketch symbol, it could be anything, like, anything
        Dim oPoints As New List(Of SketchedSymbol)
        Dim oSketchedSymbol As SketchedSymbol
        Dim oSketchedSymbols As SketchedSymbols
        oSketchedSymbols = oSheet.SketchedSymbols
        Dim SketchedSymbolPrompt As MsgBoxResult
        Dim AreThereAny As Boolean

        For Each oSketchedSymbol In oSketchedSymbols
            If Not TypePoint.Contains(oSketchedSymbol.Name) Then
                AreThereAny = True
            End If
        Next

        If AreThereAny = True Then
            SketchedSymbolPrompt = MsgBox("We detected that you had some points with no name referring to some points, are you sure ALL of your Sketched Symbols are POINTS?", MsgBoxStyle.YesNo)
            If SketchedSymbolPrompt = 7 Then
                MsgBox("Please select your coordinate points, then click ok")

                Dim i As Integer
                i = 1
                For Each oSketchedSymbol In selSet
                    oPoints.Add(selSet(i))
                    i = i + 1
                Next

            ElseIf SketchedSymbolPrompt <> 6 Then
                Return
            Else
                Dim i As Integer
                i = 1
                For Each oSketchedSymbol In oSketchedSymbols
                    oPoints.Add(oSketchedSymbols(i))
                    i = i + 1
                Next
            End If
        Else
            Dim i As Integer
            i = 1
            For Each oSketchedSymbol In oSketchedSymbols
                oPoints.Add(oSketchedSymbols(i))
                i = i + 1
            Next
        End If


        'We now have a list of all the points the user defined, wich includes all the X Y coords
        'They most probably are in the right order
        'We now have to create a table and add these coords to each one of their respective row and column
        'After we have to add baloons foreach named "AXX"

        Dim oTitles() As String
        ReDim oTitles(1)
        oTitles(0) = "X"
        oTitles(1) = "Y"



        Dim oColumn As Integer
        Dim oRow As Integer
        Dim oContent As Array

        'Add headers to the worksheet on row 1

        ThreadSec = New System.Threading.Thread(Sub(t)
                                                    'Transfer the array to the worksheet starting at cell A2
                                                    Debug.Print("Thread2 NOW ACTIVE")

                                                    'Check si excel is open avec
                                                    If Process.GetProcessesByName("Excel.Application").Count > 0 Then
                                                    End If
                                                    oExcel = CreateObject("Excel.Application")
                                                    'Start a new workbook in Excel

                                                    oExcel.Visible = False
                                                    oBook = oExcel.Workbooks.Add
                                                    oSheetExcl = oBook.Worksheets(1)
                                                    Call CheckOrientation()

                                                    Dim oOr As String
                                                    If oOrientationne = "Vertical" Then
                                                        oOr = "A1:C1"
                                                        Dim array(0 To 2) As String
                                                        array(0) = "Point"
                                                        array(1) = "X"
                                                        array(2) = "Y"

                                                        oSheetExcl.Range(oOr).Value = array
                                                    ElseIf oOrientationne = "Horizontal" Then
                                                        oOr = "A1:A3"
                                                        Dim array(0 To 2, 0) As String
                                                        array(0, 0) = "Point"
                                                        array(1, 0) = "X"
                                                        array(2, 0) = "Y"

                                                        oSheetExcl.Range(oOr).Value = array
                                                    End If

                                                    If oOrientationne = "Vertical" Then
                                                        oSheetExcl.Range("A2").Resize(oPoints.Count, 3).Value = Orientation(oOrientationne, oPoints, oOrigin, oContent, oColumn, oRow)
                                                    ElseIf oOrientationne = "Horizontal" Then
                                                        oSheetExcl.Range("B1").Resize(3, oPoints.Count).Value = Orientation(oOrientationne, oPoints, oOrigin, oContent, oColumn, oRow)
                                                    End If
                                                    'Save the Workbook and Quit Excel
                                                    Dim SavePath As String = ExcelFilePath()
                                                    Dim ExcelSavePAth As String
                                                    My.Computer.FileSystem.CreateDirectory(SavePath)
                                                    ExcelSavePAth = String.Concat(SavePath, excl)
                                                    oBook.SaveAs(ExcelSavePAth)

                                                    Dim InsP As Point2d
                                                    InsP = _invApp.TransientGeometry.CreatePoint2d(15, 15)

                                                    Dim oCustomTable As CustomTable
                                                    MsgBox("about TypeOf create table")
                                                    oCustomTable = oSheet.CustomTables.AddExcelTable(ExcelSavePAth, InsP, String.Concat("Coordonnée", " ", "(", DomainUpDown1.Text, ")"))

                                                    oExcel.Visible = False
                                                    oExcel.Quit
                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)

                                                End Sub)
        ThreadSec.Start()





        'oCustomTable = oSheet.CustomTables.Add("MAXI", InsP, 2, 32, oTitles, oContent, 3)

        'oCustomTable = oSheet.CustomTables.Add("Coordonnée", InsP, oColumn, oRow, oTitles, oContent)


        'We need to now add balloons... TD 6
        Dim AverageX As Double
        Dim AverageY As Double
        AverageX = oView.Width / 2
        AverageY = oView.Height / 2

        Dim MoinOrigine As Point2d
        MoinOrigine = oOrigin.Intent.PointOnSheet

        GenNotes = oSheet.DrawingNotes.GeneralNotes

        Dim OffX As Double
        Dim OffY As Double

        Dim b As Integer

        For Each ThisPoint In oPoints
            Dim oLeaderPoints As ObjectCollection
            oLeaderPoints = _invApp.TransientObjects.CreateObjectCollection

            b = b + 1

            If ThisPoint.Position.X - MoinOrigine.X >= AverageX Then
                OffX = 1
            ElseIf ThisPoint.Position.X - MoinOrigine.X < AverageX Then
                OffX = -1
            End If

            If ThisPoint.Position.Y - MoinOrigine.Y + 1 >= AverageY Then
                OffY = 0.25
            ElseIf ThisPoint.Position.Y - MoinOrigine.Y - 1 < AverageY Then
                OffY = -0.25
            End If

            Dim pos As Point2d = _invApp.TransientGeometry.CreatePoint2d(ThisPoint.Position.X + OffX, ThisPoint.Position.Y + OffY)
            'oLeaderPoints.Add(_invApp.TransientGeometry.CreatePoint2d(60, 45))
            'oLeaderPoints.Add(ThisPoint.Position)

            'Dim GeomIntentText As GeometryIntent =
            'oSheet.CreateGeometryIntent(ThisPoint)

            'oLeaderPoints.Add(GeomIntentText)

            Dim f As String
            f = String.Concat("A", b.ToString)
            Dim AlreadyGenNote As Boolean
            AlreadyGenNote = TextNoteAlreadyExists(f)


            If Not AlreadyGenNote Then
                oSheet.DrawingNotes.GeneralNotes.AddFitted(pos, f)
            End If

            'Dim OLeaderNote As LeaderNote =
            'oSheet.DrawingNotes.LeaderNotes.Add(oLeaderPoints, "aiii")

            'Dim oFirstNode As LeaderNode =
            'OLeaderNote.Leader.RootNode.ChildNodes.Item(1)
            'Dim oSeconNode As LeaderNode =
            'oFirstNode.ChildNodes.Item(1)

            'oFirstNode.InsertNode(oSeconNode, _invApp.TransientGeometry.CreatePoint2d(ThisPoint.Position.X, ThisPoint.Position.Y))


        Next
    End Sub

    Function Orientation(ByVal oOrientation As String, ByVal oPoints As List(Of SketchedSymbol), ByVal oOrigin As OriginIndicator, ByRef oContente() As Double, ByRef Column As Double, ByRef row As Double, Optional oOptional As Boolean = False) As Array
        Dim ThisPoint As SketchedSymbol
        Dim MoinOrigine As Point2d
        MoinOrigine = oOrigin.Intent.PointOnSheet

        Dim nb As Integer = oPoints.Count
        Dim r As Integer = -1
        If oOrientation = "Vertical" Then

            Dim oContent(0 To nb, 0 To 2)
            For Each ThisPoint In oPoints
                r = r + 1
                oContent(r, 0) = "A" & (r + 1).ToString
                oContent(r, 1) = Round((ThisPoint.Position.X - MoinOrigine.X) / oView.Scale * UnitMod)
                oContent(r, 2) = Round((ThisPoint.Position.Y - MoinOrigine.Y) / oView.Scale * UnitMod)
            Next
            Return oContent

        ElseIf oOrientation = "Horizontal" Then

            Dim oContent(0 To 2, 0 To nb - 1)
            For Each ThisPoint In oPoints
                r = r + 1
                oContent(0, r) = "A" & (r + 1).ToString
                oContent(1, r) = Round((ThisPoint.Position.X - MoinOrigine.X) / oView.Scale * UnitMod)
                oContent(2, r) = Round((ThisPoint.Position.Y - MoinOrigine.Y) / oView.Scale * UnitMod)
            Next
            Return oContent
        End If
    End Function
    Public Function CheckOrientation()
        If oOrientationne = "Automatic" Then
            If 1.3 * oView.Width >= oView.Height Then
                oOrientationne = "Horizontal"
            Else
                oOrientationne = "Vertical"
            End If
        End If
    End Function
    Private Sub PictureBox2_Click_1(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Process.Start("Chrome", "https://github.com/migui06/DrawingsCoords-PalardyAdd-ins")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        oOrientationne = ComboBox1.Text
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Form2.ActiveForm.Close()
    End Sub
#Region " Move Form "

    ' [ Move Form ]
    '
    ' // By Elektro 

    Public MoveForm As Boolean
    Public MoveForm_MousePosition As System.Drawing.Point

    Public Sub MoveForm_MouseDown(sender As Object, e As MouseEventArgs) Handles _
    MyBase.MouseDown ' Add more handles here (Example: PictureBox1.MouseDown)

        If e.Button = MouseButtons.Left Then
            MoveForm = True
            Me.Cursor = Cursors.NoMove2D
            MoveForm_MousePosition = e.Location
        End If

    End Sub

    Public Sub MoveForm_MouseMove(sender As Object, e As MouseEventArgs) Handles _
    MyBase.MouseMove ' Add more handles here (Example: PictureBox1.MouseMove)

        If MoveForm Then
            Me.Location = Me.Location + (e.Location - MoveForm_MousePosition)
        End If

    End Sub

    Public Sub MoveForm_MouseUp(sender As Object, e As MouseEventArgs) Handles _
    MyBase.MouseUp ' Add more handles here (Example: PictureBox1.MouseUp)

        If e.Button = MouseButtons.Left Then
            MoveForm = False
            Me.Cursor = Cursors.Default
        End If

    End Sub
#End Region
    Dim Path As String
    Dim ExtraDirectory As String
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = 1 Then
            TextBox1.Visible = True
            Label4.Visible = True
            TextBox1.Text = "ExcelCoords"
            TextBox1.SelectAll()
        Else
            TextBox1.Visible = False
            Label4.Visible = False
            TextBox1.Text = ""
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ExtraDirectory = String.Concat("\", TextBox1.Text)
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Visible = False
        Label4.Visible = False
        DomainUpDown1.SelectedIndex = 0
        DomainUpDown2.SelectedIndex = 2
        ComboBox1.SelectedIndex = 2
        ComboBox2.SelectedIndex = 0
        TrucsImportants.Add("https://www.zalkincapping.com/")

        'If i >= 15 Then
        '    My.Computer.Audio.Play(My.Resources.bomb, AudioPlayMode.WaitToComplete)
        '    Process.Start("chrome", "https://www.youtube.com/watch?v=dQw4w9WgXcQ")
        'End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Path = ComboBox2.Text
    End Sub
    Dim TrucsImportants As New List(Of String)
    Dim BoolTrucsImportants As Boolean

    Dim excl As String = "\Coords1.xlsx"
    Public Function ExcelFilePath() As String
        Dim actualpath() As String
        Dim actualpath2() As String
        Dim NewPath As String
        If Path = "Same as drawing" Then
            actualpath = Split(oDrwDoc.File.FullFileName, "\")

            actualpath2 = RemoveAt(actualpath, actualpath.Count - 1)
            NewPath = String.Join("\", actualpath2)
            NewPath = String.Concat(NewPath, ExtraDirectory)
            Return NewPath
        ElseIf Path = "Same as part" Then
        ElseIf Path Is Nothing Then
            Return "%Appdata%\Local\André Zalkin\Book1.xlsx"
        End If
        '*User Defined Path
        '*eg. C:\Program Files\ZalkinCappin\Capping\Utilisateur\André Zalkin\Fichier sensible
    End Function
    Public Function Check()
        If Not ComboBox2.Items.Contains(oOrientationne) Then
        End If


    End Function
    Function RemoveAt(Of T)(ByVal arr As T(), ByVal index As Integer) As T()
        Dim uBound = arr.GetUpperBound(0)
        Dim lBound = arr.GetLowerBound(0)
        Dim arrLen = uBound - lBound

        If index < lBound OrElse index > uBound Then
            Throw New ArgumentOutOfRangeException(
            String.Format("Index must be from {0} to {1}.", lBound, uBound))

        Else
            'create an array 1 element less than the input array
            Dim outArr(arrLen - 1) As T
            'copy the first part of the input array
            Array.Copy(arr, 0, outArr, 0, index)
            'then copy the second part of the input array
            Array.Copy(arr, index + 1, outArr, index, uBound - index)

            Return outArr
        End If
    End Function

    Dim UnitMod As Double
    Private Sub DomainUpDown1_SelectedItemChanged(sender As Object, e As EventArgs) Handles DomainUpDown1.SelectedItemChanged
        Dim Unit As String = DomainUpDown1.Text
        If Unit = "in" Then
            UnitMod = 1 * 10 / 25.4
        ElseIf Unit = "ft" Then
            UnitMod = 1 * 10 / 25.4 / 12
        ElseIf Unit = "mm" Then
            UnitMod = 1 * 10
        ElseIf Unit = "cm" Then
            UnitMod = 1
        ElseIf Unit = "m" Then
            UnitMod = 1 / 100
        ElseIf Unit = "km" Then
            UnitMod = 1 / 100 / 1000
        ElseIf Unit = "au" Then
            UnitMod = 1 / 100 / 149597870700
        ElseIf Unit = "ly" Then
            UnitMod = 1 / 100 / 9460730472580800
        End If
    End Sub
    Dim PrecisionMod As String
    Private Sub DomainUpDown2_SelectedItemChanged(sender As Object, e As EventArgs) Handles DomainUpDown2.SelectedItemChanged
        Dim Precision As String = DomainUpDown2.Text
        If Precision = "More?" Then
            PrecisionMod = "Don't"
        ElseIf Precision = "???" Then
            PrecisionMod = "???"
            oOrientationne = "Vertical"
        Else
            PrecisionMod = Precision
        End If
    End Sub
    Dim QuantitéDieu As Integer
    Dim Dieu() As String = Split((My.Resources.ParoleAndrer), ".")
    Public Function Round(Coord As Double)
        If PrecisionMod = "Don't" Then
            Return Coord
        ElseIf PrecisionMod = "???" Then
            QuantitéDieu = QuantitéDieu + 1
            Try
                Return Dieu(QuantitéDieu - 1)
            Catch ex As Exception
                Return "Andrer"
            End Try
        ElseIf PrecisionMod Then
            Return Math.Round(Coord, PrecisionMod.Count - 2)
        End If
    End Function
    Dim GenNotes As GeneralNotes
    Public Function TextNoteAlreadyExists(f As String) As Boolean
        Dim i As Integer
        For Each obj As Object In GenNotes
            i = i + 1
            Try
                If GenNotes(i).Text.Contains(f) Then
                    Return True
                End If
            Catch ex As Exception

            End Try

        Next
    End Function
End Class