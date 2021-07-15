Imports Excel = Microsoft.Office.Interop.Excel
Public Class ExportData
    Public xExcelapp As New Excel.Application
    Public CatiaFactory As CATIA_Property = New CATIA_Property
    Public MechanismProd As ProductStructureTypeLib.Product
    Public MechanismDoc As INFITF.Document
    Private ExtractParameter() As KnowledgewareTypeLib.Parameter
    Private X_Parameter As KnowledgewareTypeLib.Parameter
    Public X_AxisData(0) As String
    Public Y_AxisData(0, 0) As String
    Sub CATMain()
        MechanismDoc = CatiaFactory.ProductDocument
        MechanismProd = MechanismDoc.Product
        'MsgBox(MechanismDoc.Name)
    End Sub

    Private Sub ExportData_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label3.Text = "● Please Select X_Axis Data."
        Label4.Text = "● Please Select Parameters that you want to apply to chart."
        Label5.Text = "● Please Type Step Value."
        Label6.Text = "● Please Type Total Value."
        Button2.Enabled = False
        CatiaFactory = CATIA_Property.SetInitialCATIA
        Call CATMain()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ReDim ExtractParameter(0)
        Dim sFilter(0)
        sFilter(0) = "Parameter"
        Dim SelectObject As String
        SelectObject = "要加入的參數"
        Dim Instruction As String = "請選擇【" & SelectObject & "】"
        Dim oSel As INFITF.Selection = Me.CatiaFactory.Selection
        Me.CatiaFactory.Selection.Clear()
        Dim Status As String = CATIA_Property.C_SelectMuti(oSel, sFilter, Instruction)

        Dim str1 As String = "● Please Select Parameters that you want to apply to chart."
        If Status <> "Normal" Then
            Me.Label4.Text = str1
            Me.Label4.ForeColor = Color.Black
            Exit Sub
        End If
        Dim selCount As Integer = oSel.Count
        For selCount = 1 To oSel.Count2
            ReDim Preserve ExtractParameter(selCount)
            ExtractParameter(selCount) = oSel.Item2(selCount).Value
        Next
        oSel.Clear()
        If ExtractParameter(1).Name <> "" Then
            Dim str2 As String = Me.Label4.Text.Replace(str1, "Parameters Selected")
            Me.Label4.Text = str2
            Me.Label4.ForeColor = Color.Green
        End If
        Call CheckItem()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If Val(TextBox1.Text) > Val(TextBox2.Text) Then
            MsgBox("step value can not greater than total value!")
            Exit Sub
        End If
        Dim WbName As String = "KinematicsWorkbench"
        Dim cTheMachanisms As INFITF.Workbench = MechanismDoc.GetWorkbench(WbName)
        Dim Mechanisms As KinTypeLib.Mechanisms = cTheMachanisms.Mechanisms
        Dim oFirstMechanism As KinTypeLib.Mechanism = Mechanisms.Item(1)
        Dim iNbProduct As Integer = oFirstMechanism.NbProducts

        Dim oMovingPart As ProductStructureTypeLib.Product = oFirstMechanism.GetProduct(1)
        Dim dValcmd(1)
        Dim dMotion(11)
        oFirstMechanism.GetCommandValues(dValcmd)

        Call SetParameterTitle()
        Dim TotalValue As Double = 7.19 * 25.4
        Dim StepValue = 5
        Do Until dValcmd(0) > (TotalValue) - StepValue
            dValcmd(0) = dValcmd(0) + StepValue
            oFirstMechanism.PutCommandValues(dValcmd)
            'oFirstMechanism.GetProductMotion(oMovingPart, dMotion)
            'oMovingPart.Move.Apply(dMotion)
            oFirstMechanism.Update()
            Call GetParameterData()
        Loop
        If dValcmd(0) <> TotalValue Then
            dValcmd(0) = TotalValue
            oFirstMechanism.PutCommandValues(dValcmd)
            'oFirstMechanism.GetProductMotion(oMovingPart, dMotion)
            'oMovingPart.Move.Apply(dMotion)
            oFirstMechanism.Update()
            Call GetParameterData()
        End If
        'oFirstMechanism.ResetCmdValueToZero(oFirstMechanism.Commands.Item(1))
        Do Until dValcmd(0) < StepValue
            dValcmd(0) = dValcmd(0) - StepValue
            oFirstMechanism.PutCommandValues(dValcmd)
            'oFirstMechanism.GetProductMotion(oMovingPart, dMotion)
            'oMovingPart.Move.Apply(dMotion)
            oFirstMechanism.Update()
            Call GetParameterData()
        Loop
        If dValcmd(0) <> 0 Then
            dValcmd(0) = 0
            oFirstMechanism.PutCommandValues(dValcmd)
            'oFirstMechanism.GetProductMotion(oMovingPart, dMotion)
            'oMovingPart.Move.Apply(dMotion)
            oFirstMechanism.Update()
            Call GetParameterData()
        End If
        Call WriteExcel()



        MsgBox("圖表輸出完成!")
    End Sub
    Sub SetParameterTitle()
        ReDim Y_AxisData(UBound(ExtractParameter), 0)
        Dim Position As Long = InStrRev(X_Parameter.Name, "\")
        Dim str As String = X_Parameter.Name.Remove(0, Position)
        X_AxisData(0) = str
        Dim i As Integer
        For i = 1 To UBound(ExtractParameter)
            Position = InStrRev(ExtractParameter(i).Name, "\")
            Y_AxisData(i, 0) = ExtractParameter(i).Name.Remove(0, Position)
        Next
    End Sub
    Sub GetParameterData()
        ReDim Preserve X_AxisData(UBound(X_AxisData) + 1)
        ReDim Preserve Y_AxisData(UBound(Y_AxisData, 1), UBound(Y_AxisData, 2) + 1)
        X_AxisData(UBound(X_AxisData)) = Val(X_Parameter.ValueAsString)
        Dim i As Integer
        For i = 1 To UBound(ExtractParameter)
            Y_AxisData(i, UBound(Y_AxisData, 2)) = Val(ExtractParameter(i).ValueAsString)
        Next
    End Sub
    Sub WriteExcel()
        Dim Act As Excel.Worksheet
        Dim Wb As Excel.Workbook
        Wb = xExcelapp.Workbooks.Add
        Act = Wb.Sheets(1)
        Dim i As Integer
        For i = 0 To UBound(X_AxisData)
            Act.Cells(i + 1, 1).Value = X_AxisData(i)
        Next
        Dim j As Integer
        For i = 1 To UBound(Y_AxisData, 1)
            For j = 0 To UBound(Y_AxisData, 2)
                Act.Cells(j + 1, i + 1).Value = Y_AxisData(i, j)
            Next
        Next
        Call ExcelEdit(Act)
        Dim FileName As String = MechanismDoc.FullName.Replace(MechanismDoc.Name, "") + "Relation_" + X_AxisData(0) + ".xlsx"
        Act.SaveAs(FileName)
        Wb.Close()
        xExcelapp.Quit()
        Act = Nothing
        Wb = Nothing
        xExcelapp = Nothing

    End Sub
    Sub ExcelEdit(ByRef Act As Excel.Worksheet)
        'Dim Act As Excel.Worksheet
        'Dim Workbooks As Excel.Workbooks
        'Dim Workbook As Excel.Workbook
        Dim i
        Dim Row As Integer
        Dim Columns As Integer
        Dim FileName As String
        'FileName = CatiaFactory.MyCATIA.FileSelectionBox("請選擇已輸出的Excel", "*.xls;*.xlsx", INFITF.CatFileSelectionMode.CatFileSelectionModeOpen)
        'If FileName = "" Then Exit Sub
        ''Open FileName For Input As #1
        'Workbook = xExcelapp.Workbooks.Open(FileName)
        i = 1
        'Act = Workbook.Worksheets(1)
        Do Until Act.Cells(i + 1, 1).value Is Nothing
            If Act.Cells(i, 1).value IsNot Nothing Then
                i = i + 1
            End If
        Loop
        Row = i
        i = 1
        Do Until Act.Cells(1, i + 1).value Is Nothing
            If Act.Cells(1, i).value IsNot Nothing Then
                i = i + 1
            End If
        Loop
        Columns = i
        For j = 2 To Columns
            Call DrawingGraph(Row, j, Act)
        Next
        'Workbook.Save()
        'Workbook.Close()
        'xExcelapp.Quit()
        ''xExcel.Workbooks.Close
        ''Close #1
        'MsgBox("圖表輸出完成!")
    End Sub
    Sub DrawingGraph(ByVal aRow As Integer, ByVal aColumns As Integer, ByRef Act As Excel.Worksheet)
        Dim Xaxis As Excel.Range
        Dim Yaxis As Excel.Range
        xaxis = Act.Range(Act.Cells(2, 1), Act.Cells(aRow, 1))
        yaxis = Act.Range(Act.Cells(2, aColumns), Act.Cells(aRow, aColumns))
        Dim c As Excel.Chart
        c = xExcelapp.ActiveWorkbook.Charts.Add
        c.Name = "Chart" & aColumns - 1
        c = c.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:=Act.Name)
        With c
            .ChartType = Excel.XlChartType.xlLine

            On Error Resume Next
            '    .PlotArea.Interior.ColorIndex = xlNone
            '    If Err.Number <> 0 Then
            '        MsgBox Err.Description
            '    End If
            On Error GoTo 0
            ' set other chart properties
        End With
        'ActiveChart.ChartTitle = "Parameter" & Columns
        Dim xChart As Excel.ChartObject
        xChart = Act.ChartObjects(aColumns - 1)
        Dim ChartCount As Integer
        If aColumns - 1 = 1 Then
            ChartCount = 1
        Else
            ChartCount = (aColumns - 2) * 21 + 1
        End If
        Dim Cell As String
        Cell = "J" & Trim(Str(ChartCount))
        With xChart
            .Height = 340
            .Width = 700
            .Top = Act.Range(Cell).Top
            .Left = Act.Range("J1").Left
        End With
        With xChart.Chart
            .PlotArea.Interior.ColorIndex = 2
        End With
        Dim s As Excel.Series
        Do Until c.SeriesCollection.Count = 0
            c.SeriesCollection(1).Delete()
        Loop
        s = c.SeriesCollection.NewSeries
        With s
            '.Name = ExtractParameter(aColumns - 1).Name
            .ClearFormats()
            .Values = Yaxis
            .XValues = Xaxis
            .MarkerBackgroundColorIndex = Excel.XlColorIndex.xlColorIndexNone
            .MarkerForegroundColor = RGB(256 * Rnd(), 256 * Rnd(), 256 * Rnd())
            .MarkerStyle = Int((3 * Rnd()) + 1)
            .MarkerSize = 4
            ' set other series properties
        End With
        With s.Border
            .Color = RGB(256 * Rnd(), 256 * Rnd(), 256 * Rnd())
            .Weight = Excel.XlBorderWeight.xlMedium
            .LineStyle = Excel.XlLineStyle.xlContinuous

        End With
        With c
            .Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).HasTitle = True
            .Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).AxisTitle.Characters.Text = Act.Range("A1").Value
            .Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasTitle = True
            .Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).AxisTitle.Characters.Text = Act.Cells(1, aColumns).value
            .HasLegend = False

        End With

        xChart.Chart.PlotArea.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone

        xaxis = Nothing
        yaxis = Nothing
        c = Nothing
        s = Nothing

    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        X_Parameter = Nothing
        Dim sFilter(0)
        sFilter(0) = "Parameter"
        Dim SelectObject As String
        SelectObject = "X軸座標參數"
        Dim Instruction As String = "請選擇【" & SelectObject & "】"
        Dim oSel As INFITF.Selection = Me.CatiaFactory.Selection
        Me.CatiaFactory.Selection.Clear()
        Dim Status As String = CATIA_Property.C_Select(oSel, sFilter, Instruction)

        Dim str1 As String = "● Please Select X_Axis Data."
        If Status <> "Normal" Then
            Me.Label3.Text = str1
            Me.Label3.ForeColor = Color.Black
            Exit Sub
        End If
        Dim selCount As Integer = oSel.Count
        X_Parameter = oSel.Item2(1).Value
        oSel.Clear()


        If X_Parameter.Name <> "" Then
            Dim str2 As String = Me.Label3.Text.Replace(str1, "X Axis Data Selected")
            Me.Label3.Text = str2
            Me.Label3.ForeColor = Color.Green
        End If
        Call CheckItem()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress, TextBox2.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim digitsOnly As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("[^\d]")
        TextBox1.Text = digitsOnly.Replace(TextBox1.Text, "")
        Dim str1 As String = "● Please Type Step Value."
        If TextBox1.Text <> "" Then
            Dim str2 As String = Me.Label5.Text.Replace(str1, "Finished input.")
            Me.Label5.Text = str2
            Me.Label5.ForeColor = Color.Green
        Else
            Me.Label5.Text = str1
            Me.Label5.ForeColor = Color.Black
        End If
        Call CheckItem()
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim digitsOnly As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("[^\d]")
        TextBox2.Text = digitsOnly.Replace(TextBox2.Text, "")
        Dim str1 As String = "● Please Type Total Value."
        If TextBox2.Text <> "" Then
            Dim str2 As String = Me.Label6.Text.Replace(str1, "Finished input.")
            Me.Label6.Text = str2
            Me.Label6.ForeColor = Color.Green
        Else
            Me.Label6.Text = str1
            Me.Label6.ForeColor = Color.Black
        End If
        Call CheckItem()
    End Sub
    Sub CheckItem()
        If Label3.ForeColor = Color.Green And Label4.ForeColor = Color.Green And Label5.ForeColor = Color.Green And Label6.ForeColor = Color.Green Then
            Button2.Enabled = True
        Else
            Button2.Enabled = False
        End If
    End Sub
End Class
