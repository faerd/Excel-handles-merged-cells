' Excel 辅助工具宏代码 - 优化修正版
' 版本：2.3

' ==================== 撤销管理模块（支持3级撤销） ====================

Private Const MAX_UNDO_LEVELS As Integer = 3  ' 支持3级撤销

Private Type UndoRecord
    TargetAddress As String
    OldValues As Variant
    OldFormats As Variant
    OldMergeInfo As String
    OldRowHeights As Variant
    OldColumnWidths As Variant
    wasMerged As Boolean
    undoType As String  ' 记录操作类型
End Type

Private UndoStack(1 To 3) As UndoRecord
Private UndoStackPointer As Integer
Private IsUndoing As Boolean

Private Sub InitializeUndoStack()
    UndoStackPointer = 0
    IsUndoing = False
    Dim i As Integer
    For i = 1 To MAX_UNDO_LEVELS
        UndoStack(i).TargetAddress = ""
        UndoStack(i).wasMerged = False
        UndoStack(i).undoType = ""
    Next i
End Sub

Public Sub PushUndo(rng As Range, Optional undoType As String = "General")
    On Error Resume Next

    If UndoStackPointer = 0 Then
        InitializeUndoStack
    End If

    ' 如果栈已满，移除最旧的记录
    If UndoStackPointer >= MAX_UNDO_LEVELS Then
        Dim i As Integer
        For i = 1 To MAX_UNDO_LEVELS - 1
            UndoStack(i) = UndoStack(i + 1)
        Next i
        UndoStackPointer = MAX_UNDO_LEVELS - 1
    End If

    UndoStackPointer = UndoStackPointer + 1

    With UndoStack(UndoStackPointer)
        .TargetAddress = rng.Address
        .wasMerged = False
        .OldMergeInfo = ""
        .undoType = undoType

        ' 保存值和格式
        ReDim .OldValues(1 To rng.Cells.count)
        ReDim .OldFormats(1 To rng.Cells.count)

        ' 保存行高（如果是多行选择）
        If rng.rows.count > 1 Then
            ReDim .OldRowHeights(1 To rng.rows.count)
            Dim r As Long
            For r = 1 To rng.rows.count
                .OldRowHeights(r) = rng.rows(r).RowHeight
            Next r
        End If

        ' 保存列宽（如果是多列选择）
        If rng.Columns.count > 1 And rng.Columns.count < 100 Then  ' 限制避免太大
            ReDim .OldColumnWidths(1 To rng.Columns.count)
            Dim c As Long
            For c = 1 To rng.Columns.count
                .OldColumnWidths(c) = rng.Columns(c).ColumnWidth
            Next c
        End If

        Dim j As Long
        Dim cell As Range
        Dim mergeInfo As String
        mergeInfo = ""

        For j = 1 To rng.Cells.count
            Set cell = rng.Cells(j)
            .OldValues(j) = cell.Value
            .OldFormats(j) = cell.NumberFormat

            If cell.MergeCells Then
                .wasMerged = True
                mergeInfo = mergeInfo & j & ":" & cell.mergeArea.Address & ";"
            End If
        Next j

        .OldMergeInfo = mergeInfo
    End With

    On Error GoTo 0
End Sub

Public Sub PopUndo()
    If UndoStackPointer <= 0 Then
        Exit Sub
    End If

    IsUndoing = True
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    With UndoStack(UndoStackPointer)
        Dim rng As Range
        Set rng = Range(.TargetAddress)

        ' 取消当前合并
        Dim cell As Range
        For Each cell In rng.Cells
            If cell.MergeCells Then
                cell.mergeArea.UnMerge
            End If
        Next cell

        ' 恢复值和格式
        Dim i As Long
        For i = 1 To rng.Cells.count
            rng.Cells(i).Value = .OldValues(i)
            rng.Cells(i).NumberFormat = .OldFormats(i)
        Next i

        ' 恢复合并状态
        If .wasMerged And .OldMergeInfo <> "" Then
            Dim mergeParts() As String
            mergeParts = Split(.OldMergeInfo, ";")

            Dim part As Variant
            For Each part In mergeParts
                If part <> "" Then
                    Dim colonPos As Long
                    colonPos = InStr(part, ":")
                    If colonPos > 0 Then
                        Dim areaAddr As String
                        areaAddr = Mid(part, colonPos + 1)
                        On Error Resume Next
                        Range(areaAddr).Merge
                        On Error GoTo ErrorHandler
                    End If
                End If
            Next part
        End If

        ' 恢复行高
        If Not IsEmpty(.OldRowHeights) Then
            For r = 1 To UBound(.OldRowHeights)
                If r <= rng.rows.count Then
                    rng.rows(r).RowHeight = .OldRowHeights(r)
                End If
            Next r
        End If

        ' 恢复列宽
        If Not IsEmpty(.OldColumnWidths) Then
            For c = 1 To UBound(.OldColumnWidths)
                If c <= rng.Columns.count Then
                    rng.Columns(c).ColumnWidth = .OldColumnWidths(c)
                End If
            Next c
        End If
    End With

    UndoStackPointer = UndoStackPointer - 1

    GoTo CleanUp

ErrorHandler:
    ' 静默处理错误

CleanUp:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    IsUndoing = False
End Sub

' ==================== 功能1: 合并单元格并保留首个内容 ====================
Public Sub MergeKeepFirst()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Dim rng As Range
    Set rng = Selection

    If rng.Cells.count = 1 Then Exit Sub

    PushUndo rng, "MergeKeepFirst"

    Dim firstValue As Variant
    firstValue = rng.Cells(1).Value

    rng.Merge
    rng.Value = firstValue

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.ScreenUpdating = True
End Sub

' ==================== 功能2: 向下批量合并单元格 ====================
Public Sub BatchMergeDown()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim rng As Range
    Set rng = Selection

    If rng.Cells.count = 1 Then GoTo CleanUp

    PushUndo rng, "BatchMergeDown"

    Dim col As Range
    For Each col In rng.Columns
        ProcessColumnMerge col
    Next col

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Sub ProcessColumnMerge(col As Range)
    Dim cellsInCol As Range
    Set cellsInCol = col.Cells

    If cellsInCol.count = 1 Then Exit Sub

    Dim firstNonEmpty As Long
    firstNonEmpty = 1

    Do While firstNonEmpty <= cellsInCol.count
        If IsEmpty(cellsInCol(firstNonEmpty)) Or Trim(cellsInCol(firstNonEmpty).Value) = "" Then
            firstNonEmpty = firstNonEmpty + 1
        Else
            Exit Do
        End If
    Loop

    If firstNonEmpty > 1 Then
        If firstNonEmpty > 2 Then
            Range(cellsInCol(1), cellsInCol(firstNonEmpty - 1)).Merge
        End If
    End If

    If firstNonEmpty > cellsInCol.count Then
        cellsInCol.Merge
        Exit Sub
    End If

    Dim currentPos As Long
    currentPos = firstNonEmpty

    Do While currentPos <= cellsInCol.count
        Dim currentCell As Range
        Set currentCell = cellsInCol(currentPos)
        Dim currentValue As String
        currentValue = CStr(currentCell.Value)

        Dim mergeEnd As Long
        mergeEnd = currentPos

        Dim checkPos As Long
        For checkPos = currentPos + 1 To cellsInCol.count
            Dim checkValue As String
            checkValue = CStr(cellsInCol(checkPos).Value)

            If Trim(checkValue) = "" Or checkValue = currentValue Then
                mergeEnd = checkPos
            Else
                Exit For
            End If
        Next checkPos

        If mergeEnd > currentPos Then
            Range(currentCell, cellsInCol(mergeEnd)).Merge
            currentCell.Value = currentValue
        End If

        currentPos = mergeEnd + 1
    Loop
End Sub

' ==================== 功能3: 合并单元格并保留所有内容 ====================
Public Sub MergeKeepAll()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Dim rng As Range
    Set rng = Selection

    If rng.Cells.count = 1 Then Exit Sub

    PushUndo rng, "MergeKeepAll"

    Dim allContent As String
    Dim cell As Range
    Dim first As Boolean
    first = True

    For Each cell In rng.Cells
        If Not IsEmpty(cell) And Trim(cell.Value) <> "" Then
            If Not first Then
                allContent = allContent & vbLf
            End If
            allContent = allContent & Trim(cell.Value)
            first = False
        End If
    Next cell

    rng.Merge
    rng.Value = allContent
    rng.WrapText = True

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.ScreenUpdating = True
End Sub

' ==================== 功能4: 向下批量合并空白单元格 ====================
Public Sub BatchMergeBlanks()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim rng As Range
    Set rng = Selection

    PushUndo rng, "BatchMergeBlanks"

    Dim col As Range
    For Each col In rng.Columns
        ProcessColumnMergeBlanks col
    Next col

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Sub ProcessColumnMergeBlanks(col As Range)
    Dim cellsInCol As Range
    Set cellsInCol = col.Cells

    If cellsInCol.count = 1 Then Exit Sub

    Dim i As Long
    i = 1

    Do While i <= cellsInCol.count
        If IsEmpty(cellsInCol(i)) Or Trim(cellsInCol(i).Value) = "" Then
            i = i + 1
        Else
            Dim contentCell As Range
            Set contentCell = cellsInCol(i)
            Dim contentValue As String
            contentValue = CStr(contentCell.Value)

            Dim endRow As Long
            endRow = i

            Dim j As Long
            For j = i + 1 To cellsInCol.count
                If IsEmpty(cellsInCol(j)) Or Trim(cellsInCol(j).Value) = "" Then
                    endRow = j
                Else
                    Exit For
                End If
            Next j

            If endRow > i Then
                Range(contentCell, cellsInCol(endRow)).Merge
                contentCell.Value = contentValue
            End If

            i = endRow + 1
        End If
    Loop
End Sub

' ==================== 功能5: 合并单元格并填充内容 ====================
Public Sub MergeAndFill()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Dim rng As Range
    Set rng = Selection

    PushUndo rng, "MergeAndFill"

    Dim allContent As String
    Dim rw As Range
    Dim cell As Range
    Dim first As Boolean
    first = True

    For Each rw In rng.rows
        For Each cell In rw.Cells
            If Not IsEmpty(cell) And Trim(cell.Value) <> "" Then
                If Not first Then
                    allContent = allContent & vbLf
                End If
                allContent = allContent & Trim(cell.Value)
                first = False
            End If
        Next cell
    Next rw

    rng.Merge
    rng.Value = allContent
    rng.WrapText = True

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.ScreenUpdating = True
End Sub

' ==================== 功能6: 拆分单元格内容（修复撤销） ====================
Public Sub SplitCellContent()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim rng As Range
    Set rng = Selection

    ' 计算需要的最大行数
    Dim maxParaCount As Long
    maxParaCount = 0

    Dim checkCell As Range
    For Each checkCell In rng.Cells
        If Not IsEmpty(checkCell) Then
            Dim contentCheck As String
            contentCheck = Replace(Replace(CStr(checkCell.Value), vbCrLf, vbLf), vbCr, vbLf)
            Dim parasCheck() As String
            parasCheck = Split(contentCheck, vbLf)

            Dim nonEmptyCount As Long
            nonEmptyCount = 0
            Dim pc As Long
            For pc = 0 To UBound(parasCheck)
                If Trim(parasCheck(pc)) <> "" Then nonEmptyCount = nonEmptyCount + 1
            Next pc

            If nonEmptyCount > maxParaCount Then maxParaCount = nonEmptyCount
        End If
    Next checkCell

    ' 扩大撤销范围以包含可能插入的行
    Dim undoRange As Range
    If maxParaCount > rng.rows.count Then
        Set undoRange = rng.Resize(maxParaCount + 2, rng.Columns.count)  ' 多预留2行
    Else
        Set undoRange = rng
    End If

    PushUndo undoRange, "SplitCellContent"

    Dim targetCell As Range
    For Each targetCell In rng.Cells
        ' 处理合并单元格
        If targetCell.MergeCells Then
            Dim mergeArea As Range
            Set mergeArea = targetCell.mergeArea
            mergeArea.UnMerge
            Set targetCell = mergeArea.Cells(1)
        End If

        Dim content As String
        content = Trim(CStr(targetCell.Value))

        If content = "" Then GoTo NextCell

        ' 统一换行符
        content = Replace(content, vbCrLf, vbLf)
        content = Replace(content, vbCr, vbLf)

        Dim paragraphs() As String
        paragraphs = Split(content, vbLf)

        ' 收集非空段落
        Dim paraList() As String
        ReDim paraList(0 To UBound(paragraphs))
        Dim paraCount As Long
        paraCount = 0

        Dim i As Long
        For i = 0 To UBound(paragraphs)
            If Trim(paragraphs(i)) <> "" Then
                paraList(paraCount) = Trim(paragraphs(i))
                paraCount = paraCount + 1
            End If
        Next i

        If paraCount <= 1 Then GoTo NextCell

        ReDim Preserve paraList(0 To paraCount - 1)

        ' 检查是否需要插入行
        Dim neededRows As Long
        neededRows = paraCount

        If neededRows > 1 Then
            Dim emptyRowsBelow As Long
            emptyRowsBelow = 0

            Dim checkRow As Long
            For checkRow = 1 To neededRows - 1
                If targetCell.row + checkRow <= rows.count Then
                    If Application.WorksheetFunction.CountA(rows(targetCell.row + checkRow)) = 0 Then
                        emptyRowsBelow = emptyRowsBelow + 1
                    Else
                        Exit For
                    End If
                End If
            Next checkRow

            Dim rowsToInsert As Long
            rowsToInsert = (neededRows - 1) - emptyRowsBelow

            If rowsToInsert > 0 Then
                targetCell.Offset(1, 0).Resize(rowsToInsert).EntireRow.Insert
            End If
        End If

        ' 填充内容
        For i = 0 To paraCount - 1
            targetCell.Offset(i, 0).Value = paraList(i)
        Next i

        ' 清空多余单元格
        For i = paraCount To neededRows - 1
            targetCell.Offset(i, 0).ClearContents
        Next i

NextCell:
    Next targetCell

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ==================== 功能7: 取消合并并自动填充 ====================
Public Sub UnMergeAndFill()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim rng As Range
    Set rng = Selection

    PushUndo rng, "UnMergeAndFill"

    Dim cell As Range
    Dim unmergedRange As Range

    For Each cell In rng.Cells
        If cell.MergeCells Then
            Dim mergeRange As Range
            Set mergeRange = cell.mergeArea
            Dim mergeValue As Variant
            mergeValue = cell.Value

            mergeRange.UnMerge

            Dim fillCell As Range
            For Each fillCell In mergeRange
                fillCell.Value = mergeValue
            Next fillCell

            If unmergedRange Is Nothing Then
                Set unmergedRange = mergeRange
            Else
                Set unmergedRange = Union(unmergedRange, mergeRange)
            End If
        End If
    Next cell

    If Not unmergedRange Is Nothing Then
        FillDownBlanksCore unmergedRange, False
    Else
        FillDownBlanksCore rng, False
    End If

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ==================== 功能8: 自动向下填充 ====================
Public Sub FillDownBlanks()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim rng As Range
    Set rng = Selection

    PushUndo rng, "FillDownBlanks"

    FillDownBlanksCore rng, False

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Sub FillDownBlanksCore(rng As Range, showMsg As Boolean)
    Dim col As Range
    For Each col In rng.Columns
        Dim lastValue As Variant
        Dim hasValue As Boolean
        hasValue = False

        Dim cell As Range
        For Each cell In col.Cells
            If IsEmpty(cell) Or Trim(cell.Value) = "" Then
                If hasValue Then
                    cell.Value = lastValue
                End If
            Else
                lastValue = cell.Value
                hasValue = True
            End If
        Next cell
    Next col
End Sub

' ==================== 功能9: 单元格内容合并 ====================
Public Sub MergeCellContents()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Dim rng As Range
    Set rng = Selection

    If rng.Cells.count = 1 Then Exit Sub

    PushUndo rng, "MergeCellContents"

    Dim firstCell As Range
    Set firstCell = rng.Cells(1)

    Dim allContent As String
    Dim rw As Range
    Dim cell As Range
    Dim first As Boolean
    first = True

    For Each rw In rng.rows
        For Each cell In rw.Cells
            If Not IsEmpty(cell) And Trim(cell.Value) <> "" Then
                If Not first Then
                    allContent = allContent & vbLf
                End If
                allContent = allContent & Trim(cell.Value)
                first = False
            End If
        Next cell
    Next rw
    ' 保存第一个单元格的合并状态
    Dim wasMerged As Boolean
    Dim mergeAddr As String
    wasMerged = firstCell.MergeCells

    If wasMerged Then
        mergeAddr = firstCell.mergeArea.Address
    End If

    ' 清空所有单元格
    rng.ClearContents

    ' 恢复合并状态
    If wasMerged Then
        Range(mergeAddr).Merge
    End If

    firstCell.Value = allContent

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.ScreenUpdating = True
End Sub

' ==================== 功能10: 增加固定行高（优化版） ====================
Public Sub AddFixedRowHeight()
    On Error GoTo ErrorHandler

    Dim rng As Range
    Set rng = Selection

    Dim addHeight As String
    addHeight = InputBox("请输入要增加的行高（磅）：", "增加行高", "5")

    If addHeight = "" Then Exit Sub

    Dim addValue As Double
    On Error Resume Next
    addValue = CDbl(addHeight)
    On Error GoTo ErrorHandler

    If addValue <= 0 Then
        MsgBox "请输入有效的正数", vbExclamation
        Exit Sub
    End If

    PushUndo rng, "AddFixedRowHeight"

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 批量设置
    Dim rw As Range
    For Each rw In rng.rows
        rw.RowHeight = rw.RowHeight + addValue
    Next rw

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ==================== 功能11: 智能设置行高（优化均分逻辑） ====================
Public Sub SmartSetRowHeight()
    On Error GoTo ErrorHandler

    Dim rng As Range
    Set rng = Selection

    Dim minHeightStr As String
    Dim addHeightStr As String

    minHeightStr = InputBox("请输入最小行高（磅）：", "最小行高", "15")
    If minHeightStr = "" Then Exit Sub

    addHeightStr = InputBox("请输入增加的固定行高（磅）：", "增加行高", "3")
    If addHeightStr = "" Then Exit Sub

    Dim minHeight As Double
    Dim addHeight As Double

    On Error Resume Next
    minHeight = CDbl(minHeightStr)
    addHeight = CDbl(addHeightStr)
    On Error GoTo ErrorHandler

    If minHeight <= 0 Or addHeight < 0 Then
        MsgBox "请输入有效的数值", vbExclamation
        Exit Sub
    End If

    PushUndo rng, "SmartSetRowHeight"

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 收集所有需要处理的行（去重）
    Dim rowList As Collection
    Set rowList = New Collection

    Dim cell As Range
    For Each cell In rng.Cells
        Dim rowNum As Long
        rowNum = cell.row

        On Error Resume Next
        rowList.Add rowNum, CStr(rowNum)
        On Error GoTo ErrorHandler
    Next cell

    ' 处理每一行
    Dim rowKey As Variant
    For Each rowKey In rowList
        ProcessRowHeightWithDistribution CLng(rowKey), minHeight, addHeight, rng.Worksheet
    Next rowKey

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub ProcessRowHeightWithDistribution(targetRow As Long, minHeight As Double, addHeight As Double, ws As Worksheet)
    Dim rw As Range
    Set rw = ws.rows(targetRow)

    ' 先自动调整
    rw.AutoFit

    Dim baseHeight As Double
    baseHeight = rw.RowHeight

    ' 计算目标行高（不含合并单元格特殊处理）
    Dim targetHeight As Double
    If baseHeight < minHeight Then
        targetHeight = minHeight
    Else
        targetHeight = baseHeight + addHeight
    End If

    ' 检查这一行是否有合并单元格（且是合并区域的第一行）
    Dim mergeInfo As Collection
    Set mergeInfo = New Collection

    Dim cell As Range
    For Each cell In rw.Cells
        If cell.MergeCells Then
            Dim mergeArea As Range
            Set mergeArea = cell.mergeArea

            ' 只处理合并区域的第一行
            If mergeArea.row = targetRow Then
                On Error Resume Next
                mergeInfo.Add mergeArea, mergeArea.Address
                On Error GoTo 0
            End If
        End If
    Next cell

    ' 如果没有合并单元格，直接设置行高
    If mergeInfo.count = 0 Then
        rw.RowHeight = targetHeight
        Exit Sub
    End If

    ' 处理合并单元格 - 计算需要额外分配的高度
    Dim ma As Variant
    For Each ma In mergeInfo
        Dim mergeRng As Range
        Set mergeRng = ma

        ' 计算合并单元格所需总高度
        Dim requiredTotalHeight As Double
        requiredTotalHeight = CalculateMergeAreaRequiredHeight(mergeRng, minHeight, addHeight)

        ' 计算当前各行已有的高度总和
        Dim currentTotalHeight As Double
        currentTotalHeight = 0
        Dim r As Long
        For r = mergeRng.row To mergeRng.row + mergeRng.rows.count - 1
            currentTotalHeight = currentTotalHeight + ws.rows(r).RowHeight
        Next r

        ' 计算需要额外增加的高度
        Dim extraHeightNeeded As Double
        extraHeightNeeded = requiredTotalHeight - currentTotalHeight

        If extraHeightNeeded > 0 Then
            ' 需要额外增加高度，均分到各行
            Dim rowsCount As Long
            rowsCount = mergeRng.rows.count
            Dim extraPerRow As Double
            extraPerRow = extraHeightNeeded / rowsCount

            ' 给合并单元格涉及的每一行增加额外高度
            For r = mergeRng.row To mergeRng.row + rowsCount - 1
                Dim currentRowHeight As Double
                currentRowHeight = ws.rows(r).RowHeight

                ' 计算该行应有的高度：当前高度 + 均分的额外高度
                Dim newRowHeight As Double
                newRowHeight = currentRowHeight + extraPerRow

                ' 确保不低于最小行高
                If newRowHeight < minHeight Then
                    newRowHeight = minHeight
                End If

                ws.rows(r).RowHeight = newRowHeight
            Next r
        Else
            ' 不需要额外高度，但当前行仍需满足最小/增加要求
            If targetRow >= mergeRng.row And targetRow < mergeRng.row + mergeRng.rows.count Then
                ' 当前行在合并区域内，已经处理过
            Else
                rw.RowHeight = targetHeight
            End If
        End If
    Next ma
End Sub

Private Function CalculateMergeAreaRequiredHeight(mergeRng As Range, minHeight As Double, addHeight As Double) As Double
    ' 创建或使用临时工作表计算
    Static tempSheet As Worksheet
    Static initialized As Boolean

    If Not initialized Then
        On Error Resume Next
        Set tempSheet = ThisWorkbook.Worksheets("TempCalcSheet")
        If tempSheet Is Nothing Then
            Set tempSheet = ThisWorkbook.Worksheets.Add
            tempSheet.Name = "TempCalcSheet"
            tempSheet.Visible = xlSheetVeryHidden
        End If
        initialized = True
        On Error GoTo 0
    End If

    If tempSheet Is Nothing Then
        CalculateMergeAreaRequiredHeight = minHeight * mergeRng.rows.count
        Exit Function
    End If

    ' 复制内容到临时工作表计算
    Dim content As String
    content = CStr(mergeRng.Cells(1).Value)

    With tempSheet.Range("A1")
        .Value = content
        .WrapText = True
        .ColumnWidth = mergeRng.Width / 6

        .rows.AutoFit

        Dim neededHeight As Double
        neededHeight = .RowHeight

        ' 应用最小行高和增加逻辑
        If neededHeight < minHeight Then
            neededHeight = minHeight
        Else
            neededHeight = neededHeight + addHeight
        End If

        CalculateMergeAreaRequiredHeight = neededHeight

        ' 清理
        .ClearContents
        .WrapText = False
    End With
End Function

' ==================== 功能12: 设置标准格式 ====================
Public Sub SetStandardFormat()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False

    With ws.PageSetup
        .TopMargin = Application.CentimetersToPoints(2)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin = Application.CentimetersToPoints(1.2)
        .RightMargin = Application.CentimetersToPoints(1.2)

        .PaperSize = xlPaperA4

        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False

        .CenterHorizontally = True
        .CenterVertically = False

        .CenterFooter = " &""Times New Roman""&10&B& " & " &P / &N "
    End With

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
    Application.ScreenUpdating = True
End Sub

' ==================== 功能13: 设置顶端标题行 ====================
Public Sub SetPrintTitleRows()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim rng As Range
    Set rng = Selection

    Dim firstRow As Long
    Dim lastRow As Long
    firstRow = rng.row
    lastRow = firstRow + rng.rows.count - 1

    Dim actualFirstRow As Long
    Dim actualLastRow As Long
    actualFirstRow = firstRow
    actualLastRow = lastRow

    Dim cell As Range
    For Each cell In rng.Cells
        If cell.MergeCells Then
            Dim mergeArea As Range
            Set mergeArea = cell.mergeArea
            If mergeArea.row < actualFirstRow Then actualFirstRow = mergeArea.row
            If mergeArea.row + mergeArea.rows.count - 1 > actualLastRow Then
                actualLastRow = mergeArea.row + mergeArea.rows.count - 1
            End If
        End If
    Next cell

    ws.PageSetup.PrintTitleRows = "$" & actualFirstRow & ":$" & actualLastRow

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
End Sub

' ==================== 功能14: 取消顶端标题行 ====================
Public Sub ClearPrintTitleRows()
    On Error GoTo ErrorHandler

    ActiveSheet.PageSetup.PrintTitleRows = ""

    GoTo CleanUp

ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical

CleanUp:
End Sub

' ==================== 撤销操作入口 ====================
Public Sub UndoLastOperation()
    PopUndo
End Sub

' ==================== 创建备份（优化版） ====================
Public Sub CreateBackup()
    On Error GoTo ErrorHandler

    Dim backupPath As String
    Dim desktopPath As String
    Dim fileExt As String

    ' 统一使用桌面路径
    #If Mac Then
        desktopPath = MacScript("return POSIX path of (path to desktop folder as string)")
        ' 移除末尾斜杠
        If Right(desktopPath, 1) = "/" Then
            desktopPath = Left(desktopPath, Len(desktopPath) - 1)
        End If
    #Else
        desktopPath = Environ("USERPROFILE") & "\Desktop"
    #End If

    ' 根据内容决定扩展名
    If ThisWorkbook.HasVBProject Then
        fileExt = ".xlsm"  ' 有宏
    Else
        fileExt = ".xlsx"  ' 无宏
    End If

    ' 生成文件名：Backup_年月日_时分秒_原文件名.扩展名
    backupPath = desktopPath & "/Backup_" & Format(Now, "yyyymmdd_hhmmss") & "_" & ThisWorkbook.Name

    ' 如果原文件名有扩展名，替换为正确的
    If InStr(backupPath, ".xls") > 0 Then
        backupPath = Left(backupPath, InStrRev(backupPath, ".") - 1) & fileExt
    Else
        backupPath = backupPath & fileExt
    End If

    ' 静默保存副本（不打开，不切换）
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs FileName:=backupPath
    Application.DisplayAlerts = True

    ' 可选：清理旧备份（保留最近5个）
    ' Call CleanOldBackups(desktopPath, 5)

    ' 可选：显示提示（如需静默，删除下面这行）
    ' MsgBox "备份已创建到桌面：" & vbCrLf & backupPath, vbInformation

    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "备份失败：" & Err.Description, vbCritical
End Sub

' ==================== 可选：自动清理旧备份 ====================
Private Sub CleanOldBackups(folderPath As String, keepCount As Integer)
    On Error Resume Next

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim files As Collection
    Dim i As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Set files = New Collection

    ' 收集所有备份文件
    For Each file In folder.files
        If Left(file.Name, 7) = "Backup_" And (Right(file.Name, 5) = ".xlsx" Or Right(file.Name, 5) = ".xlsm" Or Right(file.Name, 5) = ".xlsb") Then
            files.Add file
        End If
    Next file

    ' 如果超过保留数量，删除最旧的
    If files.count > keepCount Then
        ' 按修改时间排序（简化：直接删除前几个）
        For i = 1 To files.count - keepCount
            files(i).Delete
        Next i
    End If

    Set fso = Nothing
End Sub

