Option Compare Text

Sub SortOutput()
    On Error GoTo ErrorCode  'Sometimes the network drive times out.
    Const file As String = "\\fileserv\SharedFiles\Shipping Output\Customer list.txt"
    Const fDir As String = "\AXOutput"              'File directory.
    Const fName As String = "\Customer list.txt"    'File name.
    Dim appdata As String   'Directory of local user application files.
    Dim FileNum As Integer  'Var for file number.
    Dim fWhole As String    'Var to read in entire consolidation list.
    Dim conList As Variant  'Consolidation List.
    Dim aLength As Integer  'Consolidation list array size.
    Dim cLength As Integer  'Quantity of rows in the table.
    Dim lastColumn As Long  'Last used column in table.
    lastColumn = Worksheets("Sheet1").UsedRange.Columns.Count + 1
    Dim a As Long           'variable for loop.
    Dim b As Long           'variable for loop.
    Dim Today               'Placeholder for date.
    'Variables for column reference.
    Dim colSalesPool As Integer
    Dim colPriority As Integer
    Dim colShipDate As Integer
    Dim colMFGDate As Integer
    Dim colDelName As Integer
    Dim colProdName As Integer
    
    'Find column integers for reference since every output screen is in a different order.
    colSalesPool = ColumnNumberByHeader("Pool")
    colPriority = ColumnNumberByHeader("Priority")
    colShipDate = ColumnNumberByHeader("FWM ship date")
    colMFGDate = ColumnNumberByHeader("Manufacturing due date")
    colDelName = ColumnNumberByHeader("Delivery Name")
    colProdName = ColumnNumberByHeader("Product Name")
    
    appdata = Environ("USERPROFILE")
    Today = Date
    aLength = 0
    
    'Check for the existance of local directory for customer list.
    'If the local directory does not exist, create it.
    If Len(Dir(appdata & fDir, vbDirectory)) = 0 Then
        MkDir (appdata & fDir)
    End If
    
    'Check for customer list on shared network.
    If Len(Dir$(file)) = 0 Then
        'Customer list missing or does not exist on shared network.
        'Check for local copy.
        If Len(Dir$(appdata & fDir & fName)) = 0 Then
            'Local copy of file does not exist.
        Else
            'Local copy exists, pull in data.
            'Should probably check for read only status.
            FileNum = FreeFile()
            Open appdata & fDir & fName For Input As FileNum
            fWhole = Input(LOF(FileNum), FileNum)
            Close FileNum
            conList = Split(fWhole, vbNewLine)
            'Get length of consolidation array.
            aLength = UBound(conList, 1) - LBound(conList, 1) + 1
        End If
    Else
        'Customer list found.
        'Copy or update customer list locally.
        If Len(Dir$(appdata & fDir & fName)) = 0 Then
            'Local copy of file does not exist.
            'Make a local copy of file.
            If GetAttr(file) <> 1 Then
                'File is normal and not in use.
                SetAttr file, vbReadOnly
                FileCopy file, appdata & fDir & fName
                SetAttr file, vbNormal
            Else
                'File is read only, must be in use.
                'We should probably do something here
            End If
        Else
            'Local copy exists.
            'Check file dates, update local file if old.
            If FileDateTime(file) > FileDateTime(appdata & fDir & fName) Then
                'Shared file is newer than local file.
                'Make a local copy of file.
                If GetAttr(file) <> 1 Then
                    'File is normal and not in use.
                    SetAttr file, vbReadOnly
                    FileCopy file, appdata & fDir & fName
                    SetAttr file, vbNormal
                Else
                    'File is read only, must be in use.
                    'We should probably do something here
                End If
            Else
                'Local file is newer or the same as shared file.
            End If
        End If
        'pull in data.
        FileNum = FreeFile()
        Open file For Input As FileNum
        fWhole = Input(LOF(FileNum), FileNum)
        Close FileNum
        conList = Split(fWhole, vbNewLine)
        'Get length of consolidation array.
        aLength = UBound(conList, 1) - LBound(conList, 1) + 1
    End If
    
Skipfiles:       'Jumps to here if there's an issue accessing shared drive, Error 52.
    
    'Remove non-material lines from list
    cLength = 1
    For a = 2 To Rows.Count
        If Cells(a, 1).Value = "Sales order" Then
            'Row is in table
            'Check for spool line.
            If InStr(1, Cells(a, colProdName), "Spool") = 1 Or InStr(1, Cells(a, colProdName), "Customer Supplied") = 1 Then
                'Row is for a non-material line, delete.
                Cells(a, 1).EntireRow.Delete
                a = a - 1
            Else
                'Row is a material line.
                'Check and replace blank shipping dates.
                If Cells(a, colShipDate) = "" Then
                    Cells(a, colShipDate) = Cells(a, colMFGDate)
                End If
                cLength = cLength + 1
            End If
        Else
            'End of table
            Exit For
        End If
    Next a
    
    ' velocity              0
    ' hot                   1
    ' late                  2
    ' due today             3
    ' or sooner status      4
    ' regular status        5
    ' kanban items          5
    
    ' Loop through each row and define the priority for each output order in the last column of the table.
    For a = 2 To cLength
        ' Check for Velocity
        If Cells(a, colSalesPool).Value = "Velocity" Or Cells(a, colSalesPool).Value = "Vel/Bal" Or Cells(a, colSalesPool).Value = "FA Velocit" Then
            Cells(a, lastColumn).Value = "0"
        ' Check for Hot
        ElseIf Cells(a, colPriority).Value = "Hot" Then
            Cells(a, lastColumn).Value = "1"
        ' Check if late
        ElseIf Cells(a, colShipDate) < Today And Cells(a, colShipDate) <> "" Then
            Cells(a, lastColumn).Value = "2"
        ElseIf Cells(a, colShipDate) = Today Then
        ' The order is due today by shipdate
            Cells(a, lastColumn).Value = "3"
        ElseIf Cells(a, colPriority) = "Or Sooner" Then
        'Order status is or sooner.
            Cells(a, lastColumn).Value = "4"
        ElseIf (Cells(a, colSalesPool) = "Standard" Or Cells(a, colSalesPool) = "Std/Balanc") And Cells(a, colPriority) = "Regular" Then
        'Order status is regular.
            Cells(a, lastColumn).Value = "5"
        ElseIf (Cells(a, colSalesPool) = "Ireland" Or Cells(a, colSalesPool) = "IRE/Bal") And Cells(a, colPriority) = "Regular" Then
        'Order is for Ireland container, same priority as a Standard order.
            Cells(a, lastColumn).Value = "5"
        ElseIf (Cells(a, colSalesPool) = "KB" Or Cells(a, colSalesPool) = "KB/Balance") And Cells(a, colPriority) = "Regular" Then
        'Order is for KanBan, sort by date with regular orders.
            Cells(a, lastColumn).Value = "5"
        ElseIf (Cells(a, colSalesPool) = "HFR" Or Cells(a, colSalesPool) = "HFR/Balanc") And Cells(a, colPriority) = "Regular" Then
        'Hold for release, work on by date.
            Cells(a, lastColumn).Value = "5"
        ElseIf (Cells(a, colSalesPool) = "FA Standar" Or Cells(a, colSalesPool) = "FA Balance" Or Cells(a, colSalesPool) = "FA Remake" Or Cells(a, colSalesPool) = "FA Rework") And Cells(a, colPriority) = "Regular" Then
        'Order is a First Article, standard priority.
            Cells(a, lastColumn).Value = "5"
        ElseIf (Cells(a, colSalesPool) = "Remake" Or Cells(a, colSalesPool) = "REM/Balanc" Or Cells(a, colSalesPool) = "Rework") And Cells(a, colPriority) = "Regular" Then
        'Remake
            Cells(a, lastColumn).Value = "5"
        Else
        'Status combination not accounted for.
            Cells(a, lastColumn).Value = "99"
        End If
    Next a
    
    'Change priority for customers in our consolidation list
    'Customers in this list often are "or sooner" but can't actually ship sooner than requested.
    If aLength = 0 Then
        'No consolidation list, do nothing
    Else
        For a = 2 To cLength
            'Compare customer name to exclusion array
            For b = 0 To aLength - 1
                If Cells(a, colDelName).Value = conList(b) Then
                    If Cells(a, lastColumn) >= 4 Then
                        'Customer matches, change priority.
                        Cells(a, lastColumn).Value = "5"
                    End If
                Else
                    'Customer is not in the list.
                End If
            Next b
        Next a
    End If
    
    'Sort table
    Dim lo As Excel.ListObject
    Set lo = ActiveWorkbook.Worksheets("Sheet1").ListObjects("table")
    With lo
    .Sort.SortFields.Clear
        .Sort.SortFields.Add _
            Key:=Range("table[Column1]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        .Sort.SortFields.Add _
            Key:=Range("table[FWM ship date]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    
    Exit Sub    'Actual end of routine, error code below only runs on error
ErrorCode:
    'Currently only addressing error 52, due to time out on tryin to access shared fileserver.
        If Err.Number = 52 Then
            GoTo Skipfiles
        End If
End Sub

Sub AdjustColumns()
    On Error Resume Next    'Most likely error is due to header name.
    Dim a As Long
    Dim colNumber As Integer
    colNumber = ColumnNumberByHeader("Number")
    Dim colQuantity As Integer
    colQuantity = ColumnNumberByHeader("Inventory order quantity")
    Dim lastColumn As Long  'Last used column in table.
    lastColumn = Worksheets("Sheet1").UsedRange.Columns.Count
    
    'Rename column
    Cells(1, colQuantity).Value = "Quantity"
    Cells(1, colNumber).Value = "Sales order"
    
    'Adjust column width
    For a = 1 To lastColumn
        Columns(a).EntireColumn.AutoFit
    Next a
    
    For a = 1 To lastColumn
        Select Case Cells(1, a).Value
            Case Is = "Manufacturing due date"
                Columns(a).EntireColumn.Hidden = True
            Case Is = "Inventory order"
                Columns(a).EntireColumn.Hidden = True
            Case Is = "Lot ID"
                Columns(a).EntireColumn.Hidden = True
            Case Is = "Item number"
                Columns(a).EntireColumn.Hidden = True
            Case Is = "Consolidation"
                Columns(a).EntireColumn.Hidden = True
            Case Is = "Reference"
                Columns(a).EntireColumn.Hidden = True
            Case Is = "Column1"
                Columns(a).EntireColumn.Hidden = True
            Case Is = "Warehouse"
                Columns(a).EntireColumn.Hidden = True
        End Select
    Next a
End Sub

Sub PrintOutput()
'
' PrintOutput Macro
'
' Keyboard Shortcut: Ctrl+j
'
    SortOutput
    AdjustColumns
    
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.text = ""
        .EvenPage.CenterHeader.text = ""
        .EvenPage.RightHeader.text = ""
        .EvenPage.LeftFooter.text = ""
        .EvenPage.CenterFooter.text = ""
        .EvenPage.RightFooter.text = ""
        .FirstPage.LeftHeader.text = ""
        .FirstPage.CenterHeader.text = ""
        .FirstPage.RightHeader.text = ""
        .FirstPage.LeftFooter.text = ""
        .FirstPage.CenterFooter.text = ""
        .FirstPage.RightFooter.text = ""
    End With
    Application.PrintCommunication = True
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.text = ""
        .EvenPage.CenterHeader.text = ""
        .EvenPage.RightHeader.text = ""
        .EvenPage.LeftFooter.text = ""
        .EvenPage.CenterFooter.text = ""
        .EvenPage.RightFooter.text = ""
        .FirstPage.LeftHeader.text = ""
        .FirstPage.CenterHeader.text = ""
        .FirstPage.RightHeader.text = ""
        .FirstPage.LeftFooter.text = ""
        .FirstPage.CenterFooter.text = ""
        .FirstPage.RightFooter.text = ""
    End With
    Application.PrintCommunication = True
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub

Sub CreationtoShipped()
    Dim a As Long
    Dim lastColumn As Long
    Dim colLoaded As Integer
    Dim colCreated As Integer
    Dim buckets(1 To 8) As Integer
    lastColumn = Worksheets("Sheet1").UsedRange.Columns.Count + 1
    colLoaded = ColumnNumberByHeader("Loaded Timestamp")
    colCreated = ColumnNumberByHeader("Created date and time")
    
    For a = 2 To Rows.Count
        If Cells(a, 1).Value <> "" Then
            'Row is in table
            If Cells(a, colLoaded).Value <> "" Then
            'If the loaded timestamp is defined
                Cells(a, lastColumn).Value = Cells(a, colLoaded).Value - Cells(a, colCreated).Value
                Select Case Cells(a, lastColumn)
                    Case Is < 0.99
                        Cells(a, lastColumn + 1).Value = "0 days"
                        buckets(1) = buckets(1) + 1
                    Case 1 To 3.99
                        Cells(a, lastColumn + 1).Value = "1 to 3 days"
                        buckets(2) = buckets(2) + 1
                    Case 4 To 7.99
                        Cells(a, lastColumn + 1).Value = "4 to 7 days"
                        buckets(3) = buckets(3) + 1
                    Case 8 To 14.99
                        Cells(a, lastColumn + 1).Value = "8 to 14 days"
                        buckets(4) = buckets(4) + 1
                    Case 15 To 21.99
                        Cells(a, lastColumn + 1).Value = "15 to 21 days"
                        buckets(5) = buckets(5) + 1
                    Case 22 To 90.99
                        Cells(a, lastColumn + 1).Value = "22 to 90 days"
                        buckets(6) = buckets(6) + 1
                    Case 91 To 180.99
                        Cells(a, lastColumn + 1).Value = "91 to 180 days"
                        buckets(7) = buckets(7) + 1
                    Case Else
                        Cells(a, lastColumn + 1).Value = "181+"
                        buckets(8) = buckets(8) + 1
                End Select
            Else
                Cells(a, lastColumn + 1).Value = "Not shipped."
            End If
        Else
            'End of table
            Exit For
        End If
    Next a
End Sub

Function ColumnNumberByHeader(text As String, Optional headerRange As Range) As Long
    Dim foundRange As Range
    If (headerRange Is Nothing) Then
        Set headerRange = Range("1:1")
    End If

    Set foundRange = headerRange.Find(text)
    If (foundRange Is Nothing) Then
        'Could not find column number by header.
        ColumnNumberByHeader = 0
    Else
        ColumnNumberByHeader = foundRange.Column
    End If
End Function
