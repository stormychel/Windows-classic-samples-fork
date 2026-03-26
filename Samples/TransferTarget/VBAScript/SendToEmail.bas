Attribute VB_Name = "SendToEmail"
Option Explicit

' =============================================================================
' SendToEmail
'
' Adds a button with optional logo to the active sheet. Clicking it sends
' the current tab's contents as an email via the TransferTarget API.
'
' Prerequisites:
'   1. TransferTargetBridge.comhost.dll registered (regsvr32)
'   2. Windows 11 Build 26100.7015+ with TransferTarget API support
' =============================================================================

' Optional: Set this to the path of a custom icon for the button.
' If empty or not found, a plain shape button is used as a fallback.
Private Const ICON_PATH As String = ""
Private Const BTN_NAME As String = "btnSendToEmail"

' Set this to the AppId of the email app you sent to send the sheet to.
' This is the AppId for Outlook (New), but you can substitute whatever app you want.
Private Const APPID As String = "Microsoft.OutlookForWindows_8wekyb3d8bbwe!Microsoft.OutlookforWindows"

#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Function FindWindowA Lib "user32" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

' ---------------------------------------------------------------------------
' Call this once to add the Send To Email button to the active sheet.
' The button persists with the workbook when saved as .xlsm.
' ---------------------------------------------------------------------------
Public Sub AddSendToEmailButton()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Remove existing button if present
    RemoveButton ws

    ' If an icon is requested and is present on disk, then create an icon button
    Dim hasIcon As Boolean
    hasIcon = False
    If Len(ICON_PATH) > 0 Then
        If Dir(ICON_PATH) <> "" Then
            hasIcon = True
        End If
    End If

    If hasIcon Then
        AddIconButton ws
    Else
        AddShapeButton ws
    End If

    MsgBox "Button added! Click it to send this sheet's contents as an email.", vbInformation
End Sub

' ---------------------------------------------------------------------------
' Create a button with an icon
' ---------------------------------------------------------------------------
Private Sub AddIconButton(ws As Worksheet)

    ' Insert the plugin logo as a picture
    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture( _
        Filename:=ICON_PATH, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=ws.Cells(1, 1).Left + 4, _
        Top:=ws.Cells(1, 1).Top + 4, _
        Width:=32, _
        Height:=32)

    pic.Name = BTN_NAME & "_Icon"
    pic.Placement = xlFreeFloating

    ' Add a label shape next to the icon
    Dim lbl As Shape
    Set lbl = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        pic.Left + pic.Width + 4, pic.Top, 130, 32)

    With lbl
        .Name = BTN_NAME & "_Label"
        .Placement = xlFreeFloating
        .TextFrame2.TextRange.Text = "Send to Email"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 4
        .TextFrame2.MarginRight = 4
        .Fill.ForeColor.RGB = RGB(0, 120, 212)
        .Line.Visible = msoFalse
    End With

    ' Group them together
    Dim grp As Shape
    Set grp = ws.Shapes.Range(Array(pic.Name, lbl.Name)).Group
    grp.Name = BTN_NAME
    grp.OnAction = "SendCurrentTabToEmail"

    ' Position at top-right area (column B, row 1)
    grp.Left = ws.Cells(1, 2).Left
    grp.Top = ws.Cells(1, 1).Top + 2
End Sub

' ---------------------------------------------------------------------------
' Create a plain shape button
' ---------------------------------------------------------------------------
Private Sub AddShapeButton(ws As Worksheet)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Cells(1, 2).Left, ws.Cells(1, 1).Top + 2, 170, 34)

    With btn
        .Name = BTN_NAME
        .Placement = xlFreeFloating
        .TextFrame2.TextRange.Text = ChrW(9993) & "  Send to Email"
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.ForeColor.RGB = RGB(0, 120, 212)
        .Line.Visible = msoFalse
        .OnAction = "SendCurrentTabToEmail"
    End With
End Sub

' ---------------------------------------------------------------------------
' Remove existing button from a worksheet
' ---------------------------------------------------------------------------
Private Sub RemoveButton(ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, Len(BTN_NAME)) = BTN_NAME Then
            shp.Delete
        End If
    Next shp
End Sub

' ---------------------------------------------------------------------------
' Main handler: reads current tab, sends to specified app via TransferTarget
' ---------------------------------------------------------------------------
Public Sub SendCurrentTabToEmail()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Check there is data to send
    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
        MsgBox "The active sheet is empty. Nothing to send.", vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "Preparing sheet contents..."

    ' Build plain text and HTML versions of the sheet
    Dim subject As String
    subject = ActiveWorkbook.Name & " - " & ws.Name

    Dim plainText As String
    plainText = RangeToPlainText(ws.UsedRange)

    Dim htmlBody As String
    htmlBody = SheetToHtml(ws)

    ' Create the COM bridge
    On Error Resume Next
    Dim oBridge As Object
    Set oBridge = CreateObject("TransferTargetBridge.Helper")
    On Error GoTo 0

    If oBridge Is Nothing Then
        MsgBox "Could not create TransferTargetBridge.Helper." & vbNewLine & _
               "Ensure the DLL is registered with regsvr32.", vbCritical
        Exit Sub
    End If

    ' Get window handle for UI positioning
#If VBA7 Then
    Dim hWnd As LongPtr
#Else
    Dim hWnd As Long
#End If
    hWnd = Application.hWnd

    ' Send to target app via TransferTarget
    Dim sResult As String
    sResult = oBridge.ShareEmailToApp(APPID, subject, plainText, htmlBody, CLngPtr(hWnd))

    If Left(sResult, 7) = "SUCCESS" Then
        MsgBox "Sheet sent successfully!", vbInformation
    Else
        MsgBox sResult, vbExclamation
    End If

    Set oBridge = Nothing
End Sub

' ---------------------------------------------------------------------------
' Convert sheet to HTML using Excel's built-in PublishObjects
' ---------------------------------------------------------------------------
Private Function SheetToHtml(ws As Worksheet) As String
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\ExcelToOutlook_" & Format(Now, "yyyymmddhhnnss") & ".htm"

    On Error GoTo HtmlFallback

    ' Use Excel's native HTML export (preserves formatting, colors, fonts)
    With ActiveWorkbook.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
        Filename:=tempFile, _
        Sheet:=ws.Name, _
        Source:=ws.UsedRange.Address, _
        HtmlType:=xlHtmlStatic)
        .Publish True
    End With

    ' Read the exported HTML file
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(tempFile) Then
        Dim ts As Object
        Set ts = fso.OpenTextFile(tempFile, 1, False, -1)  ' -1 = TristateTrue (Unicode)
        SheetToHtml = ts.ReadAll
        ts.Close
        fso.DeleteFile tempFile
    Else
        GoTo HtmlFallback
    End If

    Set fso = Nothing
    Exit Function

HtmlFallback:
    ' If PublishObjects fails, build a simple HTML table manually
    SheetToHtml = RangeToHtmlTable(ws.UsedRange)
End Function

' ---------------------------------------------------------------------------
' Fallback: manually build an HTML table from a range
' ---------------------------------------------------------------------------
Private Function RangeToHtmlTable(rng As Range) As String
    Dim html As String
    Dim r As Long, c As Long
    Dim cell As Range

    html = "<html><body>"
    html = html & "<table border='1' cellpadding='5' cellspacing='0' " & _
           "style='border-collapse:collapse;font-family:Calibri,sans-serif;font-size:11pt;'>" & vbNewLine

    For r = 1 To rng.Rows.Count
        html = html & "<tr>"
        For c = 1 To rng.Columns.Count
            Set cell = rng.Cells(r, c)

            ' Use <th> for first row (header)
            Dim tag As String
            If r = 1 Then tag = "th" Else tag = "td"

            ' Build inline style for basic formatting
            Dim sty As String
            sty = ""
            If cell.Font.Bold Then sty = sty & "font-weight:bold;"
            If cell.Font.Italic Then sty = sty & "font-style:italic;"
            If cell.Font.Color <> 0 Then
                sty = sty & "color:" & ColorToHex(cell.Font.Color) & ";"
            End If
            If cell.Interior.ColorIndex <> xlNone And cell.Interior.Color <> 16777215 Then
                sty = sty & "background-color:" & ColorToHex(cell.Interior.Color) & ";"
            End If

            ' Alignment
            Select Case cell.HorizontalAlignment
                Case xlRight: sty = sty & "text-align:right;"
                Case xlCenter: sty = sty & "text-align:center;"
            End Select

            html = html & "<" & tag
            If Len(sty) > 0 Then html = html & " style='" & sty & "'"
            html = html & ">"

            ' Escape HTML special characters
            Dim cellText As String
            cellText = cell.Text
            cellText = Replace(cellText, "&", "&amp;")
            cellText = Replace(cellText, "<", "&lt;")
            cellText = Replace(cellText, ">", "&gt;")
            html = html & cellText

            html = html & "</" & tag & ">"
        Next c
        html = html & "</tr>" & vbNewLine
    Next r

    html = html & "</table></body></html>"
    RangeToHtmlTable = html
End Function

' ---------------------------------------------------------------------------
' Convert Excel color (BGR Long) to CSS hex color (#RRGGBB)
' ---------------------------------------------------------------------------
Private Function ColorToHex(clr As Long) As String
    Dim r As Long, g As Long, b As Long
    r = clr Mod 256
    g = (clr \ 256) Mod 256
    b = (clr \ 65536) Mod 256
    ColorToHex = "#" & Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2)
End Function

' ---------------------------------------------------------------------------
' Convert range to tab-separated plain text
' ---------------------------------------------------------------------------
Private Function RangeToPlainText(rng As Range) As String
    Dim lines() As String
    ReDim lines(1 To rng.Rows.Count)
    Dim r As Long, c As Long

    For r = 1 To rng.Rows.Count
        Dim cols() As String
        ReDim cols(1 To rng.Columns.Count)
        For c = 1 To rng.Columns.Count
            cols(c) = CStr(rng.Cells(r, c).Text)
        Next c
        lines(r) = Join(cols, vbTab)
    Next r

    RangeToPlainText = Join(lines, vbNewLine)
End Function
