Attribute VB_Name = "Module4"
' Global variable declarations
Global g_HostSettleTime%
Global g_szPassword$

Sub Main()

Dim wkb As Excel.Workbook
Dim wks As Excel.Worksheet
Dim WS_Count As Integer
Dim Last_col As Integer
Dim I As Integer
Dim rng As Range
Dim c As Integer

Set wkb = ThisWorkbook
Set wks = wkb.Worksheets("List") 'Change the worksheet name here if you would like

'--------------------------------------------------------------------------------
' Get the main system object
    Dim Sessions As Object
    Dim System As Object
    Set System = CreateObject("EXTRA.System")   ' Gets the system object
    If (System Is Nothing) Then
        MsgBox "Could not create the EXTRA System object.  Stopping macro playback."
        Stop
    End If
    Set Sessions = System.Sessions

    If (Sessions Is Nothing) Then
        MsgBox "Could not create the Sessions collection object.  Stopping macro playback."
        Stop
    End If
'--------------------------------------------------------------------------------
' Set the default wait timeout value
    g_HostSettleTime = 3000     ' milliseconds

    OldSystemTimeout& = System.TimeoutValue
    If (g_HostSettleTime > OldSystemTimeout) Then
        System.TimeoutValue = g_HostSettleTime
    End If

' Get the necessary Session Object
    Dim Sess0 As Object
    Set Sess0 = System.ActiveSession
    If (Sess0 Is Nothing) Then
        MsgBox "Could not create the Session object.  Stopping macro playback."
        Stop
    End If
    If Not Sess0.Visible Then Sess0.Visible = True
    Sess0.Screen.WaitHostQuiet (g_HostSettleTime)
    
' This section of code contains the recorded events

    Call AddSheets
    Call CreateIndex

'
'    Application.ScreenUpdating = False
    fName = Environ("USERPROFILE") & "\My Documents\ReflectionScreen.bmp" 'Where the screenshot would be saved under
    Last_row = wks.Cells(wks.Rows.Count, "A").End(xlUp).Row 'Find the last row in column A
    
    Set rng = wks.Range("A2:A" & Last_row) 'Set the range to be worked on
    WS_Count = ThisWorkbook.Worksheets.Count
    wks.Activate
    
    I = 3
    Do Until I = WS_Count + 1
        For c = 1 To rng.Rows.Count
        Sess0.Screen.SendKeys ("H<Tab>" & rng.Cells(c, 1).Value & "<ENTER>") 'Depending on what page needs to be captured, this line will require significant modification

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The follwoinglines might be required if you need to screenshot only selective pages.
'        Do
'        Application.Visible = True
'        intmsg = MsgBox("Capture this page?", vbYesNo)
'            If intmsg = vbNo Then
'            Sess0.Screen.SendKeys ("<ENTER>")
'            End If
'        Loop Until intmsg = vbYes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Now actual screenshot/copy/pastes starts

        Application.Wait (Now + TimeValue("0:00:02"))
        DoEvents
        AppActivate "IGHOST - EXTRA! X-treme"
        Application.SendKeys "(%{1068})" 'print screen
        DoEvents
        Sess0.Screen.SendKeys ("<PF12>")
        DoEvents
        Worksheets(I).Activate
        Worksheets(I).Range("A1").PasteSpecial
        DoEvents
        Application.CutCopyMode = False
        I = I + 1
        Next c
    Loop
    
    Application.SendKeys "{Numlock}" 'Turn back on the Numlock key
    Sess0.Screen.WaitHostQuiet (g_HostSettleTime)
    System.TimeoutValue = OldSystemTimeout
    MsgBox ("Done")

'    Application.ScreenUpdating = True

End Sub

Sub AddSheets()

    Dim xRg As Excel.Range
    Dim wSh As Excel.Worksheet
    Dim wBk As Excel.Workbook

    Set wBk = ThisWorkbook
    Set wSh = wBk.Worksheets("List")
    Last_row = wSh.Cells(wSh.Rows.Count, "A").End(xlUp).Row 'Find the last row in column A
    
    Application.ScreenUpdating = False
    For Each xRg In wSh.Range("A2:A" & Last_row)
        With wBk
            .Sheets.Add after:=.Sheets(.Sheets.Count)
            On Error Resume Next
            ActiveSheet.Name = xRg.Value
            If Err.Number = 1004 Then
              Debug.Print xRg.Value & " already used as a sheet name"
            End If
            On Error GoTo 0
        End With
    Next xRg
    Application.ScreenUpdating = True
End Sub

Sub CreateIndex()

    Dim xAlerts As Boolean
    Dim I  As Long
    Dim xShtIndex As Worksheet
    Dim xSht As Variant
    xAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("Index").Delete
    On Error GoTo 0
    Set xShtIndex = Sheets.Add(Sheets(1))
    xShtIndex.Name = "Index"
    I = 1
    Cells(1, 1).Value = "INDEX"
    For Each xSht In ThisWorkbook.Sheets
        If xSht.Name <> "Index" Then
            I = I + 1
            xShtIndex.Hyperlinks.Add Cells(I, 1), "", "'" & xSht.Name & "'!A1", , xSht.Name
        End If
    Next
    Application.DisplayAlerts = xAlerts
End Sub





