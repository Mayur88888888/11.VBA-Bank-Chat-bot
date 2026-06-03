VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "VBA Code Assistance (IDEA from Mayur S ) v1.2"
   ClientHeight    =   5172
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12432
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================================================
' COMPLETE VBA CODE ASSISTANT - ALL IN ONE FORM
' No external files needed - patterns built-in
' Only working buttons included
' ===================================================

Option Explicit

Private colFavorites As collection
Private currentResults As collection

' ========== BUILT-IN PATTERNS (No external file needed) ==========
Private Function GetBuiltInPatterns() As collection
    Dim patterns As New collection
    
    ' WORKSHEET OPERATIONS
    patterns.Add "Sheets(""Sheet1"").Activate"
    patterns.Add "Sheets(""Sheet1"").Select"
    patterns.Add "Dim ws As Worksheet" & vbCrLf & "Set ws = ThisWorkbook.Sheets.Add" & vbCrLf & "ws.Name = ""NewSheet"""
    patterns.Add "ThisWorkbook.Sheets(""Sheet1"").Delete"
    patterns.Add "Sheets(""Sheet1"").Copy After:=Sheets(Sheets.Count)"
    patterns.Add "Sheets(""Sheet1"").Visible = xlSheetHidden"
    patterns.Add "For Each ws In ThisWorkbook.Sheets" & vbCrLf & "    Debug.Print ws.Name" & vbCrLf & "Next ws"
    
    ' RANGE & CELL OPERATIONS
    patterns.Add "Range(""A1"").Value = ""Hello World"""
    patterns.Add "Range(""A1"").Copy Destination:=Range(""B1"")"
    patterns.Add "Range(""A1:A10"").Copy Destination:=Range(""C1"")"
    patterns.Add "Range(""A1"").ClearContents"
    patterns.Add "Range(""A1"").Font.Bold = True"
    patterns.Add "Range(""A1"").Interior.Color = RGB(255, 255, 0)"
    patterns.Add "Dim lastRow As Long" & vbCrLf & "lastRow = Cells(Rows.Count, 1).End(xlUp).Row"
    patterns.Add "Dim lastCol As Long" & vbCrLf & "lastCol = Cells(1, Columns.Count).End(xlToLeft).Column"
    patterns.Add "Cells(1, 1).Value = ""Hello"""
    patterns.Add "Range(""A1"").Offset(1, 0).Value = ""Below"""
    
    ' LOOPS
    patterns.Add "For i = 1 To 10" & vbCrLf & "    Cells(i, 1).Value = i" & vbCrLf & "Next i"
    patterns.Add "For Each cell In Range(""A1:A10"")" & vbCrLf & "    Debug.Print cell.Value" & vbCrLf & "Next cell"
    patterns.Add "Do While i <= 10" & vbCrLf & "    i = i + 1" & vbCrLf & "Loop"
    patterns.Add "For i = 1 To 10 Step 2" & vbCrLf & "    Debug.Print i" & vbCrLf & "Next i"
    
    ' CONDITIONS
    patterns.Add "If Range(""A1"").Value > 100 Then" & vbCrLf & "    MsgBox ""Greater than 100""" & vbCrLf & "Else" & vbCrLf & "    MsgBox ""100 or less""" & vbCrLf & "End If"
    patterns.Add "If Range(""A1"").Value = """" Then" & vbCrLf & "    MsgBox ""Cell is empty""" & vbCrLf & "End If"
    patterns.Add "Select Case Range(""A1"").Value" & vbCrLf & "    Case 1" & vbCrLf & "        MsgBox ""One""" & vbCrLf & "    Case 2, 3, 4" & vbCrLf & "        MsgBox ""Two, three, or four""" & vbCrLf & "    Case Else" & vbCrLf & "        MsgBox ""Other""" & vbCrLf & "End Select"
    
    ' VARIABLES & ARRAYS
    patterns.Add "Dim myString As String" & vbCrLf & "myString = ""Hello"""
    patterns.Add "Dim myNumber As Integer" & vbCrLf & "myNumber = 100"
    patterns.Add "Dim myArray(1 To 5) As String" & vbCrLf & "myArray(1) = ""Apple""" & vbCrLf & "myArray(2) = ""Banana"""
    patterns.Add "Const PI As Double = 3.14159"
    
    ' USER INTERACTION
    patterns.Add "Dim userInput As String" & vbCrLf & "userInput = InputBox(""Enter your name:"")" & vbCrLf & "MsgBox ""Hello "" & userInput"
    patterns.Add "If MsgBox(""Continue?"", vbYesNo) = vbYes Then" & vbCrLf & "    MsgBox ""You clicked Yes""" & vbCrLf & "End If"
    patterns.Add "MsgBox ""Message here"", vbInformation, ""Title"""
    
    ' ERROR HANDLING
    patterns.Add "On Error Resume Next" & vbCrLf & "' Your code here" & vbCrLf & "On Error GoTo 0"
    patterns.Add "On Error GoTo ErrorHandler" & vbCrLf & "' Your code here" & vbCrLf & "Exit Sub" & vbCrLf & "ErrorHandler:" & vbCrLf & "MsgBox ""Error: "" & Err.Description"
    
    ' PERFORMANCE
    patterns.Add "Application.ScreenUpdating = False" & vbCrLf & "Application.Calculation = xlCalculationManual" & vbCrLf & "' Your code here" & vbCrLf & "Application.ScreenUpdating = True" & vbCrLf & "Application.Calculation = xlCalculationAutomatic"
    
    ' FILE OPERATIONS
    patterns.Add "Dim fso As Object" & vbCrLf & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & "If fso.FileExists(""C:\file.txt"") Then" & vbCrLf & "    MsgBox ""File exists""" & vbCrLf & "End If"
    patterns.Add "If Not fso.FolderExists(""C:\NewFolder"") Then" & vbCrLf & "    fso.CreateFolder ""C:\NewFolder""" & vbCrLf & "End If"
    patterns.Add "Workbooks.Open ""C:\file.xlsx"""
    patterns.Add "ThisWorkbook.Save"
    patterns.Add "ThisWorkbook.SaveAs ""C:\newfile.xlsx"""
    patterns.Add "ActiveWorkbook.Close SaveChanges:=True"
    
    ' STRING MANIPULATION
    patterns.Add "Dim fullName As String" & vbCrLf & "fullName = ""John"" & "" "" & ""Doe"""
    patterns.Add "If InStr(""Hello World"", ""World"") > 0 Then" & vbCrLf & "    MsgBox ""Found""" & vbCrLf & "End If"
    patterns.Add "newText = Replace(""Hello World"", ""World"", ""VBA"")"
    patterns.Add "Left(""Hello"", 2)"
    patterns.Add "Right(""Hello"", 2)"
    patterns.Add "Mid(""Hello"", 2, 3)"
    patterns.Add "LCase(""HELLO"")"
    patterns.Add "UCase(""hello"")"
    patterns.Add "Trim(""  Hello  "")"
    
    ' DATE & TIME
    patterns.Add "Dim today As Date" & vbCrLf & "today = Date"
    patterns.Add "Dim nowTime As Date" & vbCrLf & "nowTime = Now"
    patterns.Add "MsgBox Format(Date, ""dd-mmm-yyyy"")"
    patterns.Add "MsgBox Format(Now, ""hh:mm:ss"")"
    patterns.Add "DateAdd(""d"", 7, Date)"
    patterns.Add "DateDiff(""d"", #1/1/2024#, Date)"
    
    ' USER FORMS
    patterns.Add "UserForm1.Show"
    patterns.Add "Unload UserForm1"
    patterns.Add "Me.Hide"
    
    ' FIND & SEARCH
    patterns.Add "Set foundCell = Range(""A:A"").Find(""search text"")" & vbCrLf & "If Not foundCell Is Nothing Then" & vbCrLf & "    MsgBox foundCell.Address" & vbCrLf & "End If"
    
    ' FILTER & SORT
    patterns.Add "Range(""A1:C10"").AutoFilter Field:=1, Criteria1:=""<100"""
    patterns.Add "ActiveSheet.AutoFilterMode = False"
    patterns.Add "Range(""A1:C10"").Sort Key1:=Range(""A1""), Order1:=xlAscending, Header:=xlYes"
    
    ' EMAIL
    patterns.Add "Dim OutApp As Object, OutMail As Object" & vbCrLf & "Set OutApp = CreateObject(""Outlook.Application"")" & vbCrLf & "Set OutMail = OutApp.CreateItem(0)" & vbCrLf & "With OutMail" & vbCrLf & "    .To = ""email@example.com""" & vbCrLf & "    .Subject = ""Hello""" & vbCrLf & "    .Body = ""Message body""" & vbCrLf & "    .Send" & vbCrLf & "End With"
    
    Set GetBuiltInPatterns = patterns
End Function

Private Sub chkCaseSensitive_Click()

End Sub

Private Sub chkExactMatch_Click()

End Sub

Private Sub cmdAddPattern_Click()

End Sub

Private Sub cmdEditPattern_Click()

End Sub

Private Sub cmdDeletePattern_Click()

End Sub

Private Sub cmdExport_Click()

End Sub

Private Sub cmdFavorite_Click()

End Sub

Private Sub cmdSmartSuggest_Click()

End Sub

Private Sub txtDescription_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblStatus_Click()

End Sub

Private Sub UserForm_Initialize()
    ' Initialize collections
    Set colFavorites = New collection
    Set currentResults = New collection
    LoadFavorites
    
    ' Setup categories
    cmbCategory.Clear
    cmbCategory.AddItem "All Categories"
    cmbCategory.AddItem "Worksheets"
    cmbCategory.AddItem "Ranges & Cells"
    cmbCategory.AddItem "Loops"
    cmbCategory.AddItem "Conditions"
    cmbCategory.AddItem "Variables & Arrays"
    cmbCategory.AddItem "User Interaction"
    cmbCategory.AddItem "Error Handling"
    cmbCategory.AddItem "File Operations"
    cmbCategory.AddItem "Strings & Text"
    cmbCategory.AddItem "Dates & Time"
    cmbCategory.ListIndex = 0
    
    ' Setup listbox
    lstResults.Clear
    txtCode.text = ""
    
    ' Status
    lblStatus.Caption = "Ready - " & GetBuiltInPatterns().count & " patterns loaded"
    txtSearch.SetFocus
    
    ' Set form title
    Me.Caption = "VBA Code Assistance (IDEA from Mayur S ) v1.2"
End Sub

Private Sub btnSearch_Click()
    Dim keyword As String
    Dim category As String
    Dim allPatterns As collection
    Dim codeText As String
    Dim results As New collection
    Dim i As Integer
    Dim codeItem As Variant
    
    keyword = Trim(LCase(txtSearch.text))
    
    If cmbCategory.ListIndex > 0 Then
        category = cmbCategory.List(cmbCategory.ListIndex)
    Else
        category = ""
    End If
    
    If keyword = "" And category = "" Then
        MsgBox "Type a keyword or select a category", vbExclamation
        Exit Sub
    End If
    
    lblStatus.Caption = "Searching..."
    Me.Repaint
    
    Set allPatterns = GetBuiltInPatterns()
    
    For Each codeItem In allPatterns
        codeText = CStr(codeItem)
        If MatchesPattern(codeText, keyword, category) Then
            results.Add Array(GetShortDescription(codeText), codeText)
        End If
    Next codeItem
    
    lstResults.Clear
    txtCode.text = ""
    
    If results.count = 0 Then
        lblStatus.Caption = "No matches - try different words"
        MsgBox "No code found. Try different keywords.", vbInformation
    Else
        Set currentResults = results
        For i = 1 To results.count
            lstResults.AddItem results(i)(0)
        Next i
        lblStatus.Caption = "Found " & results.count & " snippets"
    End If
End Sub

Private Function MatchesPattern(codeText As String, keyword As String, category As String) As Boolean
    Dim lcCode As String
    lcCode = LCase(codeText)
    
    ' Check keyword
    If keyword <> "" Then
        If InStr(lcCode, keyword) = 0 Then
            MatchesPattern = False
            Exit Function
        End If
    End If
    
    ' Check category
    If category <> "" And category <> "All Categories" Then
        Select Case category
            Case "Worksheets"
                If InStr(lcCode, "sheet") = 0 And InStr(lcCode, "worksheets") = 0 Then
                    MatchesPattern = False
                    Exit Function
                End If
            Case "Ranges & Cells"
                If InStr(lcCode, "range") = 0 And InStr(lcCode, "cell") = 0 And InStr(lcCode, "cells") = 0 Then
                    MatchesPattern = False
                    Exit Function
                End If
            Case "Loops"
                If InStr(lcCode, "for") = 0 And InStr(lcCode, "while") = 0 And InStr(lcCode, "do") = 0 Then
                    MatchesPattern = False
                    Exit Function
                End If
            Case "Conditions"
                If InStr(lcCode, "if") = 0 And InStr(lcCode, "select case") = 0 Then
                    MatchesPattern = False
                    Exit Function
                End If
            Case "Variables & Arrays"
                If InStr(lcCode, "dim") = 0 And InStr(lcCode, "const") = 0 And InStr(lcCode, "array") = 0 Then
                    MatchesPattern = False
                    Exit Function
                End If
            Case "User Interaction"
                If InStr(lcCode, "msgbox") = 0 And InStr(lcCode, "inputbox") = 0 Then
                    MatchesPattern = False
                    Exit Function
                End If
            Case "Error Handling"
                If InStr(lcCode, "error") = 0 And InStr(lcCode, "resume") = 0 Then
                    MatchesPattern = False
                    Exit Function
                End If
            Case "File Operations"
                If InStr(lcCode, "file") = 0 And InStr(lcCode, "folder") = 0 And InStr(lcCode, "open") = 0 Then
                    MatchesPattern = False
                    Exit Function
                End If
            Case "Strings & Text"
                If InStr(lcCode, "instr") = 0 And InStr(lcCode, "replace") = 0 And InStr(lcCode, "left") = 0 Then
                    MatchesPattern = False
                    Exit Function
                End If
            Case "Dates & Time"
                If InStr(lcCode, "date") = 0 And InStr(lcCode, "now") = 0 Then
                    MatchesPattern = False
                    Exit Function
                End If
        End Select
    End If
    
    MatchesPattern = True
End Function

Private Function GetShortDescription(codeText As String) As String
    Dim lines() As String
    lines = Split(codeText, vbCrLf)
    Dim firstLine As String
    
    If UBound(lines) >= 0 Then
        firstLine = Trim(lines(0))
        If Len(firstLine) > 65 Then
            firstLine = Left(firstLine, 62) & "..."
        End If
        GetShortDescription = firstLine
    Else
        GetShortDescription = "Code snippet"
    End If
End Function

Private Sub lstResults_Click()
    On Error GoTo ClickError
    
    If lstResults.ListIndex >= 0 Then
        If Not currentResults Is Nothing Then
            If lstResults.ListIndex + 1 <= currentResults.count Then
                txtCode.text = currentResults(lstResults.ListIndex + 1)(1)
                lblStatus.Caption = "Code loaded - Click Copy to use"
                
                If IsFavorite(txtCode.text) Then
                    btnFavorite.Caption = "? Remove"
                Else
                    btnFavorite.Caption = "? Save"
                End If
            End If
        End If
    End If
    
    Exit Sub
    
ClickError:
    lblStatus.Caption = "Error loading code"
End Sub

Private Sub btnCopy_Click()
    If txtCode.text = "" Then
        MsgBox "Select a code snippet first", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo CopyError
    
    Dim DataObj As New MSForms.DataObject
    DataObj.SetText txtCode.text
    DataObj.PutInClipboard
    
    lblStatus.Caption = "? Copied! Paste anywhere with Ctrl+V"
    
    Dim oldCaption As String
    oldCaption = btnCopy.Caption
    btnCopy.Caption = "? Copied!"
    Application.Wait Now + TimeValue("00:00:01")
    btnCopy.Caption = oldCaption
    
    Exit Sub
    
CopyError:
    MsgBox "Could not copy: " & Err.Description, vbCritical
End Sub

Private Sub btnFavorite_Click()
    If txtCode.text = "" Then
        MsgBox "Select a code snippet first", vbExclamation
        Exit Sub
    End If
    
    If IsFavorite(txtCode.text) Then
        RemoveFavorite txtCode.text
        btnFavorite.Caption = "? Save"
        lblStatus.Caption = "Removed from favorites"
    Else
        AddFavorite txtCode.text
        btnFavorite.Caption = "? Remove"
        lblStatus.Caption = "Saved to favorites"
    End If
    
    SaveFavorites
End Sub

Private Sub LoadFavorites()
    Dim path As String, fnum As Integer, line As String
    path = ThisWorkbook.path & "\VBA_Favorites.txt"
    Set colFavorites = New collection
    
    If Dir(path) <> "" Then
        fnum = FreeFile
        Open path For Input As #fnum
        Do While Not EOF(fnum)
            Line Input #fnum, line
            On Error Resume Next
            colFavorites.Add line
            On Error GoTo 0
        Loop
        Close #fnum
    End If
End Sub

Private Function IsFavorite(code As String) As Boolean
    Dim fav As Variant
    IsFavorite = False
    For Each fav In colFavorites
        If fav = code Then
            IsFavorite = True
            Exit Function
        End If
    Next fav
End Function

Private Sub AddFavorite(code As String)
    On Error Resume Next
    colFavorites.Add code
    On Error GoTo 0
End Sub

Private Sub RemoveFavorite(code As String)
    Dim i As Integer
    For i = colFavorites.count To 1 Step -1
        If colFavorites(i) = code Then
            colFavorites.Remove i
            Exit For
        End If
    Next i
End Sub

Private Sub SaveFavorites()
    Dim path As String, fnum As Integer, fav As Variant
    path = ThisWorkbook.path & "\VBA_Favorites.txt"
    fnum = FreeFile
    Open path For Output As #fnum
    For Each fav In colFavorites
        Print #fnum, fav
    Next fav
    Close #fnum
End Sub

Private Sub btnSmartSuggest_Click()
    Dim context As String
    Dim allPatterns As collection
    Dim codeItem As Variant
    Dim codeText As String
    Dim results As New collection
    Dim score As Integer
    Dim i As Integer
    Dim keywords() As String
    Dim kw As Variant
    
    If txtSearch.text <> "" Then
        context = txtSearch.text
    Else
        context = InputBox("What do you want to do?", "Smart Suggest")
        If context = "" Then Exit Sub
    End If
    
    lblStatus.Caption = "Finding best matches..."
    Me.Repaint
    
    keywords = Split(LCase(context), " ")
    Set allPatterns = GetBuiltInPatterns()
    
    For Each codeItem In allPatterns
        codeText = CStr(codeItem)
        score = 0
        
        For Each kw In keywords
            If Len(kw) > 2 Then
                If InStr(LCase(codeText), kw) > 0 Then
                    score = score + 10
                End If
            End If
        Next kw
        
        If InStr(LCase(codeText), LCase(context)) > 0 Then
            score = score + 25
        End If
        
        If score > 0 Then
            results.Add Array("[?] " & GetShortDescription(codeText), codeText)
        End If
    Next codeItem
    
    lstResults.Clear
    txtCode.text = ""
    
    If results.count = 0 Then
        lblStatus.Caption = "No suggestions"
        MsgBox "No matches found. Try different words.", vbInformation
    Else
        Set currentResults = results
        For i = 1 To results.count
            lstResults.AddItem results(i)(0)
        Next i
        lblStatus.Caption = "Found " & results.count & " suggestions"
    End If
End Sub

Private Sub btnAddPattern_Click()
    Dim newCode As String
    
    newCode = InputBox("Paste or type your VBA code:", "Add New Pattern")
    If newCode = "" Then Exit Sub
    
    ' Save to external file for persistence
    Dim path As String, fnum As Integer
    path = ThisWorkbook.path & "\VBA_MyPatterns.txt"
    fnum = FreeFile
    
    Open path For Append As #fnum
    Print #fnum, ""
    Print #fnum, newCode
    Print #fnum, ""
    Close #fnum
    
    MsgBox "Pattern saved to " & path & vbCrLf & _
           "Restart assistant to see it.", vbInformation
End Sub

Private Sub cmbCategory_Change()
    If txtSearch.text <> "" Then
        btnSearch_Click
    End If
End Sub

Private Sub btnSettings_Click()
    MsgBox "=== VBA Code Assistant ===" & vbCrLf & vbCrLf & _
           "?? Built-in patterns: " & GetBuiltInPatterns().count & vbCrLf & _
           "? Favorites saved to: " & ThisWorkbook.path & "\VBA_Favorites.txt" & vbCrLf & _
           "?? Custom patterns saved to: " & ThisWorkbook.path & "\VBA_MyPatterns.txt" & vbCrLf & vbCrLf & _
           "?? Shortcuts:" & vbCrLf & _
           "  Ctrl+F - Focus search" & vbCrLf & _
           "  Ctrl+S - AI Suggest" & vbCrLf & _
           "  Ctrl+C - Copy code" & vbCrLf & _
           "  Ctrl+D - Save/Remove Favorite" & vbCrLf & _
           "  Esc - Close", vbInformation
End Sub

Private Sub btnExport_Click()
    If txtCode.text = "" Then
        MsgBox "Select a code snippet first", vbExclamation
        Exit Sub
    End If
    
    Dim savePath As Variant
    Dim fileName As String
    
    fileName = InputBox("File name:", "Export Code", "MyCode")
    If fileName = "" Then Exit Sub
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder"
        If .Show = -1 Then
            savePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Dim fullPath As String
    fullPath = savePath & "\" & fileName & ".txt"
    
    On Error GoTo ExportError
    Open fullPath For Output As #1
    Print #1, txtCode.text
    Close #1
    
    MsgBox "Saved to:" & vbCrLf & fullPath, vbInformation
    lblStatus.Caption = "Exported to " & fileName & ".txt"
    Exit Sub
    
ExportError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Keyboard shortcuts
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 67: If Shift = 2 Then btnCopy_Click          ' Ctrl+C
        Case 83: If Shift = 2 Then btnSmartSuggest_Click  ' Ctrl+S
        Case 70: If Shift = 2 Then                        ' Ctrl+F
            txtSearch.SetFocus
            txtSearch.SelStart = 0
            txtSearch.SelLength = Len(txtSearch.text)
        Case 68: If Shift = 2 Then btnFavorite_Click      ' Ctrl+D
        Case 69: If Shift = 2 Then btnExport_Click        ' Ctrl+E
        Case 65: If Shift = 2 Then btnAddPattern_Click    ' Ctrl+A
        Case 27: Unload Me                                 ' Esc
    End Select
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 Then
        UserForm_KeyDown KeyCode, Shift
    End If
    If KeyCode = 13 Then  ' Enter key
        btnSearch_Click
    End If
End Sub

Private Sub txtCode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 Then
        UserForm_KeyDown KeyCode, Shift
    End If
End Sub

