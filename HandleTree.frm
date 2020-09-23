VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "HandleTree"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   Icon            =   "HandleTree.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   18
      Top             =   6000
      Width           =   492
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show Rect"
      Height          =   375
      Left            =   8160
      TabIndex        =   15
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show"
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make List"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8493
      _Version        =   327682
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   3120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   5055
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "HandleTree.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "HandleTree.frx":0624
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6720
      Width           =   6495
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   8040
      Width           =   9255
   End
   Begin VB.Label Label8 
      Caption         =   "Label5"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7680
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Label5"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   7200
      Width           =   4695
   End
   Begin VB.Label Label6 
      Caption         =   "Label5"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   6480
      Width           =   5895
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   6000
      Width           =   5775
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label lbresults 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   5520
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pp As WINDOWPLACEMENT
Dim classTab As Long
Dim windowTab As Long
Dim SelectedHandle As Long
Dim SelectedClass As String
Dim SelectedTitle  As String

Public Function GetStyle(l As Long) As String
    If l - WS_OVERLAPPED > 0 Then
        GetStyle = "WS_OVERLAPPED"
        l = l - WS_OVERLAPPED
    End If
    If l < 0 Then
        GetStyle = GetStyle & "|" & "WS_POPUP"
        l = l - &H80000000
    End If
    If l - WS_CHILD >= 0 Then
        GetStyle = GetStyle & " | " & "WS_CHILD"
        l = l - WS_CHILD
    End If
    If l - WS_MINIMIZE >= 0 Then
        GetStyle = GetStyle & " | " & "WS_MINIMIZE"
        l = l - WS_MINIMIZE
    End If
    If l - WS_VISIBLE >= 0 Then
        GetStyle = GetStyle & " | " & "WS_VISIBLE"
        l = l - WS_VISIBLE
    End If
    If l - WS_DISABLED >= 0 Then
        GetStyle = GetStyle & " | " & "WS_DISABLED"
        l = l - WS_DISABLED
    End If
    If l - WS_CLIPSIBLINGS >= 0 Then
        GetStyle = GetStyle & " | " & "WS_CLIPSIBLINGS"
        l = l - WS_CLIPSIBLINGS
    End If
    If l - WS_CLIPCHILDREN >= 0 Then
        GetStyle = GetStyle & " | " & "WS_CLIPCHILDREN"
        l = l - WS_CLIPCHILDREN
    End If
    If l - WS_MAXIMIZE >= 0 Then
        GetStyle = GetStyle & " | " & "WS_MAXIMIZE"
        l = l - WS_MAXIMIZE
    End If
    If l - WS_CAPTION >= 0 Then
        GetStyle = GetStyle & " | " & "WS_CAPTION"
        l = l - WS_CAPTION
    End If
    If l - WS_BORDER >= 0 Then
        GetStyle = GetStyle & " | " & "WS_BORDER"
        l = l - WS_BORDER
    End If
    If l - WS_DLGFRAME >= 0 Then
        GetStyle = GetStyle & " | " & "WS_DLGFRAME"
        l = l - WS_DLGFRAME
    End If
    If l - WS_VSCROLL >= 0 Then
        GetStyle = GetStyle & " | " & "WS_VSCROLL"
        l = l - WS_VSCROLL
    End If
    If l - WS_HSCROLL >= 0 Then
        GetStyle = GetStyle & " | " & "WS_HSCROLL"
        l = l - WS_HSCROLL
    End If
    If l - WS_SYSMENU >= 0 Then
        GetStyle = GetStyle & " | " & "WS_SYSMENU"
        l = l - WS_SYSMENU
    End If
    If l - WS_THICKFRAME >= 0 Then
        GetStyle = GetStyle & " | " & "WS_THICKFRAME"
        l = l - WS_THICKFRAME
    End If
    If l - WS_GROUP >= 0 Then
        GetStyle = GetStyle & " | " & "WS_GROUP"
        l = l - WS_GROUP
    End If
    If l - WS_TABSTOP >= 0 Then
        GetStyle = GetStyle & " | " & "WS_TABSTOP"
        l = l - WS_TABSTOP
    End If
    If l - WS_MINIMIZEBOX >= 0 Then
        GetStyle = GetStyle & " | " & "WS_MINIMIZEBOX"
        l = l - WS_MINIMIZEBOX
    End If
    If l - WS_MAXIMIZEBOX >= 0 Then
        GetStyle = GetStyle & " | " & "WS_MAXIMIZEBOX"
        l = l - WS_MAXIMIZEBOX
    End If
End Function

Public Function GetP(S As String) As Long
    Dim i%
    Dim A$
    Dim b As Integer
    b = 0
    For i% = 1 To Len(S)
        If Mid$(S, i%, 1) = " " Then b = b + 1
        If b = 1 Then A$ = A$ + Mid$(S, i%, 1)
    Next i%
    
    GetP = Val(A$)
End Function

Public Function GetS(S As String) As String
    Dim i%
    Dim A$
    Dim b As Integer
    b = 0
    For i% = 1 To Len(S)
        If Mid$(S, i%, 1) = " " Then b = b + 1
        If b >= 2 Then A$ = A$ + Mid$(S, i%, 1)
    Next i%
    
    GetS = A$
End Function

Private Sub Command1_Click()
       Dim sWindowText As String
    Dim sClassname As String
 '  On Error GoTo Err_Handler
    Dim hWnds() As Long
    Dim i%, j%
    Dim hWnd As Long
    Dim imgX As ListImage
    Dim lw As Long
    Dim pwnd As Long
    Dim k$
    Dim iconn% 'The current icon number in a file
    Dim iconfilename$ 'The filename of the icon file(.EXE, .DLL, .ICO)
    Dim numicons% 'The number of icons in a file
    Dim iconmod$
    Dim Iconh As Long
    Dim r As Long
    Dim hModule As Long
    TreeView1.Nodes.Clear
    List2.Clear
    List1.Clear
      Dim nodX As Node
      Pic.BackColor = vbWhite
  r = FindWindowLike(hWnds(), 0, "*", "*")
    If r Then lbresults = "Found : " & r & " windows."
    sWindowText = Space(255)
    r = GetComputerName(sWindowText, 255)
    Set nodX = TreeView1.Nodes.Add(, , "r0", sWindowText)
    nodX.Sorted = True
    Set TreeView1.ImageList = ImageList1
    For i% = 0 To List2.ListCount - 1
        Set nodX = TreeView1.Nodes.Add("r" & GetParent(GetP(List2.List(i%))), tvwChild, "r" & GetP(List2.List(i%)), GetS(List2.List(i%)))
        If GetParent(GetP(List2.List(i%))) = 0 Then
            hModule = GetModuleHandle(0) 'gets handle
            iconfilename$ = GetExeFromHandle(GetP(List2.List(i%)))
            iconmod$ = iconfilename$ + Chr$(0) 'prepares filename
            Iconh = ExtractIcon(hModule, iconmod$, -1) 'gets number of icons
            numicons% = Iconh 'puts it into a variable
            numicons% = numicons% - 1 'Accounts for the first icon, at number 0
            If numicons% > 0 Then
                Iconh = ExtractIcon(hModule, iconmod$, 1) 'Extracts the first icon
                If Iconh <> 0 Then
                    Pic.Cls
                    r = DrawIcon(Pic.hDC, 0, 0, Iconh)
                    Set imgX = ImageList1.ListImages.Add(, , Pic.Image)
                    nodX.Image = ImageList1.ListImages.Count
                End If
            End If
        End If
        nodX.Sorted = True
    Next i%
  
         For i% = 1 To TreeView1.Nodes.Count
            If TreeView1.Nodes(i%).Image = 0 Then
                TreeView1.Nodes(i%).Image = 2
            Else
                TreeView1.Nodes(i%).EnsureVisible
            End If
            
        Next i%
        sWindowText = Space(255)
        r = GetWindowText(128, sWindowText, 255)
        sWindowText = Left(sWindowText, r)
        
        sClassname = Space(255)
        r = GetClassName(128, sClassname, 255)
        sClassname = Left(sClassname, r)
        TreeView1.Nodes(1).Image = 1
       Set TreeView1.SelectedItem = TreeView1.Nodes(1)
       Pic.BackColor = Form1.BackColor
 Exit Sub
Err_Handler:
   If Err = 35601 Then
    Set nodX = TreeView1.Nodes.Add("r0", tvwChild, "r" & GetP(List2.List(i%)), List2.List(i%))
    nodX.Sorted = True

    Resume Next
End If
End Sub

Function FindWindowLike(hWndArray() As Long, ByVal hWndStart As Long, WindowText As String, Classname As String) As Integer
    Dim i%, j%
    Dim hWnd As Long
    Dim hprn As Long
    Dim sWindowText As String
    Dim sClassname As String
    Dim r As Long
    Dim nodX As Node
   'Hold the level of recursion and
   'hold the number of matching windows
    Static level As Integer
    Static found As Integer
  
   'Initialize if necessary
    If level = 0 Then
      found = 0
      ReDim hWndArray(0 To 0)
      If hWndStart = 0 Then
        hWndStart = GetDesktopWindow()
        sWindowText = Space(255)
        r = GetWindowText(hWndStart, sWindowText, 255)
        sWindowText = Left(sWindowText, r)
        
        sClassname = Space(255)
        r = GetClassName(hWndStart, sClassname, 255)
        sClassname = Left(sClassname, r)
        List2.AddItem level & " " & Format$(hWndStart, "0000000000") & " " & sClassname & "     " & sWindowText
      End If
   End If
  
   'Increase recursion counter
    level = level + 1
  
   'Get first child window
    hWnd = GetWindow(hWndStart, GW_CHILD)

    Do Until hWnd = 0
      
       'Search children by recursion
        r = FindWindowLike(hWndArray(), hWnd, WindowText, Classname)
      
       'Get the window text and class name
        sWindowText = Space(255)
        r = GetWindowText(hWnd, sWindowText, 255)
        sWindowText = Left(sWindowText, r)
        
        sClassname = Space(255)
        r = GetClassName(hWnd, sClassname, 255)
        sClassname = Left(sClassname, r)
      
       'Check that window matches the search parameters:
        If (sWindowText Like WindowText) And (sClassname Like Classname) Then
         
            found = found + 1
            ReDim Preserve hWndArray(0 To found)
            hWndArray(found) = hWnd
            If GetParent(hWnd) = 0 Then
                List2.AddItem level & " " & Format$(hWndArray(found), "0000000000") & " " & sClassname & "     " & sWindowText
            Else
                List2.AddItem level + 1 & " " & Format$(hWndArray(found), "0000000000") & " " & sClassname & "     " & sWindowText
            End If

        End If
      
       'Get next child window:
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
        
    Loop
   'Decrement recursion counter
    level = level - 1
  
   'Return the number of windows found
    FindWindowLike = found

End Function



Private Sub Command2_Click()
    Dim r As Long
    Dim hWnd As Long
    Dim wp As WINDOWPLACEMENT
    hWnd = Right$(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 1)
    wp.Length = Len(wp)
    r = GetWindowPlacement(hWnd, wp)
    wp.showCmd = SW_HIDE
    wp.Length = Len(wp)
    r = SetWindowPlacement(hWnd, wp)
End Sub

Private Sub Command3_Click()
    Dim r As Long
    Dim hWnd As Long
    Dim wp As WINDOWPLACEMENT
    hWnd = Right$(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 1)
    wp.Length = Len(wp)
    r = GetWindowPlacement(hWnd, wp)
    wp.showCmd = SW_SHOW
    wp.Length = Len(wp)
    r = SetWindowPlacement(hWnd, wp)
    wp.Length = Len(wp)
    r = GetWindowPlacement(phnd, wp)
    If wp.showCmd = SW_SHOWMINIMIZED Then
    wp.showCmd = SW_SHOWNORMAL
        wp.Length = Len(wp)
        r = SetWindowPlacement(phnd, wp)
    End If
    r = BringWindowToTop(phnd)
End Sub

Private Sub Command4_Click()
    Form1.Caption = "Press ESCAPE to go back"
    Form2.Show vbModal
    Form1.Caption = "HandleTree"
End Sub

Private Sub Form_Activate()
    Form1.Visible = True
End Sub

Private Sub Form_Load()
'     Command1_Click
'   TreeView1_Click
End Sub

Private Sub TreeView1_Click()
  '  On Error GoTo Err_Handler
    Dim sWindowText As String
    Dim sClassname As String
    Dim r As Long
    Dim hWnd As Long
    Dim lpar As Long
    Dim b As Boolean
    Dim p As POINTAPI
    Dim wp As WINDOWPLACEMENT
    
    Dim pwnd As Long
    Dim k$
    Dim iconn% 'The current icon number in a file
    Dim iconfilename$ 'The filename of the icon file(.EXE, .DLL, .ICO)
    Dim numicons% 'The number of icons in a file
    Dim iconmod$
    Dim Iconh As Long
    Dim hModule As Long
    hWnd = Right$(TreeView1.SelectedItem.Key, Len(TreeView1.SelectedItem.Key) - 1)
    Label1 = "Handle:" & hWnd
    Label2 = "Parent:" & GetParent(hWnd)
    sWindowText = Space(255)
    r = GetWindowText(hWnd, sWindowText, 255)
    sWindowText = Left(sWindowText, r)
    
    sClassname = Space(255)
    r = GetClassName(hWnd, sClassname, 255)
    sClassname = Left(sClassname, r)
    Label3 = "Text:" & sWindowText
    Label4 = "Class:" & sClassname
    wp.Length = Len(wp)
    r = GetWindowPlacement(hWnd, wp)
    Label5 = "MaxPos:(" & wp.ptMaxPosition.x & "," & wp.ptMaxPosition.Y & ")"
    Label6 = "MinPos: (" & wp.ptMinPosition.x & "," & wp.ptMinPosition.Y & ")"
    Label7 = "NormPos: (" & wp.rcNormalPosition.Left & "," & wp.rcNormalPosition.Top & "," & wp.rcNormalPosition.Right & "," & wp.rcNormalPosition.Bottom & ")"
    Label10 = GetExeFromHandle(hWnd)
    kk = wp.rcNormalPosition
    
    Select Case wp.showCmd
    Case 0
        Label8 = "SW_HIDE"
    Case 1
        Label8 = "SW_SHOWNORMAL"
    Case 2
        Label8 = "SW_SHOWMINIMIZED"
    Case 3
        Label8 = "SW_SHOWMAXIMIZED"
    Case 4
        Label8 = "SW_SHOWNOACTIVATE"
    Case 5
        Label8 = "SW_SHOW"
    Case 6
        Label8 = "SW_MINIMIZE"
    Case 7
        Label8 = "SW_SHOWMINNOACTIVE"
    Case 8
        Label8 = "SW_SHOWNA"
    Case 9
        Label8 = "SW_RESTORE"
    Case 10
        Label8 = "SW_MAX"
    End Select
    r = GetWindowLong(hWnd, GWL_STYLE)
    Label9.Caption = GetStyle(r)
    phnd = hWnd
    While phnd <> 0
        lpar = phnd
        phnd = GetParent(phnd)
    Wend
    phnd = lpar
    r = GetClientRect(phnd, pk)
    p.x = pk.Left
    p.Y = pk.Top
    r = ClientToScreen(phnd, p)
    pk.Top = p.Y
    pk.Left = p.x
    p.x = pk.Right
    p.Y = pk.Bottom
    r = ClientToScreen(phnd, p)
    pk.Bottom = p.Y
    pk.Right = p.x
    If phnd = hWnd Then
        pk.Top = 0
        pk.Left = 0
    End If
    hModule = GetModuleHandle(0) 'gets handle
    iconfilename$ = GetExeFromHandle(phnd)
    iconmod$ = iconfilename$ + Chr$(0) 'prepares filename
    Iconh = ExtractIcon(hModule, iconmod$, -1) 'gets number of icons
    numicons% = Iconh 'puts it into a variable
    numicons% = numicons% - 1 'Accounts for the first icon, at number 0
    Iconh = ExtractIcon(hModule, iconmod$, 1) 'Extracts the first icon
    Pic.Cls
    If Iconh <> 0 Then
        r = DrawIcon(Pic.hDC, 0, 0, Iconh)
    Else
        If TreeView1.SelectedItem.Key = "r0" Then
            Set Pic.Picture = ImageList1.ListImages(1).Picture
        Else
            Set Pic.Picture = ImageList1.ListImages(2).Picture
        End If
    End If

Err_Handler:
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
    TreeView1_Click
End Sub
