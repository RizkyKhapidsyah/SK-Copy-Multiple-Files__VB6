VERSION 5.00
Begin VB.Form frmCopyfiles 
   Caption         =   "Copy Files"
   ClientHeight    =   4515
   ClientLeft      =   1545
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6585
   Begin VB.PictureBox Picture2 
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   5235
      TabIndex        =   15
      Top             =   480
      Width           =   5295
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   5295
      End
      Begin VB.TextBox txtDestination 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808000&
      Caption         =   "Delete directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Delete selected files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Create new directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   960
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy Files"
         Height          =   495
         Left            =   720
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CopytoClipboard 
         Caption         =   "Copy"
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   2040
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopy2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Copy files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.DirListBox Dir2 
      Height          =   2565
      Left            =   4680
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Select files"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Source folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Destination folder"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "frmCopyfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdCopy_Click()
Dim result As Long, fileop As SHFILEOPSTRUCT
With fileop
        .hWnd = Me.hWnd
        .wFunc = FO_COPY
        .pFrom = Text1 & vbNullChar & vbNullChar
        .pTo = txtDestination & vbNullChar & vbNullChar
        .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
End With
result = SHFileOperation(fileop)
If result <> 0 Then
        
        MsgBox Err.LastDllError
Else
        If fileop.fAnyOperationsAborted <> 0 Then
                      MsgBox "Operation Failed"
         End If
End If

End Sub


Public Function ExtractName(SpecIn As String) As String
   
Dim i As Integer
Dim SpecOut As String
   
On Error Resume Next
   
For i = Len(SpecIn) To 1 Step -1
If Mid(SpecIn, i, 1) = "\" Then
   SpecOut = Mid(SpecIn, i + 1)
   Exit For
End If
Next i

ExtractName = SpecOut

End Function





Private Sub cmdCopy2_Click()
CopytoClipboard_Click
Dim lstCount As Long
Dim lstIndex As Long
Dim Files() As String
   Dim nRet As Long
   Dim i As Long
  Text1.Text = ""
  lstCount = File1.ListCount - 1
   lstIndex = 0
   Do
    nRet = clipPasteFiles(Files)
   If nRet Then
    If lstIndex = nRet Then Exit Sub
     If lstIndex = nRet - 1 Then

End If
If lstIndex = nRet Then Exit Sub
    Text3 = Val(lstIndex) + 1
    If Val(lstIndex) = nRet Then Exit Sub
      Text1 = Text1 & Files(lstIndex)
     cmdCopy_Click
  DoEvents
   End If
   
      lstIndex = lstIndex + 1
   Text1.Text = ""
   Loop Until lstIndex > lstCount
End Sub

Private Sub Command1_Click()
 On Error GoTo errhandler
    Dim currdir As String, newDir As String
    currdir = Dir2.List(Dir2.ListIndex)
again:
    newDir = InputBox("Type full directory specification:", _
        "Create directory", currdir)
    If newDir = "" Then
         Exit Sub
    End If
    MkDir newDir
    DoEvents
    Dir2.Refresh
    Exit Sub
errhandler:
    If Err.Number = 75 Then
        MsgBox "Directory already exists/access error"
        GoTo again
    End If
   
End Sub

Private Sub Command2_Click()
   If File1.ListCount < 1 Then
         MsgBox "No file in current dir"
         Exit Sub
    End If
    If File1.filename = "" Then
         MsgBox "No file selected yet"
         Exit Sub
    End If
   ' If MsgBox("Sure to delete the selected file(s) ", vbYesNo + vbQuestion) = vbNo Then
    'Exit Sub
  ' End If
   CopytoClipboard_Click
    Dim lstCount As Long
Dim lstIndex As Long
Dim Files() As String
   Dim nRet As Long
   Dim i As Long
  Text1.Text = ""
  lstCount = File1.ListCount - 1
   lstIndex = 0
   Do
    nRet = clipPasteFiles(Files)
   If nRet Then
    If lstIndex = nRet Then Exit Sub
     If lstIndex = nRet - 1 Then
   '
End If 'For i = 0 To nRet - 1
If lstIndex = nRet Then Exit Sub
    Text3 = Val(lstIndex) + 1
    If Val(lstIndex) = nRet Then Exit Sub
      Text1 = Text1 & Files(lstIndex) ' & vbCrLf
     On Error Resume Next
     Kill Text1.Text
  DoEvents
   End If
   
      lstIndex = lstIndex + 1
   Text1.Text = ""
   Loop Until lstIndex > lstCount
End Sub

Private Sub Command3_Click()
On Error GoTo errhandler
    If MsgBox("Sure to delete " & Dir2.List(Dir2.ListIndex) & vbLf & _
           "and all its contents?", vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
    End If
    Dim currdir As String, delDir As String
    currdir = CurDir
    delDir = Dir2.List(Dir2.ListIndex)
    ChDir delDir
    On Error Resume Next
    Kill "*.*"
    On Error GoTo errhandler
    ChDir currdir
    RmDir Dir2.List(Dir2.ListIndex)
    Dir2.Path = drvList.Drive
    Dir2.Refresh
    Exit Sub
errhandler:
   
End Sub

Private Sub CopytoClipboard_Click()
Dim Files() As String
   Dim Path As String
   Dim i As Long, n As Long
   Text1.Text = ""
   Text2.Text = ""
   ' Make sure path has trailing backslash.
   Path = Dir1.Path
   If Right(Path, 1) <> "\" Then
      Path = Path & "\"
   End If
   
   ' Build array of files.
   With File1
      For i = 0 To .ListCount - 1
         If .Selected(i) Then
            ReDim Preserve Files(0 To n) As String
            Files(n) = Path & .List(i)
            n = n + 1
         End If
      Next i
   End With
   
   ' Copy to clipboard.
   If clipCopyFiles(Files) Then
      Text2 = "Success... "
   Else
      Text2 = "failed..."
   End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Dir2_Change()
txtDestination.Text = Dir2.Path
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmAbout.Show vbModal
End Sub
