VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Binary metamorphosis to VBA EXCEL"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14010
   FillColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   14010
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame size_row_c 
      Caption         =   "HEX row size: 70 characters/line "
      Height          =   1575
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   13575
      Begin VB.TextBox File_name 
         BackColor       =   &H00996633&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1320
         TabIndex        =   10
         Text            =   "tini.exe"
         ToolTipText     =   "Extracted file name when the source code is executed under the name ..."
         Top             =   480
         Width           =   8535
      End
      Begin VB.CommandButton Save_BAS 
         Caption         =   "Save as .BAS file"
         Height          =   975
         Left            =   10080
         TabIndex        =   9
         Top             =   360
         Width           =   3255
      End
      Begin VB.HScrollBar Row_hex_size 
         Height          =   375
         Left            =   240
         Max             =   300
         Min             =   10
         TabIndex        =   5
         Top             =   960
         Value           =   70
         Width           =   9615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unpack as:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   555
         Width           =   855
      End
   End
   Begin VB.PictureBox bar 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   240
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   901
      TabIndex        =   2
      Top             =   9120
      Width           =   13575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Open EXE file"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   13575
      Begin VB.CommandButton Open_file 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Open file ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text_path 
         BackColor       =   &H00996633&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3240
         TabIndex        =   6
         Text            =   "..."
         Top             =   360
         Width           =   9975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Pack:"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Timer Timer_Row_hex_size 
      Interval        =   100
      Left            =   16560
      Top             =   7320
   End
   Begin VB.TextBox VBCODE 
      BackColor       =   &H00996633&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5655
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   3120
      Width           =   13575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Conversion status:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   8880
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'########################################## -->
'#                                        # -->
'#  Name   : Binary to EXCEL VBA          # -->
'#  Use    : Pack EXE to VB6 source code  # -->
'#  Author : Dr. Paul A. Gagniuc          # -->
'#  Area   : European Union               # -->
'#  Date   : 01/01/2013                   # -->
'#                                        # -->
'#  Mode   : Visual Basic 6.0             # -->
'#                                        # -->
'########################################## -->

Dim sFile As String
Dim BAS_CONTENT As String


Private Sub Convert_to_VB_Click()

    Dim c As String
    
    FF = FreeFile
    Trial = Text_path.Text
    
    Open Trial For Binary Access Read As #FF
        c = Space$(LOF(FF))
        Get #1, 1, c
    Close #FF
    
    
    VBHEX = VBHEX & HexString(c)
    
    step_line = Row_hex_size.Value
    
    u = 0
    
    For i = 1 To Len(VBHEX) Step step_line
        u = u + 1
        VB_line = VB_line & "   f$(" & u & ") = " & Chr(34) & Mid(VBHEX, i, step_line) & Chr(34) & vbCrLf
    Next i
    
    VB_text = VB_text & "Sub Main()" & vbCrLf
    VB_text = VB_text & vbCrLf
    VB_text = VB_text & "   Dim f$(" & u & ")" & vbCrLf
    VB_text = VB_text & vbCrLf & VB_line & vbCrLf
    VB_text = VB_text & "   d = FreeFile" & vbCrLf
    VB_text = VB_text & "   Open Application.ActiveWorkbook.Path & " & Chr(34) & "\" & File_name.Text & Chr(34) & " For Output As #d" & vbCrLf
    VB_text = VB_text & "       For i = 1 To " & u & vbCrLf
    VB_text = VB_text & "           a$ = f$(i)" & vbCrLf
    VB_text = VB_text & "           While Len(a$) > 0" & vbCrLf
    VB_text = VB_text & "               b$ = " & Chr(34) & "&H" & Chr(34) & " & Left$(a$, 2)" & vbCrLf
    VB_text = VB_text & "               a$ = Right$(a$, Len(a$) - 2)" & vbCrLf
    VB_text = VB_text & "               Print #d, Chr$(Val(b$));" & vbCrLf
    VB_text = VB_text & "           Wend" & vbCrLf
    VB_text = VB_text & "       Next i" & vbCrLf
    VB_text = VB_text & "   Close #d" & vbCrLf
    VB_text = VB_text & vbCrLf
    VB_text = VB_text & "   MsgBox " & Chr(34) & File_name.Text & " extracted to app path !" & Chr(34) & ", vbInformation" & vbCrLf
    VB_text = VB_text & "   End" & vbCrLf
    VB_text = VB_text & vbCrLf
    VB_text = VB_text & "End Sub" & vbCrLf
    
    BAS_CONTENT = VB_text
    VBCODE.Text = VB_text

End Sub


Public Function HexString(ByVal EvalString As String) As String

    Dim intStrLen As Variant
    Dim intLoop As Variant
    Dim strHex As String
    
    bar.Cls
    
    intStrLen = Len(EvalString)
    
        For intLoop = 1 To intStrLen
            strHex = strHex & Right$(" " & Hex(Asc(Mid(EvalString, intLoop, 1))), 2)
            bar.Line (0, 100)-((bar.ScaleWidth / intStrLen) * intLoop, 0), &H996633, BF      ' &H008080FF&
            DoEvents
        Next
        
    HexString = strHex
    
End Function


Private Sub Form_Load()
    Text_path.Text = App.Path & "\tini.executable"
    Convert_to_VB_Click
End Sub


Private Sub Open_file_Click()

        Dim CC As cCommonDialog
        Set CC = New cCommonDialog
        
        If CC.VBGetOpenFileName(sFile, , True, , , True, "EXE files (*.exe, *.bin, *.jpg)|*.exe;*.bin;*.jpg|All files|*.*", , , "Open a EXE file", "", Form1.hWnd, 0) Then
            Text_path.Text = sFile
            Convert_to_VB_Click
        End If
        
End Sub


Private Sub Row_hex_size_Change()
    Convert_to_VB_Click
End Sub


Private Sub Save_BAS_Click()

    Dim fname As String
    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    
    fname = Split(File_name.Text, ".")(0) & ".bas"
    
    Dim CurEnvFile As String
    
    If CC.VBGetSaveFileName(fname, CurEnvFile, , "BAS file save (*.bas)|*.bas;|All files|*.*", , "", "bla bla", ".bas", Form1.hWnd, 0) Then
    
        file1 = FreeFile
        Open CurEnvFile For Append As file1
            Print #file1, BAS_CONTENT
        Close #file1
    
    End If
        
End Sub


Private Sub Timer_Row_hex_size_Timer()
    size_row_c.Caption = " HEX row size: [" & Row_hex_size.Value & " characters/line] "
End Sub
