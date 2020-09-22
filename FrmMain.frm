VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3ds File's Reader ** Class **  by Andrea Fontana © 2002. ( trikko@katamail.com )"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdInfo 
      Caption         =   "Object's Infos..."
      Height          =   375
      Left            =   6840
      TabIndex        =   33
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton CmdVisit 
      Caption         =   "Visit My Homepage !!!"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   7080
      Width           =   4335
   End
   Begin VB.ComboBox CmbFile 
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   5400
      Width           =   4335
   End
   Begin VB.CommandButton CmdVote 
      Caption         =   "Rate This Example At Planet-Source-Code.com"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   6480
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control Panel: "
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   4335
      Begin VB.TextBox TxtFps 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   32
         Text            =   "60"
         Top             =   1425
         Width           =   495
      End
      Begin VB.OptionButton OptCtl 
         Caption         =   "Mouse"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   29
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptCtl 
         Caption         =   "KeyBoard"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "FPS:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   1890
         X2              =   1890
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "Left Click = Zoom In  Right Click = Zoom Out"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   2760
         TabIndex        =   21
         ToolTipText     =   "Zoom Out"
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   2040
         TabIndex        =   20
         ToolTipText     =   "Zoom In"
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lnot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   19
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lnot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Z"
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   18
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lnot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   17
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   9
         Left            =   3720
         TabIndex        =   16
         ToolTipText     =   "Rotate Z"
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   8
         Left            =   3360
         TabIndex        =   15
         ToolTipText     =   "Rotate Z"
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   14
         ToolTipText     =   "Rotate Y"
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   3240
         TabIndex        =   13
         ToolTipText     =   "Rotate Y"
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   12
         ToolTipText     =   "Rotate X"
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   11
         ToolTipText     =   "Rotate X"
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   10
         ToolTipText     =   "Move Right"
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   9
         ToolTipText     =   "Move Down"
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   8
         ToolTipText     =   "Move Left"
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   7
         ToolTipText     =   "Move Up"
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.PictureBox PicCanvas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   5055
      Left            =   120
      ScaleHeight     =   335
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   583
      TabIndex        =   1
      Top             =   120
      Width           =   8775
      Begin VB.Label lblInt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "How To Display 3dStudio's Files In Your Apps!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   960
         Width           =   8775
      End
      Begin VB.Label lblInt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rate this piece of code at: www.planet-source-code.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   26
         Top             =   3360
         Width           =   8775
      End
      Begin VB.Label lblInt 
         BackStyle       =   0  'Transparent
         Caption         =   "Homepage:     www.it.owns.it - www.vbp.it/trikko"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   4
         Left            =   2295
         TabIndex        =   25
         Top             =   2520
         Width           =   4815
      End
      Begin VB.Label lblInt 
         BackStyle       =   0  'Transparent
         Caption         =   "  Class:     cls3dsReader.cls"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   24
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label lblInt 
         BackStyle       =   0  'Transparent
         Caption         =   " E-Mail:     trikko@katamail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   3
         Left            =   2625
         TabIndex        =   23
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label lblInt 
         BackStyle       =   0  'Transparent
         Caption         =   "Author:     Andrea Fontana"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   22
         Top             =   2040
         Width           =   3135
      End
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Load 3ds File !"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label status 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ready."
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   7680
      Width           =   8775
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'' What's on the world is this? >>>
Private Const pi As Single = 3.14159265358979
Dim WithEvents c3ds As cls3dsReader
Attribute c3ds.VB_VarHelpID = -1
Private Const SW_SHOW = 5       ' Displays Window in its current size
                                ' and position
Private Const SW_SHOWNORMAL = 1 ' Restores Window if Minimized or
                                ' Maximized
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
         "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
         String, ByVal lpFile As String, ByVal lpParameters As String, _
         ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias _
         "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
         String, ByVal lpResult As String) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim bMouse As Boolean
Dim xDiff As Integer
Dim yDiff As Integer


Public Sub OpenLink(Link As String)
Dim FileName As String, Dummy As String
Dim BrowserExec As String * 255
Dim RetVal As Long
Dim Filenumber As Integer
If LCase(Left(Link, 7)) = "mailto:" Then
ShellExecute FrmMain.hWnd, "OPEN", Link, vbNullString, App.Path, 1
Exit Sub
End If
FileName = App.Path + "\temphtm.HTM"
Filenumber = FreeFile                    ' Get unused file number
Open FileName For Output As #Filenumber  ' Create temp HTML file
    Write #Filenumber, "<HTML> <\HTML>"  ' Output text
Close #Filenumber                        ' Close file
BrowserExec = Space(255)
RetVal = FindExecutable(FileName, Dummy, BrowserExec)
BrowserExec = Trim(BrowserExec)
Kill FileName
' If an application is found, launch it!
If RetVal <= 32 Or IsEmpty(BrowserExec) Then ' Error
ShellExecute FrmMain.hWnd, "OPEN", Link, vbNullString, App.Path, 1
Exit Sub
Else
RetVal = ShellExecute(FrmMain.hWnd, "open", BrowserExec, _
Link, Dummy, SW_SHOWNORMAL)
Exit Sub
If RetVal <= 32 Then        ' Error
    ShellExecute FrmMain.hWnd, "OPEN", Link, vbNullString, App.Path, 1
    Exit Sub
End If
End If
ShellExecute FrmMain.hWnd, "OPEN", Link, vbNullString, App.Path, 1
End Sub
Private Sub c3ds_Loading(Percentage As Single)
status.Caption = FormatNumber(Percentage, 1) & " % Completed..."
End Sub

Private Sub c3ds_LoadingComplete()
status.Caption = "Loading Completed!"
End Sub

Private Sub CmbFile_Change()
CmbFile.SelStart = Len(CmbFile.Text)
End Sub


Private Sub CmdInfo_Click()
MsgBox "Number of Solids: " & c3ds.SolidsCount & vbCrLf & _
       "Number of Triangles: " & c3ds.PolysCount & vbCrLf & _
       "Number of Points: " & c3ds.PointsCount, vbOKOnly + vbInformation, "Statistics:"

End Sub

Private Sub CmdLoad_Click()
Dim I As Integer
If CmbFile.Text = App.Path & "\ToyDog.3ds" Then TxtFps.Text = 8
If CmbFile.Text = App.Path & "\Cube.3ds" Then TxtFps.Text = 60
If c3ds.Load3dsFile(CmbFile.Text) Then
For I = 0 To 5
    lblInt(I).Visible = False
Next
c3ds.CreateBuffer PicCanvas
c3ds.Render PicCanvas.hdc

Else
MsgBox "Error File Not Exists!!!", vbOKOnly + vbCritical, "Error!"
End If
PicCanvas.SetFocus
End Sub

Private Sub CmdVisit_Click()
OpenLink "http://www.vbp.it/trikko"
PicCanvas.SetFocus
End Sub

Private Sub CmdVote_Click()
OpenLink "http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=37484&lngWId=1"
PicCanvas.SetFocus
End Sub

Private Sub Form_Activate()
bMouse = True
OptCtl_Click 0

End Sub

Private Sub Form_Click()
PicCanvas.SetFocus
End Sub

Private Sub Form_Load()
Set c3ds = New cls3dsReader
PicCanvas.ForeColor = &HC0C0C0
PicCanvas.BackColor = &H80000001
CmbFile.AddItem App.Path & "\ToyDog.3ds"
CmbFile.AddItem App.Path & "\Cube.3ds"
CmbFile.Text = App.Path & "\Cube.3ds"
CmbFile.SelStart = Len(CmbFile.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
c3ds.DeleteBuffer
bMouse = False
End
End Sub



Private Sub Frame1_Click()
PicCanvas.SetFocus

End Sub


Private Sub lKey_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If bMouse = True Then Exit Sub
Select Case Index
    Case 1
        c3ds.MoveX -5
    Case 3
        c3ds.MoveX 5
    Case 0
        c3ds.MoveY -5
    Case 2
        c3ds.MoveY 5
    Case 4
        c3ds.RotateX -pi / 24
    Case 5
        c3ds.RotateX pi / 24
    Case 6
        c3ds.RotateY -pi / 24
    Case 7
        c3ds.RotateY pi / 24
    Case 8
        c3ds.RotateZ -pi / 24
    Case 9
        c3ds.RotateZ pi / 24
    Case 10
        c3ds.MoveZ 1.05
    Case 11
        c3ds.MoveZ 1 / 1.05
    End Select
c3ds.Render PicCanvas.hdc

End Sub

Private Sub OptCtl_Click(Index As Integer)
Dim tCount As Long

bMouse = CBool(OptCtl(0).Value)
PicCanvas.SetFocus

If bMouse = False Then Exit Sub

Do While bMouse = True
    tCount = GetTickCount()
    c3ds.RotateY CDbl((pi / 18) * (xDiff / 291))
    c3ds.RotateX CDbl((pi / 18) * (yDiff / 167))
    c3ds.Render PicCanvas.hdc
    Do While -tCount + GetTickCount < 1 / CDbl(Val(TxtFps.Text) / 1000)
        DoEvents
    Loop
    If bMouse = False Then Exit Do
Loop
End Sub



Private Sub PicCanvas_KeyDown(KeyCode As Integer, Shift As Integer)
If bMouse = True Then Exit Sub
Select Case KeyCode
    Case Asc("A")
        c3ds.MoveX -5
    Case Asc("D")
        c3ds.MoveX 5
    Case Asc("W")
        c3ds.MoveY -5
    Case Asc("S")
        c3ds.MoveY 5
    Case Asc("R")
        c3ds.RotateX -pi / 24
    Case Asc("T")
        c3ds.RotateX pi / 24
    Case Asc("F")
        c3ds.RotateY -pi / 24
    Case Asc("G")
        c3ds.RotateY pi / 24
    Case Asc("V")
        c3ds.RotateZ -pi / 24
    Case Asc("B")
        c3ds.RotateZ pi / 24
    Case Asc("Q")
        c3ds.MoveZ 1.05
    Case Asc("E")
        c3ds.MoveZ 1 / 1.05
    End Select
c3ds.Render PicCanvas.hdc
End Sub

Private Sub PicCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bMouse = False Then Exit Sub

Select Case Button
    Case 1
        c3ds.MoveZ 1.5
    Case 2
        c3ds.MoveZ 1 / 1.5
End Select

End Sub

Private Sub PicCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

xDiff = PicCanvas.ScaleWidth / 2 - X
yDiff = PicCanvas.ScaleHeight / 2 - Y


End Sub

Private Sub PicCanvas_Paint()
c3ds.Render PicCanvas.hdc
End Sub

Private Sub status_Click()
PicCanvas.SetFocus

End Sub


Private Sub TxtFps_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
 KeyAscii = 0
End If
 
End Sub
