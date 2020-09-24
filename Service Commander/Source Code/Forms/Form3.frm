VERSION 5.00
Begin VB.Form Options 
   Caption         =   "Options"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3735
   LinkTopic       =   "Form3"
   ScaleHeight     =   4335
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3945
      Index           =   2
      Left            =   7440
      TabIndex        =   48
      Top             =   255
      Visible         =   0   'False
      Width           =   3660
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Options"
         Height          =   285
         Left            =   15
         TabIndex        =   50
         Top             =   120
         Width           =   3555
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3915
      Index           =   1
      Left            =   3690
      TabIndex        =   47
      Top             =   270
      Visible         =   0   'False
      Width           =   3660
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   285
         Left            =   1185
         TabIndex        =   57
         Top             =   3450
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   2790
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   56
         Text            =   "Form3.frx":0000
         Top             =   405
         Width           =   3555
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "About Service Commander"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   49
         Top             =   90
         Width           =   3555
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3870
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   210
      Visible         =   0   'False
      Width           =   3630
      Begin VB.CommandButton Command2 
         Caption         =   "&Exit"
         Height          =   300
         Index           =   2
         Left            =   2640
         TabIndex        =   58
         Top             =   3570
         Width           =   945
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   21
         Left            =   1350
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   54
         Top             =   2700
         Width           =   225
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   300
         Index           =   1
         Left            =   1335
         TabIndex        =   53
         Top             =   3570
         Width           =   945
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   300
         Index           =   0
         Left            =   45
         TabIndex        =   52
         Top             =   3570
         Width           =   900
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use Colored Text"
         Height          =   195
         Left            =   1013
         TabIndex        =   51
         Top             =   3015
         Width           =   1605
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H0000C000&
         Height          =   225
         Index           =   0
         Left            =   135
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   24
         Top             =   315
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000C0&
         Height          =   225
         Index           =   1
         Left            =   135
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   23
         Top             =   555
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000080FF&
         Height          =   225
         Index           =   2
         Left            =   135
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   22
         Top             =   795
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0FFC0&
         Height          =   225
         Index           =   3
         Left            =   135
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   21
         Top             =   1035
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0FF&
         Height          =   225
         Index           =   4
         Left            =   135
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   20
         Top             =   1275
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C56A31&
         Height          =   225
         Index           =   5
         Left            =   1350
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   19
         Top             =   2460
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF00FF&
         Height          =   225
         Index           =   6
         Left            =   135
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   18
         Top             =   1515
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H0080C0FF&
         Height          =   225
         Index           =   7
         Left            =   135
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   17
         Top             =   1755
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00808080&
         Height          =   225
         Index           =   8
         Left            =   135
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   16
         Top             =   1995
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00008080&
         FillColor       =   &H00FF8080&
         ForeColor       =   &H00FF8080&
         Height          =   225
         Index           =   9
         Left            =   135
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   15
         Top             =   2235
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         Height          =   225
         Index           =   10
         Left            =   1350
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   14
         Top             =   315
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00800080&
         Height          =   225
         Index           =   11
         Left            =   1350
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   13
         Top             =   555
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C000&
         Height          =   225
         Index           =   12
         Left            =   1350
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   12
         Top             =   795
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000040C0&
         Height          =   225
         Index           =   13
         Left            =   1350
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   11
         Top             =   1035
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   225
         Index           =   14
         Left            =   1350
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   10
         Top             =   1275
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000080&
         Height          =   225
         Index           =   15
         Left            =   1350
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   9
         Top             =   1515
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00004080&
         Height          =   225
         Index           =   16
         Left            =   1350
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   8
         Top             =   1755
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00004040&
         Height          =   225
         Index           =   17
         Left            =   1350
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   7
         Top             =   1995
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00008000&
         Height          =   225
         Index           =   18
         Left            =   1350
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   6
         Top             =   2235
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   225
         Index           =   19
         Left            =   135
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   5
         Top             =   2475
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   20
         Left            =   135
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   4
         Top             =   2715
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Selected Text Color"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   21
         Left            =   1650
         TabIndex        =   55
         Top             =   2715
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Running"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   46
         Top             =   315
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Stopped"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   45
         Top             =   555
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Paused"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   44
         Top             =   810
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Starting"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   43
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Stopping"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   420
         TabIndex        =   42
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Selected Back Color"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   1650
         TabIndex        =   41
         Top             =   2475
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kernel Mode"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   420
         TabIndex        =   40
         Top             =   1530
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File System"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   420
         TabIndex        =   39
         Top             =   1770
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Adapter"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   420
         TabIndex        =   38
         Top             =   1995
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Driver"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   420
         TabIndex        =   37
         Top             =   2250
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Automatic"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   1650
         TabIndex        =   36
         Top             =   330
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Boot"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   1650
         TabIndex        =   35
         Top             =   555
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "System"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   1650
         TabIndex        =   34
         Top             =   810
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Manual"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   1650
         TabIndex        =   33
         Top             =   1065
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Disabled"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   1650
         TabIndex        =   32
         Top             =   1290
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Win32 Own Process"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   1650
         TabIndex        =   31
         Top             =   1530
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Color Keys"
         Height          =   225
         Index           =   0
         Left            =   690
         TabIndex        =   30
         Top             =   60
         Width           =   2250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Own Interactive Process"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   1650
         TabIndex        =   29
         Top             =   2010
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Win 32 Shared Process "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   1650
         TabIndex        =   28
         Top             =   1770
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Shared Interactive Process"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   18
         Left            =   1650
         TabIndex        =   27
         Top             =   2250
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text Color"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   420
         TabIndex        =   26
         Top             =   2505
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Back Color"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   420
         TabIndex        =   25
         Top             =   2715
         Width           =   780
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      Height          =   195
      Index           =   1
      Left            =   1140
      TabIndex        =   1
      Top             =   45
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Color Keys"
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   1110
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Width           =   1110
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000F&
      Height          =   255
      Index           =   0
      Left            =   15
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Width           =   1110
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      Height          =   195
      Index           =   2
      Left            =   2250
      TabIndex        =   2
      Top             =   45
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   2205
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Visible         =   0   'False
      Width           =   1110
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Check1_Click()

   'Colored text
   UpdateColors

End Sub

Private Sub Command1_Click()

   'Ok for about frame
   Unload Me

End Sub

Private Sub Command2_Click(Index As Integer)
On Error GoTo SaveClr
  Dim a As Long

   Select Case Index
    Case 0:
      SaveSetting "SerVice Commander\" & App.Path, "Options", "Color Set", "NO"
      GetColors

      For a = 0 To 21
         Picture1(a).BackColor = ColorListNum(a)
      Next

      Check1.Value = ColorListNum(22)
      UpdateColors

    Case 1:

      For a = 0 To 21
         ColorListNum a, Picture1(a).BackColor
      Next

      ColorListNum 22, Check1.Value

      For a = 0 To 5
         Form1.List1(a).BackColor = Picture1(20).BackColor
         Form1.List1(a).Refresh
      Next

      For a = 0 To 22
         SaveSetting "SerVice Commander\" & App.Path, "Options", "Color " & Format(a, "00"), _
            Str(ColorListNum(a))
      Next

      SaveSetting "SerVice Commander\" & App.Path, "Options", "Color Set", "YES"

      UpdateColors

    Case 2:
      Unload Me
   End Select
Exit Sub
SaveClr:
MsgBox "Service Commander cannot access your registry to save your color", 16, "Error Saving Colors"
End Sub

Private Sub Form_Load()

   Me.Icon = Form1.Icon

  Dim a As Long

   For a = 0 To 2
      Label3(a).ForeColor = vbButtonText
      Shape1(a).BackColor = vbButtonFace
   Next

   Frame1(0).Visible = True
   Frame1(0).Left = 50
   Frame1(0).Top = Shape1(0).Height
   Shape1(0).BackColor = vbButtonShadow
   Options.Width = Frame1(0).Width + 200

   For a = 0 To 21
      Picture1(a).BackColor = ColorListNum(a)
   Next

   Check1.Value = ColorListNum(22)
   Me.Show

End Sub

Private Sub Label1_Click(Index As Integer)
On Error GoTo ClrDlg
  Dim Color As Long

   Color = ShowColor(Picture1(Index).BackColor, CC_FULLOPEN Or CC_RGBINIT Or CC_ANYCOLOR, Me.Hwnd)

   If Color <> -1 Then
      Picture1(Index).BackColor = Color
      UpdateColors
   End If
Exit Sub
ClrDlg:
MsgBox "Service Commander encountered a problem accessing the color dialog", 16, "Service Commander Color Dialog"
End Sub

Private Sub Label3_Click(Index As Integer)

  Dim a As Long

   For a = 0 To 2
      Shape1(a).BackColor = vbButtonFace
      Frame1(a).Visible = False
   Next

   Options.Width = Frame1(Index).Width + 200
   Frame1(Index).Visible = True
   Frame1(Index).Left = 50
   Frame1(Index).Top = Shape1(Index).Height
   Shape1(Index).BackColor = vbButtonShadow

   If Index = 1 Then
      Label4.Caption = Label4.Caption & " " & Rev
      Command1.SetFocus
      Text1.Text = "Service Commander FREEWARE" & vbCrLf & "" & vbCrLf & "Backup and Restore" & _
         " Service Profiles" & vbCrLf & "Start, Stop, Pause ,Resume, and Uninstall Services" & _
         vbCrLf & vbCrLf & "Change any service on your system " & vbCrLf & vbCrLf & "Contact:" & _
         " Me.TheUser@Yahoo.com"

      '"Service Commander " & Rev & " - Bilgus 2008
   End If

End Sub

Private Sub Picture1_Click(Index As Integer)

  Dim Color As Long

   Color = ShowColor(Picture1(Index).BackColor, CC_FULLOPEN Or CC_RGBINIT Or CC_ANYCOLOR, Me.Hwnd)

   If Color <> -1 Then
      Picture1(Index).BackColor = Color
      UpdateColors
   End If

End Sub

Private Sub UpdateColors()

   'Color the background an foreground labels and picture boxes
  Dim a As Long

   If Check1.Value = 0 Then

      For a = 0 To 21

         If a <> 19 Then
            Label1(a).BackColor = Picture1(a).BackColor
            Label1(a).ForeColor = Picture1(19).BackColor
         End If

      Next

    Else

      For a = 0 To 20
         If a <> 20 Then Label1(a).BackColor = Picture1(20).BackColor
         Label1(a).ForeColor = Picture1(a).BackColor
      Next

   End If
   Label1(5).ForeColor = Picture1(5).BackColor
   Label1(5).BackColor = Picture1(21).BackColor
   Label1(21).BackColor = Picture1(5).BackColor
   Label1(21).ForeColor = Picture1(21).BackColor
   Label1(20).BackColor = Picture1(20).BackColor
   Label1(20).ForeColor = Picture1(19).BackColor

End Sub

