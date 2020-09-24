VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "Form2"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5610
   LinkTopic       =   "Form2"
   ScaleHeight     =   6630
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4755
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   375
      Width           =   5340
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   105
      TabIndex        =   2
      Top             =   5145
      Width           =   5370
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3570
         Top             =   225
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   315
         Index           =   1
         Left            =   4035
         TabIndex        =   11
         ToolTipText     =   "Click to Close"
         Top             =   900
         Width           =   1230
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Refresh"
         Height          =   315
         Index           =   0
         Left            =   4035
         TabIndex        =   10
         ToolTipText     =   "Refresh Info"
         Top             =   225
         Width           =   1230
      End
      Begin VB.CommandButton Command1 
         Caption         =   "C&hange"
         Height          =   315
         Left            =   4035
         TabIndex        =   6
         ToolTipText     =   "Change Service Options"
         Top             =   555
         Width           =   1230
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "Form2.frx":0006
         Left            =   1215
         List            =   "Form2.frx":0022
         TabIndex        =   5
         Text            =   "Service Type"
         ToolTipText     =   "What Type Of Service This Is"
         Top             =   195
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "Form2.frx":00BF
         Left            =   1740
         List            =   "Form2.frx":00D2
         TabIndex        =   4
         Text            =   "Start Type"
         ToolTipText     =   "Start Type"
         Top             =   525
         Width           =   1770
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         ItemData        =   "Form2.frx":0119
         Left            =   1740
         List            =   "Form2.frx":0129
         TabIndex        =   3
         Text            =   "Error Control"
         ToolTipText     =   "Error Level On Failure To Start"
         Top             =   870
         Width           =   1770
      End
      Begin VB.Label Label2 
         Caption         =   "Error Control"
         Height          =   255
         Index           =   2
         Left            =   45
         TabIndex        =   9
         Top             =   930
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Start Type"
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   8
         Top             =   555
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Service Type"
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   315
      Index           =   2
      Left            =   4140
      TabIndex        =   12
      Top             =   6060
      Width           =   1230
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   45
      Width           =   5295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SvcName As String
Private moreinfo As SvcReturn
Public SettingsOnly As Boolean


Private Sub Command1_Click()

   'Change to selected Options
   'Service Type/Start Type/ Error control
  Dim lReturn As Long
  Dim NewServiceType As Long
  Dim NewStartType As Long
  Dim NewErrorControl As Long

   NewServiceType = NewServType(Combo1(0).ListIndex)
   NewStartType = Combo1(1).ListIndex
   NewErrorControl = Combo1(2).ListIndex

   If ServType(moreinfo.StartType) = Combo1(0).ListIndex Then NewServiceType = SERVICE_NO_CHANGE
   If Val(moreinfo.StartType) = Combo1(1).ListIndex Then NewStartType = SERVICE_NO_CHANGE
   If Val(moreinfo.ErrorControl) = Combo1(2).ListIndex Then NewErrorControl = SERVICE_NO_CHANGE
   'if nothing new is selected then pass no change
   lReturn = SetServiceConfig(SvcName, NewServiceType, NewStartType, NewErrorControl, 0&)
   'Debug.Print lReturn & " SetServiceConfig Ret"
   Command2_Click 0 'refresh form1 list
   If lReturn <> 0 Then Call MsgBox(ErrLib(lReturn), 48, "Error Can't Change Service")
   If SettingsOnly = True Then Unload Me

End Sub

Private Sub Command2_Click(Index As Integer)

   'refresh and Close

   If Index = 0 Then
      'Refresh
      moreinfo = GetServiceConfig(SvcName)
      LoadInfo
      Form1.GetServiceList Form1.LastIndex

    Else
      'Close
      Unload Me
   End If

End Sub

Private Sub Command2_MouseDown(Index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

   'in unattended mode right click on close to stop countdown to close

   If Button = 2 And Index = 2 Then
      Timer1.Enabled = False
      Command2(2).Caption = "&Close"
      Command2(2).ToolTipText = "Click to Close"
   End If

End Sub

Private Sub Form_Load()

   'Make it for the more info option otherwise the log option makes it say its title
   Me.Icon = Form1.Icon
   Form1.Enabled = False
   Me.Caption = "More Info and Settings"
   moreinfo = GetServiceConfig(SvcName)
   LoadInfo

End Sub

Private Sub Form_Resize()

   If Me.WindowState <> vbMinimized And SettingsOnly = False Then

      'can't resize form on minimize
      If Me.Width < 5730 Then Me.Width = 5730
      If Me.Height < 3375 Then Me.Height = 3375
      Text1.Width = Form2.Width - 390
      Text1.Height = Form2.Height - (Frame1.Height + 1000)
      Frame1.Top = Text1.Height + Text1.Top
      Command2(2).Top = Text1.Height + Text1.Top + 500 'Close (under frame for unattended)
      'move frames,button, and textbox to proper position

   End If

   If SettingsOnly = True Then
      Me.BorderStyle = 0
      Frame1.Top = 0
      Text1.Visible = False
      Command2(2).Visible = False
      Me.Height = Frame1.Height
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   'if not unattened enable form1
   If InStr(1, Command$, "load", vbTextCompare) = 0 Then Form1.Enabled = True
   If Form1.Enabled = True Then Form1.Command1.SetFocus

End Sub

Private Sub LoadInfo()

   'loads the selected service info into textbox

   Label1.Caption = SvcName
   On Error Resume Next

   With moreinfo
      If .Account = "" Then .Account = "None"
      If .Dependencies = "" Then .Dependencies = "None"
      If .LoadOrderGroup = "" Then .LoadOrderGroup = "None"
      If .PathName = "" Then .PathName = "Unknown"
      If .StartType = "" Then .StartType = "Unknown"
      If .ServiceType = "" Then .ServiceType = "Unknown"
      Text1.Text = vbNullString
      Text1.Text = "SERVICE TYPE:" & vbCrLf & SvcType(.ServiceType) & vbCrLf & vbCrLf & "START" & _
         " TYPE:" & vbCrLf & StartType(.StartType)
      Text1.Text = Text1.Text & vbCrLf & vbCrLf
      Text1.Text = Text1.Text & "DESCRIPTION:" & vbCrLf & .DisplayName & vbCrLf & vbCrLf & "ERROR" & _
         " CONTROL:" & vbCrLf & SvcError(.ErrorControl)
      Text1.Text = Text1.Text & vbCrLf & vbCrLf
      Text1.Text = Text1.Text & "ACCOUNT:" & vbCrLf & .Account & vbCrLf & vbCrLf & "DEPENDENCIES:" _
         & vbCrLf & .Dependencies
      Text1.Text = Text1.Text & vbCrLf & vbCrLf
      Text1.Text = Text1.Text & "LOAD ORDER GROUP:" & vbCrLf & .LoadOrderGroup & vbCrLf & vbCrLf & _
         "PATH:" & vbCrLf & .PathName
      Text1.Text = Text1.Text & vbCrLf & vbCrLf
      Text1.Text = Text1.Text & "TAG ID:" & vbCrLf & .TagId & vbCrLf & vbCrLf & "CURRENT STATUS:" & _
         vbCrLf & SvcState(GetServiceStatus(SvcName))
      'Nice spaced format
      Combo1(0).ListIndex = ServType(.ServiceType)
      Combo1(1).ListIndex = .StartType
      Combo1(2).ListIndex = .ErrorControl
      'set the comboboxes to the current values
   End With

End Sub

Private Function NewServType(Stype As Variant) As Long

   'returns the service type based on combo1(0) value

   Select Case Stype
    Case 0:
      NewServType = 1
      'Kernel Mode  Driver
    Case 1:
      NewServType = 2
      'File System Driver
    Case 2:
      NewServType = 4
      'Adapter
    Case 3:
      NewServType = 8
      'Driver
    Case 4:
      NewServType = 16
      'Win32 Own Process
    Case 5:
      NewServType = 32
      'Win 32 Shared Process
    Case 6:
      NewServType = 272
      'Own Interactive Process
    Case 7:
      NewServType = 288
      'Shared Interactive Process
    Case Else:
      'Debug.Print Stype & " New Service Type"

   End Select

End Function

Private Function ServType(Stype As Variant) As Long

   'returns the combo number to combo1(0) based on service type

   Select Case Stype
    Case 1:
      ServType = 0
      'Kernel Mode  Driver
    Case 2:
      ServType = 1
      'File System Driver
    Case 4:
      ServType = 2
      'Adapter
    Case 8:
      ServType = 3
      'Driver
    Case 10, 16:
      ServType = 4
      'Win32 Own Process
    Case 20, 32:
      ServType = 5
      'Win 32 Shared Process
    Case 100, 272:
      ServType = 6
      'Own Interactive Process
    Case 120, 288:
      ServType = 7
      'Shared Interactive Process
    Case Else:
      'Debug.Print Stype & " combo1(0)"
      Combo1(0).Text = Stype
   End Select

End Function

Private Sub Timer1_Timer()

   'closes form after 30 seconds
   'used for unattended
   Timer1.Tag = Timer1.Tag + 1
   Command2(2).Caption = "&Close (" & 30 - Timer1.Tag & ")"

   If Timer1.Tag = 30 Then
      Timer1.Enabled = False
      Unload Me
      Exit Sub
   End If

End Sub

