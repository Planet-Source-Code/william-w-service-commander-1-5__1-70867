VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Service Commander 1.5 - Bilgus"
   ClientHeight    =   5790
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Please Wait..."
      Height          =   1185
      Left            =   4035
      TabIndex        =   25
      Top             =   4695
      Visible         =   0   'False
      Width           =   3810
      Begin VB.CommandButton Command5 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   2955
         TabIndex        =   27
         ToolTipText     =   "Cancel Operation"
         Top             =   795
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Saving Service 1 of 1"
         Height          =   210
         Index           =   0
         Left            =   45
         TabIndex        =   26
         Top             =   240
         Width           =   3720
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Height          =   390
         Index           =   1
         Left            =   60
         TabIndex        =   36
         Top             =   690
         Width           =   2850
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   165
         Index           =   1
         Left            =   90
         Shape           =   4  'Rounded Rectangle
         Top             =   510
         Width           =   15
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         Height          =   165
         Index           =   0
         Left            =   90
         Shape           =   4  'Rounded Rectangle
         Top             =   510
         Width           =   3600
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   5000
      Left            =   0
      Max             =   0
      SmallChange     =   100
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3450
      Width           =   11865
   End
   Begin VB.Frame Frame5 
      Caption         =   "Search Column"
      Height          =   930
      Left            =   7890
      TabIndex        =   38
      Top             =   4695
      Visible         =   0   'False
      Width           =   7275
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   720
         Left            =   3960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   43
         Text            =   "Form1.frx":0442
         Top             =   135
         Width           =   3225
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Next"
         Height          =   345
         Index           =   1
         Left            =   3060
         TabIndex        =   42
         Top             =   165
         Width           =   840
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Done"
         Height          =   330
         Index           =   0
         Left            =   3060
         TabIndex        =   40
         Top             =   555
         Width           =   840
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   45
         TabIndex        =   39
         Top             =   195
         Width           =   2985
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   45
         TabIndex        =   41
         Top             =   540
         Width           =   2430
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2940
         Left            =   90
         TabIndex        =   29
         Top             =   405
         Width           =   11625
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "CONTINUE AT YOUR OWN RISK!!!"
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1230
            TabIndex        =   35
            Top             =   1560
            Width           =   5445
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   " ALWAYS MAKE A BACKUP OF YOUR CURRENT CONFIGURATION"
            ForeColor       =   &H000000FF&
            Height          =   510
            Left            =   1185
            TabIndex        =   34
            Top             =   1935
            Width           =   5670
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Stopping or uninstalling the wrong service could cause DEVASTATING or UNDESIRED results and windows may or may not stop you."
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   1095
            TabIndex        =   33
            Top             =   1035
            Width           =   5745
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "All Rights Reserved BgsSoft 2008"
            Height          =   255
            Left            =   990
            TabIndex        =   32
            Top             =   810
            Width           =   5670
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "FREE FOR NON COMMERICAL PUBLIC USE"
            Height          =   255
            Left            =   945
            TabIndex        =   31
            Top             =   540
            Width           =   5790
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Service Commander 1.5 - Bilgus "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1050
            TabIndex        =   30
            Top             =   60
            Width           =   6315
         End
         Begin VB.Image Image1 
            Height          =   780
            Left            =   105
            Top             =   75
            Width           =   810
         End
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         CausesValidation=   0   'False
         Height          =   3210
         Left            =   -30
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2980
         Index           =   5
         ItemData        =   "Form1.frx":0498
         Left            =   10500
         List            =   "Form1.frx":049A
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   390
         Width           =   1245
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2980
         Index           =   4
         ItemData        =   "Form1.frx":049C
         Left            =   9270
         List            =   "Form1.frx":049E
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   390
         Width           =   1545
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2980
         Index           =   3
         ItemData        =   "Form1.frx":04A0
         Left            =   7215
         List            =   "Form1.frx":04A2
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   390
         Width           =   2325
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2980
         Index           =   2
         ItemData        =   "Form1.frx":04A4
         Left            =   4095
         List            =   "Form1.frx":04A6
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   390
         Width           =   3405
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2985
         Index           =   1
         ItemData        =   "Form1.frx":04A8
         Left            =   1935
         List            =   "Form1.frx":04AA
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   390
         Width           =   2430
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2980
         Index           =   0
         ItemData        =   "Form1.frx":04AC
         Left            =   90
         List            =   "Form1.frx":04AE
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   390
         Width           =   2130
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Service Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   90
         MousePointer    =   9  'Size W E
         TabIndex        =   22
         ToolTipText     =   "Right Click to Sort"
         Top             =   180
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Service Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1935
         MousePointer    =   9  'Size W E
         TabIndex        =   21
         ToolTipText     =   "Right Click to Sort"
         Top             =   180
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Service Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4095
         MousePointer    =   9  'Size W E
         TabIndex        =   20
         ToolTipText     =   "Right Click to Sort"
         Top             =   180
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Service Path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7215
         MousePointer    =   9  'Size W E
         TabIndex        =   19
         ToolTipText     =   "Right Click to Sort"
         Top             =   180
         Width           =   2070
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Service State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   9270
         MousePointer    =   9  'Size W E
         TabIndex        =   18
         ToolTipText     =   "Right Click to Sort"
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Service Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   10500
         MousePointer    =   9  'Size W E
         TabIndex        =   17
         ToolTipText     =   "Right Click to Sort"
         Top             =   180
         Width           =   1230
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   495
      Left            =   60
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5085
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   450
      Top             =   4695
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   15
      Top             =   4710
   End
   Begin VB.Frame Frame2 
      Height          =   930
      Left            =   0
      TabIndex        =   24
      Top             =   3765
      Width           =   11880
      Begin VB.CommandButton Command1 
         Caption         =   "&Get Services"
         Height          =   315
         Left            =   3135
         TabIndex        =   8
         ToolTipText     =   "Get Services"
         Top             =   180
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":04B0
         Left            =   75
         List            =   "Form1.frx":04D5
         TabIndex        =   7
         Text            =   "All Services"
         ToolTipText     =   "Sort Services By Service Type"
         Top             =   180
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   5370
         TabIndex        =   9
         ToolTipText     =   "Text From Item Selected Above"
         Top             =   195
         Width           =   4965
      End
      Begin VB.CommandButton Command2 
         Caption         =   "S&tart Service"
         Height          =   315
         Left            =   1665
         TabIndex        =   12
         ToolTipText     =   "Apply Setting At Left To Selected Service"
         Top             =   555
         Width           =   1395
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Save Current Profile"
         Height          =   315
         Index           =   0
         Left            =   3135
         TabIndex        =   13
         ToolTipText     =   "Save Current Displayed Services To File In Box At Right"
         Top             =   555
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":05F2
         Left            =   10380
         List            =   "Form1.frx":05FF
         TabIndex        =   10
         Text            =   "All States"
         ToolTipText     =   "Sort Services By Running or Inactive"
         Top             =   180
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form1.frx":062D
         Left            =   75
         List            =   "Form1.frx":0640
         TabIndex        =   11
         Text            =   "Start Service"
         ToolTipText     =   "Action To Do To Service"
         Top             =   555
         Width           =   1590
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   5310
         TabIndex        =   14
         Text            =   "CurrentConfiguration.ini"
         ToolTipText     =   "Type In Your Own File Name Here Or Select One From The List"
         Top             =   555
         Width           =   2130
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Restore Profile From &File"
         Height          =   315
         Index           =   1
         Left            =   7440
         TabIndex        =   15
         Top             =   555
         Width           =   2235
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Options"
         Height          =   315
         Left            =   9720
         TabIndex        =   16
         ToolTipText     =   "Change Color Keys and Learn About This Wonderful Program"
         Top             =   555
         Width           =   2100
      End
   End
   Begin VB.Menu MainMenu 
      Caption         =   "Main Menu"
      Visible         =   0   'False
      Begin VB.Menu moreinfo 
         Caption         =   "More &Info"
         Index           =   0
      End
      Begin VB.Menu ChgSets 
         Caption         =   "Change &Settings"
         Index           =   0
      End
      Begin VB.Menu CState 
         Caption         =   "&Change State"
         Begin VB.Menu Cserv 
            Caption         =   "&Start Service"
            Index           =   0
         End
         Begin VB.Menu Cserv 
            Caption         =   "S&top Service"
            Index           =   1
         End
         Begin VB.Menu Cserv 
            Caption         =   "&Pause Service"
            Index           =   2
         End
         Begin VB.Menu Cserv 
            Caption         =   "&Resume Service"
            Index           =   3
         End
         Begin VB.Menu Cserv 
            Caption         =   "&Uninstall Service"
            Index           =   4
         End
      End
      Begin VB.Menu SearchColumn 
         Caption         =   "Search &This Column"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------API Declares------------------------------

'Retrieves from INI
Private Declare Function GetPrivateProfileString Lib "kernel32" _
      Alias "GetPrivateProfileStringA" ( _
      ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, _
      ByVal lpDefault As String, _
      ByVal lpReturnedString As String, _
      ByVal nSize As Long, _
      ByVal lpFileName As String) As Long

'Saves to INI
Private Declare Function WritePrivateProfileString Lib "kernel32" _
      Alias "WritePrivateProfileStringA" ( _
      ByVal lpApplicationName As String, _
      ByVal lpKeyName As Any, _
      ByVal lpString As Any, _
      ByVal lpFileName As String) As Long

'--------------------Global Variables-----------------
Private SvcName As String
Private CurrentItem As Long
Private DesiredState As Long
Private CancelOpp As Boolean
Private LastCol As Long
Private LastSel As Long
Private StartPos As Long
Private Const LB_FINDSTRING = &H18F


Public Sub ChangeSvcState(State As Long, SvcNamel As String)

   'its self explanatory
   On Error Resume Next
  Dim lReturn As Long

   SvcName = SvcNamel
   'SvcName = List1(1).List(List1(Index).ListIndex)
   CurrentItem = List1(4).ListIndex

   Select Case State
    Case -1, 0: 'Start Selected Service
      lReturn = StartNTService(SvcName)
      DesiredState = 4

    Case 1: 'Stop Selected Service
      lReturn = StopNTService(SvcName)
      DesiredState = 1

    Case 2: 'Pause Selected Service
      lReturn = PauseNTService(SvcName)
      DesiredState = 7

    Case 3: 'Resume Selected Service
      lReturn = ResumeNTService(SvcName)
      DesiredState = 4

    Case 4: 'Uninstall Selected Service
      lReturn = DeleteNTService(SvcName)
      DesiredState = 0
   End Select

   If lReturn = 0 Then

      If LastIndex <> -1 And DesiredState <> 0 Then

         If DesiredState <> GetServiceStatus(SvcName) Then List1(4).List(CurrentItem) = _
            SvcState(GetServiceStatus(SvcName))
         'if the state doesn't match what it already is,  change it and enable timer 2
         'to watch till it changes to the desired state
         Timer2.Enabled = True
       Else

         'otherwise update the whole list
         GetServiceList LastIndex - 1
      End If

    Else
      Call MsgBox(SvcName & " = " & SvcState(DesiredState) & vbCrLf & ErrLib(lReturn), 48, "Error" & _
         " Cannot Complete Request")
   End If

   'End If

End Sub

Private Sub ChgSets_Click(Index As Integer)

   Form2.SvcName = List1(1).List(List1(Index).ListIndex)
   Form2.SettingsOnly = True
   Load Form2
   Form2.BorderStyle = 0
   Form2.Top = Form1.Top
   Form2.Left = Form1.Left
   Form2.Caption = vbNullString
   Form2.Show

End Sub

Private Sub Combo1_Click()

   'show selected service type
   'All/All win32/All driver/Kernel Mode  Driver Services/File System Driver Services
   '/Adapter Services/Driver Services/Win32 Own Process Services/Win 32 Shared Process Services
   '/Own Interactive Process Services/Shared Interactive Process Services
   GetServiceList

End Sub

Private Sub Combo2_Click()

   'show only selected service state
   'Active/Inactive/All
   GetServiceList

End Sub

Private Sub Combo3_Click()

   'Change the button caption to the service state selected
   Command2.Caption = Combo3.List(Combo3.ListIndex)

End Sub

Private Sub Command1_Click()

   'get Services or refresh if previously selected
   GetServiceList LastIndex

End Sub

Private Sub Command2_Click()

   'Change selected service to state choosen in combo3

   If List1(1).ListIndex <> -1 Then
      If Combo3.ListIndex < 4 Then
         Call ChangeSvcState(CLng(Combo3.ListIndex), List1(1).List(List1(1).ListIndex))
       Else

         Select Case MsgBox("Are you sure you want to uninstall this service ?" & vbCrLf & "" & _
            vbCrLf & "You can only get the service back by reinstalling." & vbCrLf & "" & vbCrLf & _
            "If it is a Windows service you will have to reinstall " & vbCrLf & "Windows or use" & _
            " System restore to undo your changes." & vbCrLf & "" & vbCrLf & "Continue?", 308, _
            "Uninstall Service")
          Case 6:
            'Yes Button Selected
            Call ChangeSvcState(CLng(Combo3.ListIndex), List1(1).List(List1(1).ListIndex))

          Case 7:
            'No Button Selected
         End Select

      End If
    Else
      Call MsgBox("Please Select a Service to apply this action to", 64, "No Service Selected")
   End If

End Sub

Private Sub Command3_Click(Index As Integer)

   'Save and restore profile
  Dim CurTxt As String

   If Index = 0 Then
      'Save Current Profile
      If Combo4.Text = "" Then Combo4.Text = "CurrentConfiguration.ini"
      If LCase$(Right$(Combo4.Text, 4)) <> ".ini" Then Combo4.Text = Combo4.Text & ".ini"
      'get selected (or typed) filename from the combo box if no .ini put one
      WriteCurConfig App.Path & "\Configurations\" & Combo4.Text
      CurTxt = Combo4.Text
      Combo4.Clear
      Command3(0).SetFocus
      GetFilesInPath App.Path & "\Configurations\", ".ini"
      'we clear the combobox each time so we can update files but we need to save the
      'current file thats where curtxt comes in

      If Combo4.ListCount = 0 Then
         Command3(1).Enabled = False
         Combo4.Text = "CurrentConfiguration.ini"
         'if there are no .ini files disable restore
       Else
         Combo4.Text = CurTxt
         Command3(1).Enabled = True
         'set the combo box to curtxt
      End If
      
    Else
      CurTxt = Combo4.Text
      Combo4.Clear
      GetFilesInPath App.Path & "\Configurations\", ".ini"

      If Combo4.ListCount = 0 Then
         Command3(1).Enabled = False
         Combo4.Text = "CurrentConfiguration.ini"
         'if no files exist set the default file name and disable restore
       Else
         Combo4.Text = CurTxt
         Command3(1).Enabled = True
         SetStoredConfig App.Path & "\Configurations\" & Combo4.Text
         'otherwise enable restore profile and restore the filename that was in there
      End If

   End If
   
End Sub

Private Sub Command3_MouseMove(Index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

   On Error GoTo InfoErr

   If LCase$(Dir(App.Path & "\Configurations\" & Combo4.Text)) = LCase$(Combo4.Text) And Index = 1 _
      Then
      'If the file exists pull info from it
      Dim Filename As String
      Dim lpBuffer As String
      Dim lpAppName As String
      Dim lpKeyName As String
      Dim nSize As Long
      Dim StrOut As String

      nSize = 255
      Filename = App.Path & "\Configurations\" & Combo4.Text
      lpBuffer = Space(255)
      lpAppName = "Scmdr Profile Data 0.1.4" 'The section of the INI...
      'left this 0.1.4 for backward compability
      lpKeyName = "Date" ' THE LINE OF THE INI
      GetPrivateProfileString lpAppName, lpKeyName, "N/A", lpBuffer, nSize, Filename
      StrOut = "Date: " & Mid(Trim$(lpBuffer), 1, Len(Trim$(lpBuffer)) - 1)

      lpBuffer = Space(255)
      lpKeyName = "Time" ' THE LINE OF THE INI
      GetPrivateProfileString lpAppName, lpKeyName, "N/A", lpBuffer, nSize, Filename
      StrOut = StrOut & " Time: " & Mid(Trim$(lpBuffer), 1, Len(Trim$(lpBuffer)) - 1)

      lpBuffer = Space(255)
      lpKeyName = "Entries" ' THE LINE OF THE INI
      GetPrivateProfileString lpAppName, lpKeyName, "N/A", lpBuffer, nSize, Filename
      StrOut = StrOut & " Entries: " & Mid(Trim$(lpBuffer), 1, Len(Trim$(lpBuffer)) - 1)

      lpBuffer = Space(255)
      lpKeyName = "Extra Info" ' THE LINE OF THE INI
      GetPrivateProfileString lpAppName, lpKeyName, "", lpBuffer, nSize, Filename
      If Asc(LCase$(Trim$(lpBuffer))) <> 0 Then StrOut = StrOut & " Info: " & Mid(Trim$(lpBuffer), _
         1, Len(Trim$(lpBuffer)) - 1)
   End If

   If StrOut = "" Then StrOut = "No Info To Display Or File Does Not Exist"
   Command3(1).ToolTipText = vbCrLf & StrOut & vbCrLf
InfoErr:

End Sub

Private Sub Command4_Click()

   Load Options

End Sub

Private Sub Command5_Click()

   'Cancel button on progressbar
   CancelOpp = True

End Sub

Private Sub Command7_Click(Index As Integer)

   If Index = 0 Then
      Frame5.Visible = False
    Else
      SearchList StartPos
   End If

End Sub

Private Sub CommandLineInput()

   'if save or load are in the command line if so find the file names

  Dim LoadPath As String
  Dim SavePath As String
  Dim Str1 As Long
  Dim Str2 As Long

   Str1 = InStr(1, Command$, "Load=", vbTextCompare) 'get where load= starts
   Str2 = InStr(1, Command$, "Save=", vbTextCompare) 'get where save= starts

   If Str1 * Str2 <> 0 Then
      'if neither = 0 then...

      If Str1 > Str2 Then
         'if load= is after save=
         SavePath = Mid(Command$, Str2 + 5, Str1 - 7)
         'save must be first so beginning of 'save=' + the len of 'save='
         'the length is from one before load= to the end of Save=

         LoadPath = Mid(Command$, Str1 + 5) 'load is at the end

      End If

      If Str1 < Str2 Then
         'if load= is before save=
         SavePath = Mid(Command$, Str2 + 5) 'save is at the end

         LoadPath = Mid(Command$, Str1 + 5, Str2 - 7)
         'load must be first so beginning of 'load=' + the len of 'load='
         'the length is from one before save= to the end of load=

      End If

    Else

      If Str1 = 0 And Str2 = 0 Then Exit Sub 'if neither is specified then exit
      Form1.Visible = False
      'if both are specified then do a save before a load Duh....

      If Str2 <> 0 Then
         SavePath = Mid(Command$, Str2 + 5) '+5 gets rid of Save=
         If SavePath = "" Then Exit Sub 'if no filename exit
         GetServiceList
         WriteCurConfig App.Path & "\Configurations\" & SavePath
      End If

      If Str1 <> 0 Then
         LoadPath = Mid(Command$, Str1 + 5) '+5 gets rid of Load=
         If LoadPath = "" Then Exit Sub 'if no filename exit
         SetStoredConfig App.Path & "\Configurations\" & LoadPath
      End If

   End If
   Unload Me 'unload form1 when done

End Sub

Private Sub Cserv_Click(Index As Integer)

   If List1(1).ListIndex <> -1 Then
      If Index < 4 Then
         Call ChangeSvcState(CLng(Index), List1(1).List(List1(1).ListIndex))
       Else

         Select Case MsgBox("Are you sure you want to uninstall this service ?" & vbCrLf & "" & _
            vbCrLf & "You can only get the service back by reinstalling." & vbCrLf & "" & vbCrLf & _
            "If it is a Windows service you will have to reinstall " & vbCrLf & "Windows or use" & _
            " System restore to undo your changes." & vbCrLf & "" & vbCrLf & "Continue?", 308, _
            "Uninstall Service")
          Case 6:
            'Yes Button Selected
            Call ChangeSvcState(CLng(Index), List1(1).List(List1(1).ListIndex))

          Case 7:
            'No Button Selected
         End Select

      End If
    Else
      Call MsgBox("Please Select a Service to apply this action to", 64, "No Service Selected")
   End If

End Sub

Private Sub Form_Load()

   If CheckIsNT = True Then
      If Command$ = "" Then
         'visible load
         Dim a As Long
         GetColors

         For a = 0 To 5
            List1(a).BackColor = ColorListNum(20)
         Next

         'get the files if there aren't any then disable restore of profiles
         GetFilesInPath App.Path & "\Configurations\", ".ini"
         If Combo4.ListCount = 0 Then Command3(1).Enabled = False
         Command3(0).Enabled = False 'disable save profile until you bring up a profile to save
         Command2.Enabled = False 'disabled until you select a service to
         'stop/start/pause/resume/uninstall
         Label2.Caption = "Service Commander " & Rev & " - Bilgus"
         Form1.Caption = "Service Commander " & Rev & " - Bilgus"
         Image1.Picture = Me.Icon 'put the icon into the warning about frame3 that
         'dissappears when you bring up your profile

         'Subclass the "Frame", to Capture the Listbox Notification Messages ...
         lPrevWndProc = SetWindowLong(Frame1.Hwnd, GWL_WNDPROC, AddressOf SubClassedList)

         If Form1.Width > Screen.Width Then
            'for screens too small to display the whole window we will
            'shrink the window to the screen width and move it to the top left
            Form1.Width = Screen.Width
            Form1.Top = 0
            Form1.Left = 0
         End If

       Else
         CommandLineInput
         'do command line options
      End If

    Else
      ' can't use on non NT OS so let them know that
      MsgBox "This program is designed for the manipulation of NT based OS services" & vbCrLf & _
         "                 " & vbCrLf & "                 It was not designed for, nor will it run" & _
         " on this OS ", 16, "Service Commander " & Rev
      Unload Me
      Exit Sub
   End If

End Sub

Private Sub Form_Resize()

   On Error GoTo FormError
  Dim FrameWidth As Long

   If Me.WindowState <> vbMinimized Then
      'we can only resize when not minimized
      'this puts the frames in their proper places
      Dim a As Long
      VScroll1.Visible = False
      VScroll1.Value = List1(0).TopIndex
      HScroll1.Visible = False
      HScroll1.Width = Form1.Width - 100
      HScroll1.Left = 0
      HScroll1.Min = 0
      HScroll1.Max = 0
      Frame1.Left = 0

      If Me.Height < 3400 Then Me.Height = 3400

      Frame1.Height = Form1.Height - (Frame2.Height + 800)
      Frame1.Width = List1(5).Left + List1(5).Width + 50
      HScroll1.Top = Frame1.Height + Frame1.Top
      Frame2.Top = Frame1.Height + Frame1.Top + HScroll1.Height

      If Frame1.Width > Frame2.Width Then
         FrameWidth = Frame1.Width
         'I just changed this to find which frame is widest
       Else
         FrameWidth = Frame2.Width
         'This fixes an error with the dynamic vertical scrollbar
      End If

      If Form1.Width < FrameWidth Then
         If Form1.Width < Frame1.Width + Frame1.Left Then VScroll1.Visible = True
         'we only want the vertical scrollbar showing if the right edge
         'of frame 1 is obscured
         HScroll1.Min = 45 + Frame1.Left

         If (FrameWidth - Form1.Width) + 100 <= 32768 Then
            'this fixes an error from the scrollbar overflow
            HScroll1.Max = (FrameWidth - Form1.Width) + 100
            HScroll1.Visible = True

          Else

            For a = 0 To 5
               List1(a).Width = 32700 \ 5
               Label1(a).Left = List1(a).Left

               If a < 5 Then
                  Label1(a).Width = List1(a).Width - 285
                  List1(a + 1).Left = List1(a).Left + List1(a).Width - 285
                  ' Match up all the edges of labels and lists (except 5 it has nothing overlapping
                  '   it)
                Else
                  Label1(a).Width = List1(a).Width
               End If

            Next

            Frame1.Width = List1(5).Left + List1(5).Width
            Frame3.Width = Frame1.Width
            HScroll1.Max = (Frame1.Width - Form1.Width) + 100
            HScroll1.Visible = True
         End If

      End If

      For a = 0 To 5
         List1(a).Height = Frame1.Height - List1(a).Top
         'set the height of the lists to the height of the frame
      Next

      Frame3.Height = List1(0).Height
      VScroll1.Top = List1(0).Top + 30
      VScroll1.Left = Form1.Width - VScroll1.Width - 100
      VScroll1.Height = List1(0).Height - 60
      VScroll1.Max = List1(0).ListCount
      VScroll1.Min = 0
   End If

   Command1.SetFocus
   Frame5.Left = Frame2.Left + 3090
   Frame5.Top = Frame2.Top
   Exit Sub
FormError:
   MsgBox "An error has occured while resizing the form." & vbCrLf & "Please Close and Restart" & _
      " Service Commander." & vbCrLf & "Error: " & Err.Number & vbCrLf & Err.Description, 0, _
      "Service Commander " & Rev

End Sub

Private Sub Form_Unload(Cancel As Integer)

   'Release the SubClassing, Very Important to Prevent Crashing!
   Call SetWindowLong(Frame1.Hwnd, GWL_WNDPROC, lPrevWndProc)
   Unload Options
   Unload Form2

End Sub

Private Sub GetFilesInPath(Path As String, FileExt As String)

   'this gets all the FileExt (.ini) files in the path (\configuration\)
   On Error GoTo PathError
  Dim objFSO As FileSystemObject

   Set objFSO = New FileSystemObject
  Dim objFiles As Files
  Dim objFile As File

   Set objFiles = objFSO.GetFolder(Path).Files

   For Each objFile In objFiles
      'if the last 4 letters match the fileext then add them to the combobox

      If LCase$(Right$(objFile.Name, Len(FileExt))) = LCase$(FileExt) Then
         Combo4.AddItem objFile.Name
      End If

   Next
   Exit Sub
PathError:

   If Err.Number = 76 Then
      ' if the folder isn't there give them an option of making one

      Select Case MsgBox("Configurations Folder Not Found!" & vbCrLf & "Make One?" & vbCrLf & "" & _
         vbCrLf & "Note: if you choose not to have a Configurations " & vbCrLf & "folder you" & _
         " cannot store or retrieve profiles", 52, "Service Commander " & Rev)
       Case 6:
         'Yes Button Selected
         MkDir App.Path & "\Configurations"
         Exit Sub

       Case 7:
         'No Button Selected
         Exit Sub
      End Select

    Else
      Call MsgBox("Get Files In Path Error" & vbCrLf & Err.Description & " = " & Err.Number, 48, _
         "Service Commander " & Rev)
   End If

End Sub

Public Sub GetServiceList(Optional LastIndex As Long = -1)

  Dim a As Long
  Dim AltDel As Long
  Dim Del As Long
  Dim ShowService As Long
  Dim State As Long

   Command3(0).Enabled = True
   Command2.Enabled = True
   Frame3.Visible = False

   For a = 0 To 5
      'reset the labels to no sorting
      Label1(a).Tag = ""
      Label1(a).ToolTipText = "Right Click to Sort"
      Label1(a).FontItalic = False

      'clear the lists before we fill them
      List1(a).Clear
   Next

   Select Case Combo2.ListIndex
      'state lets you choose what service states get displayed
    Case -1, 0: 'All States
      State = 3

    Case 1: 'Active State
      State = 1

    Case 2: 'Inactive State
      State = 2
   End Select

   Select Case Combo1.ListIndex
      'delimeter and altdelimeter lets you choose what type of services are displayed
    Case -1, 0: 'All Services
      Del = 0
      AltDel = 0
      ShowService = SERVICE_DRIVER Or SERVICE_WIN32

    Case 1: 'All Drivers
      Del = 0
      AltDel = 0
      ShowService = SERVICE_DRIVER

    Case 2: 'All Win 32 Services
      Del = 0
      AltDel = 0
      ShowService = SERVICE_WIN32

    Case 3: 'Kernel Mode  Driver Services
      Del = 1
      AltDel = 0
      ShowService = SERVICE_DRIVER

    Case 4: 'File System Driver Services
      Del = 2
      AltDel = 0
      ShowService = SERVICE_DRIVER

    Case 5: 'Adapter Services
      Del = 4
      AltDel = 0
      ShowService = SERVICE_DRIVER Or SERVICE_WIN32

    Case 6: 'Driver Services
      Del = 8
      AltDel = 0
      ShowService = SERVICE_DRIVER Or SERVICE_WIN32

    Case 7: 'Win32 Own Process Services
      Del = 10
      AltDel = 16
      ShowService = SERVICE_WIN32

    Case 8: 'Win 32 Shared Process Services
      Del = 20
      AltDel = 32
      ShowService = SERVICE_WIN32

    Case 9: 'Own Interactive Process Services
      Del = 100
      AltDel = 272
      ShowService = SERVICE_WIN32

    Case 10: 'Shared Interactive Process Services
      Del = 120
      AltDel = 288
      ShowService = SERVICE_WIN32
   End Select

   'load the services into the listboxes
   Call EnumServices(List1(0), List1(1), List1(2), List1(3), List1(4), List1(5), Del, AltDel, _
      ShowService, State)

   Command1.Caption = "&Refresh  (" & List1(1).ListCount & " Items)" 'makes it say refresh instead
   '   of get services

   If LastIndex <> -1 And LastIndex <= List1(0).ListCount Then

      'selects the item that was selected before refresh

      For a = 0 To 5
         List1(a).Selected(LastIndex) = True
      Next

      'List1(1).Selected(LastIndex) = True
      'List1(2).Selected(LastIndex) = True
      'List1(3).Selected(LastIndex) = True
      'List1(4).Selected(LastIndex) = True
      'List1(5).Selected(LastIndex) = True
   End If

   VScroll1.Max = List1(0).ListCount - 1
   If VScroll1.Max < 0 Then VScroll1.Max = 0
   VScroll1.Min = 0
   If Form1.Enabled = True Then Command1.SetFocus

End Sub

Private Sub HScroll1Frame()

   Frame1.Left = 0 - HScroll1.Value

   If Frame2.Width > Form1.Width Then
      If Frame2.Width + 0 - HScroll1.Value < Form1.Width Then
         Frame2.Left = Form1.Width - Frame2.Width - 100
       Else
         Frame2.Left = 0 - HScroll1.Value
      End If

    Else
      Frame2.Left = 0
   End If

   Frame5.Left = Frame2.Left + 3090
   Frame5.Top = Frame2.Top

   If HScroll1.Value < HScroll1.Max - 50 And Frame1.Width + Frame1.Left > Form1.Width Then
      VScroll1.Visible = True
    Else
      VScroll1.Visible = False
   End If

End Sub

Private Sub HScroll1_Change()

   'move the frames to the final position on change
   'dont need it for the mouse to move scrollbar
   'but we do for keyboard
   HScroll1Frame

End Sub

Private Sub HScroll1_Scroll()

   'move the frames on scroll
   HScroll1Frame

End Sub

Private Sub Label1_MouseDown(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

  Dim a As Long

   'make it so each time you right click a label it makes it a different sort type

   If Button = 2 Then

      Select Case Label1(Index).Tag
       Case "Sorted Asc":
         Label1(Index).Tag = "Sorted Dsc"
         Label1(Index).ToolTipText = "Sorted Descending"
         Label1(Index).FontItalic = True

       Case "Sorted Dsc":
         Label1(Index).Tag = ""
         Label1(Index).ToolTipText = "Right Click to Sort"
         Label1(Index).FontItalic = False

       Case Else:
         Label1(Index).Tag = "Sorted Asc"
         Label1(Index).ToolTipText = "Sorted Ascending"
         Label1(Index).FontItalic = True
      End Select

      For a = 0 To 5

         If Index <> a Then
            Label1(a).Tag = ""
            Label1(a).ToolTipText = "Right Click to Sort"
            Label1(a).FontItalic = False
            'reset the other labels to no sorting
         End If

      Next

      SortLists (Index)

   End If

End Sub

Private Sub Label1_MouseMove(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

   If X > Label1(Index).Width - 150 Then
      Label1(Index).MousePointer = 9
      'Give a Resize arrow towards the end of the label
    Else
      Label1(Index).MousePointer = vbNormal
      'Give a Normal arrow towards the beginning of the label
   End If

   If Button = 1 And (X > Label1(Index).Width - 150 Or Command6.Visible = True) Then
      If Label1(Index).Left + (X - Label1(Index).Left) > 50 Then
         Label1(Index).Width = (Label1(Index).Left + (X - Label1(Index).Left))
         'command 6 is my rudimentary resize guide
         ' it also triggers the mouse up event to resize upon completion
         Command6.Left = Label1(Index).Width + Label1(Index).Left
         Command6.Height = Frame1.Height
         Command6.Visible = True
         Command6.Tag = Index
       Else
         Label1_MouseUp 0, 0, 0, X, Y
         'had a problem with the mouse moving out of the movement
         'area and not triggering the mouse up event
      End If

    Else
      Label1_MouseUp 0, 0, 0, X, Y
      'had a problem with the mouse moving out of the movement
      'area and not triggering the mouse up event
   End If

End Sub

Private Sub Label1_MouseUp(Index As Integer, _
                           Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

  Dim a As Long

   'only if command6 is visible from the mouse move event does this resize the list boxes
   'command6.tag is the label index from the mouse move

   If Command6.Visible = True Then
      If Label1(Command6.Tag).Left + (X - Label1(Command6.Tag).Left) > 25 Then
         List1(Command6.Tag).Width = (Label1(Command6.Tag).Left + (X - Label1(Command6.Tag).Left)) _
            + 285
         '285 is the overlap of one listbox to the next to hide the scrollbar (i know cheap!)
         If Command6.Tag <> 5 Then List1(Command6.Tag + 1).Left = List1(Command6.Tag).Left + _
            List1(Command6.Tag).Width - 285

         For a = 0 To 5
            Label1(a).Left = List1(a).Left

            If a < 5 Then
               Label1(a).Width = List1(a).Width - 285
               List1(a + 1).Left = List1(a).Left + List1(a).Width - 285
               ' Match up all the edges of labels and lists (except 5 it has nothing overlapping it)
             Else
               Label1(a).Width = List1(a).Width
            End If

         Next

         Frame1.Width = List1(5).Left + List1(5).Width
         Frame3.Width = Frame1.Width

         Form_Resize

         Command6.Visible = False
         'hide command6 when we are done so it gets it out of the way and
         'so we don't do it again before its time
      End If

   End If

End Sub

Public Function LastIndex() As Long

   'I used this routine so much I found it easier just to make it a function
   ' it just gets the selected item number if there is one selected
  Dim a As Long
  Dim count As Long

   For a = 0 To 5
      count = count + List1(a).SelCount
   Next

   If List1(1).ListIndex <> -1 And count > 4 Then
      LastIndex = List1(0).ListIndex
    Else
      LastIndex = -1
   End If

End Function

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

   'On Error Resume Next
   'This code added per error pointed out by The_Jokez concerning arrow keys not
   'scrolling and selecting
  Dim Lindex As Long
  Dim a As Long

   'ok since we're using the keydown event everything that happens with a listbox
   'happens on the key up event so we need to figure out the key pressed
   'and add it to the current listindex without changing it or we'll be off one

   '?why not use the key up event instead and save some code you ask?
   ' the lists won't scroll together on the key up event
   Lindex = List1(Index).ListIndex

   If KeyCode = 38 Or KeyCode = 37 Then

      '38=up
      '37=left (also moves up)
      'List1(Index).Selected(Lindx) = False
      'UP
      'up and left arrows move the list up
      If Lindex > 0 Then Lindex = Lindex - 1
      ' we don't want the lindx to be less than 0 or we'll error

      SelectItem Index, Lindex, LastSel
      'we'll move all lists except the list that is recieving the key event
      'otherwise we will end up 1 off since lindx=listindex-1
   End If

   If KeyCode = 40 Or KeyCode = 39 Then
      '40=dn
      '39 = right (also moves dn)
      'Down
      'down and right arrows move the list down
      If Lindex < List1(Index).ListCount - 1 Then Lindex = Lindex + 1
      ' we don't want the lindx to be more than listcount or we'll error

      SelectItem Index, Lindex, LastSel
      'we'll move all lists except the list that is recieving the key event
      'otherwise we will end up 1 off since lindx=listindex+1
   End If

   'I really hated the annoying blinking the scrollbar did so I set tab stop to false
   'but this makes people that use only a keyboard SOL so i made the arrow keys also
   'move the scrollbar when the lists are selected

   If KeyCode = 37 Then

      If HScroll1.Value - 100 < HScroll1.Min Then
         HScroll1.Value = HScroll1.Min
       Else
         HScroll1.Value = HScroll1.Value - 100
      End If

   End If

   If KeyCode = 39 Then
      If HScroll1.Value + 100 > HScroll1.Max Then
         HScroll1.Value = HScroll1.Max
       Else
         HScroll1.Value = HScroll1.Value + 100
      End If

   End If
   'Make sure all the top items match

   For a = 0 To 5
      If Lindex <> a Then List1(a).TopIndex = List1(Index).TopIndex
   Next

End Sub

Private Sub List1_MouseDown(Index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

  Dim Lindex As Long

   On Error GoTo exitmnu
   Lindex = ListBoxLocation(List1(Index), Y)
   'get the item number the nmouse is over
   Text1.Text = ListBoxToolTip(List1(Index), Y, "")
   'put the item text into a textbox so you can copy & paste the name or desc.

   SelectItem -1, Lindex, LastSel
   LastCol = Index
   'here is the more info menu popup
   If Button = 2 And Lindex > -1 Then PopupMenu MainMenu, , X + List1(Index).Left, Y

   'again it only shows if something is selected and you right click

exitmnu:

End Sub

Private Sub List1_MouseMove(Index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

   ListBoxToolTip List1(Index), Y, Label1(Index).Caption
   ' get the item the mouse is moving over a popup
   'the text in the list for the tool tip

End Sub

Private Sub List1_MouseUp(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

  Dim Lindex As Long

   On Error GoTo exitmnu
   Lindex = ListBoxLocation(List1(Index), Y)
   'get the item number the nmouse is over
   Text1.Text = ListBoxToolTip(List1(Index), Y, "")
   'put the item text into a textbox so you can copy & paste the name or desc.
   SelectItem -1, Lindex, LastSel

exitmnu:

End Sub

Private Function ListBoxLocation(ByRef objList As Object, ByVal Y As Single) As Long

   'gets the index of the item the mouse is over
  Dim lngIndex      As Long
  Dim objParentFont As StdFont

   With objList

      ' Font adaptation & Index determination:

      With .Parent
         Set objParentFont = .Font
         Set .Font = objList.Font
         lngIndex = Y \ (.TextHeight("") + 30)
         Set .Font = objParentFont
      End With

      'If we take the font size and the top index we can figure
      'out what item the mouse is over with simple math
      ' Index evaluation:
      lngIndex = lngIndex + .TopIndex

      If lngIndex < .ListCount Then
         ListBoxLocation = lngIndex
       Else
         ListBoxLocation = -1
      End If

   End With

End Function

Private Function ListBoxToolTip(ByRef objList As Object, ByVal Y As Single, Optional ByRef _
                                DefaultToolTip As String = "") As String

   'Dim a As Long
  Dim lngIndex      As Long
  Dim objParentFont As StdFont

   With objList
      'gets the item the mouse is over and makes the tooltip its data
      ' Font adaptation & Index determination:

      With .Parent
         Set objParentFont = .Font
         Set .Font = objList.Font

         lngIndex = Y \ (.TextHeight("") + 30)
         Set .Font = objParentFont
      End With

      'If we take the font size and the top index we can figure
      'out what item the mouse is over with simple math
      ' Index evaluation:
      lngIndex = lngIndex + .TopIndex

      If lngIndex < .ListCount Then
         .ToolTipText = .List(lngIndex) & " (Right Click For More Info And Settings)"
         ListBoxToolTip = .List(lngIndex)
       Else
         .ToolTipText = DefaultToolTip
         'if there isn't an item we'll make it the default tooltip supplied (Label1(Index).Caption)
      End If

   End With

End Function

Private Sub moreinfo_Click(Index As Integer)

   'Popup Moreinfo here
   Form2.SettingsOnly = False
   Form2.SvcName = List1(1).List(List1(Index).ListIndex)
   Load Form2
   Form2.Show

End Sub

Private Sub SearchColumn_Click()

   Text2.Text = ""
   Label9.Caption = ""
   StartPos = -1
   Frame5.Caption = "Search Column " & Label1(LastCol).Caption
   Frame5.Visible = True
   Frame5.Left = Frame2.Left + 3090
   Frame5.Top = Frame2.Top

End Sub

Private Sub SearchList(Start As Long)
On Error Resume Next
  Dim Lindex As Long

   Lindex = SendMessage(List1(LastCol).Hwnd, LB_FINDSTRING, Start, ByVal Text2.Text)
   StartPos = Lindex

   If Lindex > -1 Then
      'if there is an item in the list select it

      Label9.Caption = "Item Found"

      SelectItem -1, Lindex, LastSel
    Else
      Label9.Caption = "Item Not Found"
   End If

End Sub

Private Sub SelectItem(Index As Integer, Selected As Long, ByRef LastSelected As Long)
On Error Resume Next
  Dim a As Long

   If Selected > -1 Then
      'if there is an item in the list under the mouse select it
      ' I actually chose to do this this way because
      'by default a listbox doesn't select with a right click

      For a = 0 To 5

         List1(a).Selected(LastSelected) = False

      Next

      For a = 0 To 5

         If Index <> a Then List1(a).Selected(Selected) = True

      Next
      LastSelected = Selected
      'put the selected item back into the LastSelected item for next time
   End If

End Sub

Private Sub SetStoredConfig(Filename As String)

   'this is the part i really love
   'the ability to restore a service profile
  Dim a As Long
  Dim Changed As String
  Dim NumChg As Long
  Dim Errors As String
  Dim NumErr As Long
  Dim FCnt As Long
  Dim lpAppName As String
  Dim lpBuffer As String
  Dim lpKeyName As String
  Dim lReturn As Long
  Dim NewStartType As Long
  Dim NotFound As String
  Dim NumNotFnd As Long
  Dim NoChange As String
  Dim NumNoChg As Long
  Dim nSize As Long
  Dim PbIntv As Long
   NumChg = 0
   NumErr = 0
   NumNotFnd = 0
   NumNoChg = 0
   FCnt = 0
   'hide frame 1 because it blinks while in refresh from the combo selection changing
   Frame1.Visible = False
   Combo1.ListIndex = 0
   Combo2.ListIndex = 0
   'disable interaction except the progress bar and cancel button
   Frame2.Enabled = False
   Frame4.Top = (Me.Height - Frame4.Height) / 2
   Frame4.Left = (Me.Width - Frame4.Width) / 2
   Shape1(1).BackColor = &HFF00& 'neon green
   Frame4.Visible = True
   Form1.MousePointer = vbHourglass
   Shape1(1).Width = 0
   CancelOpp = False
   'set canceled to false so we know when someone clicks it

   If List1(1).ListCount - 1 > 0 And List1(1).ListCount - 1 < Shape1(0).Width Then
      PbIntv = (Shape1(0).Width / List1(1).ListCount - 1) + 1
      Shape1(1).Width = PbIntv
      'get the how many times we need to move the bar to = 100%
      'this is only good up to 3600 items ( width of shape1(0))
      ' I don't see a pc having more than 3600 services and if so
      'they are used to waiting so they don't need a progress bar anyway
   End If

   Errors = ""
   Changed = ""
   NoChange = ""
   NotFound = ""
   nSize = 255

   For a = 0 To List1(1).ListCount - 1
      If a > 0 Then Shape1(1).Width = PbIntv + (Shape1(1).Width)
      If Shape1(1).Width > Shape1(0).Width Then Shape1(1).Width = Shape1(0).Width
      Label5(0).Caption = "Restoring Service " & a + 1 & " of " & List1(1).ListCount
      'increase the progressbar and item counts
      DoEvents
      lpBuffer = Space(255)
      lpAppName = List1(1).List(a) 'The section of the INI...
      Label5(1).Caption = lpAppName ' shows the service name in progressbar
      lpKeyName = "Description" ' THE LINE OF THE INI
      GetPrivateProfileString lpAppName, lpKeyName, 0&, lpBuffer, nSize, Filename
      'Get ini description and put it into lpbuffer

      If Mid(LCase$(RTrim$(lpBuffer)), 1, Len(RTrim$(lpBuffer)) - 1) = LCase$(List1(2).List(a)) Then
         FCnt = FCnt + 1
         lpBuffer = Space(255)
         lpKeyName = "Start State" ' THE LINE OF THE INI
         GetPrivateProfileString lpAppName, lpKeyName, 0&, lpBuffer, nSize, Filename
         ' used lcase$ so that the names and start state were not case sensitive

         If Mid(LCase$(RTrim$(lpBuffer)), 1, Len(RTrim$(lpBuffer)) - 1) <> LCase$(List1(5).List(a)) _
            Then
            'Get the number for the start state

            Select Case Mid(LCase$(RTrim$(lpBuffer)), 1, Len(RTrim$(lpBuffer)) - 1)
             Case "boot":
               NewStartType = SERVICE_BOOT_START

             Case "system":
               NewStartType = SERVICE_SYSTEM_START

             Case "automatic":
               NewStartType = SERVICE_AUTO_START

             Case "disabled":
               NewStartType = SERVICE_DISABLED

             Case "manual":
               NewStartType = SERVICE_DEMAND_START

             Case Else:
               'if the start state doesn't match we'll log an error
               '(i use (N/A) for services i don't want to change)
               NewStartType = SERVICE_NO_CHANGE
               Errors = Errors & vbCrLf & lpAppName & " / Start Type Not Recognized: " & _
                  Mid(LCase$(RTrim$(lpBuffer)), 1, Len(RTrim$(lpBuffer)) - 1)
            End Select

            'set the start state for the specified service
            lReturn = SetServiceConfig(List1(1).List(a), SERVICE_NO_CHANGE, NewStartType, _
               SERVICE_NO_CHANGE, 0&)

            If lReturn <> 0 Then

               NumErr = NumErr + 1

               Errors = Errors & vbCrLf & lpAppName & " / " & ErrLib(lReturn)
               Shape1(1).BackColor = &HFF& 'make progressbar red to signify error
               'if we get an error returned log it

             Else
               NumChg = NumChg + 1
   
               Changed = Changed & vbCrLf & lpAppName & " / Changed State - " & List1(5).List(a) & _
                  " To " & Mid(RTrim$(lpBuffer), 1, Len(RTrim$(lpBuffer)) - 1)
               'log the name and the orig start and what it was changed to
            End If

          Else
  
            NumNoChg = NumNoChg + 1
            NoChange = NoChange & vbCrLf & lpAppName & " / Matched Current State - " & _
               List1(5).List(a)
            'log that there wasn't a change in start state
         End If

       Else

         NumNotFnd = NumNotFnd + 1

         NotFound = NotFound & vbCrLf & List1(2).List(a)
      End If

      If CancelOpp = True Then

         NumErr = NumErr + 1

         Errors = Errors & vbCrLf & "User Canceled"
         'log an error that the user cancled
         Exit For
      End If

   Next a
   ' used form 2 to display the results since i already had a nice text display there
   lpAppName = "Scmdr Profile Data 0.1.4" 'The section of the INI...
   'left this 0.1.4 for backward compability
   lpBuffer = Space(255)
   lpKeyName = "Entries" ' THE LINE OF THE INI
   GetPrivateProfileString lpAppName, lpKeyName, "N/A", lpBuffer, nSize, Filename
   Errors = "Errors: " & NumErr & vbCrLf & Errors
   Changed = "Changed: " & NumChg & vbCrLf & Changed
   NoChange = "No Change: " & NumNoChg & vbCrLf & NoChange
   NotFound = "Not Found: " & NumNotFnd & vbCrLf & NotFound
   
   Load Form2
   Form2.Caption = "Set Configuration Results"
   Form2.Frame1.Visible = False
   Form2.Text1.Text = vbNullString
   Form2.Text1.Text = "Found: " & FCnt & "/" & Mid(Trim$(lpBuffer), 1, Len(Trim$(lpBuffer)) - 1) & _
      vbCrLf & vbCrLf & Errors & vbCrLf & vbCrLf & Changed & vbCrLf & vbCrLf & NoChange & vbCrLf & _
      vbCrLf & NotFound & vbCrLf
   'put all the strings together with carriage return for a nice log format
   Form2.Show

   If InStr(1, Command$, "load", vbTextCompare) <> 0 Then
      Form2.Timer1.Tag = 0
      Form2.Timer1.Enabled = True
      Form2.Command2(2).ToolTipText = "Click to Close, Right Click to Stop Timer"

      ' if you used the unattended command then the log will close in 30 seconds
    Else
      ' make everything display as it was before you restored a profile
      Form1.MousePointer = vbNormal
      GetServiceList
      Frame1.Visible = True
      Frame2.Enabled = True
      Frame4.Visible = False
   End If

End Sub

Private Sub SortLists(Index As Long)

  Dim a As Long
  Dim lCount As Long
  Dim listM() As String
  Dim Jump As Long
  Dim Temp As String
  Dim Swapped As Boolean
  Dim i As Long

   'Code Doc's sort Modified for use in list matrix setup
   'http://www.vbforums.com/showthread.php?referrerid=61394&t=471565

   ReDim listM(5, List1(0).ListCount) 'redim it to the number of items in list1(0)
   lCount = List1(0).ListCount - 1

   For a = 0 To lCount
      'Load the array with values
      listM(0, a) = List1(0).List(a)
      listM(1, a) = List1(1).List(a)
      listM(2, a) = List1(2).List(a)
      listM(3, a) = List1(3).List(a)
      listM(4, a) = List1(4).List(a)
      listM(5, a) = List1(5).List(a)
   Next

   Jump = lCount

   While Jump
      Jump = Jump \ 2
      Swapped = True
      'make the list iteration half the size each time to make sorting faster

      While Swapped
         Swapped = False

         For i = 0 To lCount - Jump

            If LCase$(listM(Index, i)) > LCase$(listM(Index, i + Jump)) Then
               'used lcase$ to make the sort case insensitive

               For a = 0 To 5
                  Temp = listM(a, i)
                  listM(a, i) = listM(a, i + Jump)
                  listM(a, i + Jump) = Temp
                  Swapped = True
               Next

            End If
         Next

      Wend
   Wend

   For a = 0 To 5
      List1(a).Clear
   Next

   Select Case Label1(Index).Tag
    Case "Sorted Asc":

      For a = 0 To lCount
         'Sorted ascending
         'Put sorted array data back into list
         List1(0).List(a) = listM(0, a)
         List1(1).List(a) = listM(1, a)
         List1(2).List(a) = listM(2, a)
         List1(3).List(a) = listM(3, a)
         List1(4).List(a) = listM(4, a)
         List1(5).List(a) = listM(5, a)
      Next

    Case "Sorted Dsc":

      For a = lCount To 0 Step -1
         'Sorted descending (flipped backwards)
         'Put sorted array data back into list
         List1(0).List(lCount - a) = listM(0, a)
         List1(1).List(lCount - a) = listM(1, a)
         List1(2).List(lCount - a) = listM(2, a)
         List1(3).List(lCount - a) = listM(3, a)
         List1(4).List(lCount - a) = listM(4, a)
         List1(5).List(lCount - a) = listM(5, a)
      Next

    Case Else:
      'no sorting
      GetServiceList
   End Select

End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = 2 And Text1.SelLength = 0 Then
      Text1.SelStart = 0
      Text1.SelLength = Len(Text1.Text)
   End If

   'if text isn't already selected on right click it will select all the text for you

End Sub

Private Sub Text2_Change()

   SearchList -1

End Sub

Private Sub Timer1_Timer()

   'lowered the interval for better scrolling
  Dim a As Long
  Dim b As Long
  Dim b2 As Long
  Dim fstr As String
  Dim MatchLoc As Long
  Dim match As Long
  Dim tempstr As String
  Dim tempstr2 As String
  Const ListBxTotal = 5

   On Error GoTo ExitTimer
   'this one is a bit complicated
   'i couldn't find a function to compare 6 numbers and find the different one
   'so i put a string around each one and used instr to compare

   For a = 0 To ListBxTotal
      tempstr = tempstr & "A" & Str$(List1(a).TopIndex) & "Z"
      b = b + List1(a).TopIndex
      ' I did this so that 2 indexes at 1 didn't look like one index at 11
      tempstr2 = tempstr2 & "A" & Str$(List1(a).ListIndex) & "Z"
      b2 = b2 + List1(a).ListIndex
      ' i compare the topindex and the list index
      ' so we can have the right matching and the right item selected
   Next

   If b / (ListBxTotal + 1) <> List1(0).TopIndex Then
      'this is a simple way to find out if the index are at the same position
      'if in 5 boxes top index is at 3 then 3x5=15 and 15/5=3 which is the topindex

      For a = 0 To ListBxTotal
         match = 0

         fstr = "A" & Str$(List1(a).TopIndex) & "Z"

         MatchLoc = InStr(1, tempstr, fstr)

         If MatchLoc <> 0 Then
            'if we found a match increase match by one
            match = match + 1
            If InStr(MatchLoc + Len(fstr), tempstr, fstr) <> 0 Then match = match + 1
         End If

         If match = 1 Then
            'if we only find one match then we know this is the item that
            'doesn't match the rest

            For b = 0 To ListBxTotal
               If a <> b Then List1(b).TopIndex = List1(a).TopIndex
            Next b

            Exit For
         End If

         'otherwise try again
      Next a

    Else
   End If

   If b2 / (ListBxTotal + 1) <> List1(0).ListIndex Then
      'this is the same as above but its for the matching of selected items

      For a = 0 To ListBxTotal
         match = 0

         fstr = "A" & Str$(List1(a).ListIndex) & "Z"

         MatchLoc = InStr(1, tempstr2, fstr)

         If MatchLoc <> 0 Then
            match = match + 1
            If InStr(MatchLoc + Len(fstr), tempstr2, fstr) <> 0 Then match = match + 1
         End If

         If match = 1 Then

            For b = 0 To ListBxTotal
               If a <> b Then List1(b).ListIndex = List1(a).ListIndex
            Next b

            Exit For
         End If

      Next a
    Else

   End If
   Exit Sub
ExitTimer:
   Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()

   'this finds if the service state has changed to the one we want
   'if so it updates the listbx and stops the timer
   On Error GoTo ExitTimer

   'Debug.Print GetServiceStatus(SvcName) & "/" & DesiredState

   If DesiredState = GetServiceStatus(SvcName) Then

      Timer2.Tag = 0
      Timer2.Enabled = False

      If List1(1).ListIndex <> -1 And List1(4).ListIndex = CurrentItem Then
         List1(4).List(CurrentItem) = SvcState(GetServiceStatus(SvcName))

       Else

         GetServiceList List1(1).ListIndex
      End If

    Else
      Timer2.Tag = Timer2.Tag + 1
      If Timer2.Tag >= 31 Then Timer2.Enabled = False
      'after 30 seconds we will stop the timer because we don't want it
      'running forever waiting on something that will never happen
      'shouldn't happen but you never know and its better to be safe than sorry
   End If

   Exit Sub
ExitTimer:
   'on error reset timer
   Timer2.Enabled = False
   Timer2.Tag = 0

End Sub

Private Sub VScroll1_Change()

   On Error Resume Next
   List1(5).TopIndex = VScroll1.Value
   'the top index is a bit off from the min and max values and
   'i can't easily find how many pages of data there are so it
   'goes on items instead and this will throw an error because
   'topindex doesn't go as high as the item count

End Sub

Private Sub VScroll1_Scroll()

   On Error Resume Next
   List1(5).TopIndex = VScroll1.Value

End Sub

Private Sub WriteCurConfig(Filename As String)

   ' writes an ini file with your current configuration
   On Error GoTo WriteError
  Dim lpAppName As String
  Dim lpKeyName As String
  Dim lpBuffer As String

  Dim a As Long
  Dim PbIntv As Long

   If LCase$(Dir(Filename)) = LCase$(Mid(Filename, InStrRev(Filename, "\") + 1)) Then

      Select Case MsgBox("File '" & Mid(Filename, InStrRev(Filename, "\") + 1) & "' already exists" _
         & vbCrLf & "Are you sure you want to overwrite this file with your current profile?", 292, _
         "Service Commander " & Rev)
       Case 6:
         'Yes Button Selected
         Kill Filename 'delete the file so we don't get conflicting profiles

       Case 7:
         'No Button Selected
         Exit Sub
      End Select

   End If
   'make everything but the progressbar disabled
   Frame1.Enabled = False
   Frame2.Enabled = False
   Frame4.Top = (Me.Height - Frame4.Height) / 2
   Frame4.Left = (Me.Width - Frame4.Width) / 2
   Frame4.Visible = True
   Shape1(1).BackColor = &HFF00& 'neon green
   Form1.MousePointer = vbHourglass
   Shape1(1).Width = 0
   CancelOpp = False

   If List1(1).ListCount - 1 > 0 And List1(1).ListCount - 1 < Shape1(0).Width Then
      PbIntv = (Shape1(0).Width / List1(1).ListCount - 1) + 1
      Shape1(1).Width = PbIntv
   End If

   lpAppName = "Scmdr Profile Data 0.1.4" 'The section of the INI...
   'left this 0.1.4 for backward compability
   lpKeyName = "Date" ' THE LINE OF THE INI
   lpBuffer = Date$
   WritePrivateProfileString lpAppName, lpKeyName, lpBuffer, Filename
   lpKeyName = "Time" ' THE LINE OF THE INI
   lpBuffer = Time$
   WritePrivateProfileString lpAppName, lpKeyName, lpBuffer, Filename
   lpKeyName = "Entries" ' THE LINE OF THE INI
   lpBuffer = List1(1).ListCount
   WritePrivateProfileString lpAppName, lpKeyName, lpBuffer, Filename
   lpKeyName = "Extra Info" ' THE LINE OF THE INI
   lpBuffer = Combo1.Text & " - " & Combo2.Text
   WritePrivateProfileString lpAppName, lpKeyName, lpBuffer, Filename

   For a = 0 To List1(1).ListCount - 1
      'go through the list and save each service in the list to the ini
      If a > 0 Then Shape1(1).Width = PbIntv + (Shape1(1).Width)
      If Shape1(1).Width > Shape1(0).Width Then Shape1(1).Width = Shape1(0).Width
      Label5(0).Caption = "Saving Service " & a + 1 & " of " & List1(1).ListCount
      DoEvents
      lpAppName = List1(1).List(a) 'The section of the INI...
      Label5(1).Caption = lpAppName 'shows service name in progressbar
      lpKeyName = "Description" ' THE LINE OF THE INI
      lpBuffer = List1(2).List(a)
      WritePrivateProfileString lpAppName, lpKeyName, lpBuffer, Filename

      lpKeyName = "Start State" ' THE LINE OF THE INI
      lpBuffer = List1(5).List(a)
      WritePrivateProfileString lpAppName, lpKeyName, lpBuffer, Filename
      If CancelOpp = True Then Exit For
   Next

   'put everything back as it was
   Frame1.Enabled = True
   Frame2.Enabled = True
   Frame4.Visible = False

   Form1.MousePointer = vbNormal
   Exit Sub
WriteError:
   If Err.Number = 53 Then Resume Next

   If Err.Number = 75 Then
      Call MsgBox("File Read Only" & vbCrLf & "Number=" & Err.Number, 48, "Write Current" & _
         " Configuration Error")
      Exit Sub
   End If

   Call MsgBox(Err.Description & vbCrLf & "Number=" & Err.Number, 48, "Write Current Configuration" & _
      " Error")
   Frame1.Enabled = True
   Frame2.Enabled = True
   Frame4.Visible = False

   Form1.MousePointer = vbNormal

End Sub

