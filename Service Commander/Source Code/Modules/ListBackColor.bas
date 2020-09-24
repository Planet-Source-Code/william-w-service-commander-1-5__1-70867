Attribute VB_Name = "ListBackColor"
Option Explicit

Private Type ClrList
   Running As Long
   Stopped As Long
   Paused As Long
   Starting As Long
   Stopping As Long
   Selected As Long
   SelectedText As Long
   Kernel As Long
   FileSys As Long
   Adapter As Long
   Driver As Long
   Automatic As Long
   Boot As Long
   System As Long
   Manual As Long
   Disabled As Long
   OwnProcess As Long
   SharedProcess As Long
   OwnInteractive As Long
   SharedInteractive As Long
   TextColor As Long
   BackColor As Long
   ColoredText As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   ItemData As Long
End Type

'-----------------API DECLARES---------------------
'C
Private Declare Function CallWindowProc Lib "user32" _
      Alias "CallWindowProcA" ( _
      ByVal lpPrevWndFunc As Long, _
      ByVal Hwnd As Long, _
      ByVal Msg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
      Alias "RtlMoveMemory" ( _
      Destination As Any, _
      Source As Any, _
      ByVal Length As Long)

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'D
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

'F
Private Declare Function FillRect Lib "user32" ( _
      ByVal hdc As Long, _
      lpRect As RECT, _
      ByVal hBrush As Long) As Long

'G
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'S
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Declare Function SendMessage Lib "user32" _
      Alias "SendMessageA" ( _
      ByVal Hwnd As Long, _
      ByVal wMsg As Long, _
      ByVal wParam As Long, _
      lParam As Any) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
      Alias "SetWindowLongA" ( _
      ByVal Hwnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long) As Long

'T
Private Declare Function TextOut Lib "gdi32" _
      Alias "TextOutA" ( _
      ByVal hdc As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal lpString As String, _
      ByVal nCount As Long) As Long

Private Const LB_GETTEXT = &H189
Private Const ODS_FOCUS = &H10
Private Const ODS_SELECTED As Long = &H1
Private Const ODT_LISTBOX = 2
Private Const WM_DRAWITEM = &H2B

Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWTEXT = 8
Public Const GWL_WNDPROC = (-4)

Public ColorList As ClrList
Public lPrevWndProc As Long


Public Function ColorListNum(Item As Long, Optional Data As Long = -1) As Long

   Select Case Item
    Case 0: 'Running

      If Data = -1 Then
         ColorListNum = ColorList.Running
       Else
         ColorList.Running = Data
      End If

    Case 1: 'Stopped

      If Data = -1 Then
         ColorListNum = ColorList.Stopped
       Else
         ColorList.Stopped = Data
      End If

    Case 2: 'Paused

      If Data = -1 Then
         ColorListNum = ColorList.Paused
       Else
         ColorList.Paused = Data
      End If

    Case 3: 'Starting

      If Data = -1 Then
         ColorListNum = ColorList.Starting
       Else
         ColorList.Starting = Data
      End If

    Case 4: 'Stopping

      If Data = -1 Then
         ColorListNum = ColorList.Stopping
       Else
         ColorList.Stopping = Data
      End If

    Case 5: 'Selected

      If Data = -1 Then
         ColorListNum = ColorList.Selected
       Else
         ColorList.Selected = Data
      End If

    Case 6: 'Kernel Mode

      If Data = -1 Then
         ColorListNum = ColorList.Kernel
       Else
         ColorList.Kernel = Data
      End If

    Case 7: 'File System

      If Data = -1 Then
         ColorListNum = ColorList.FileSys
       Else
         ColorList.FileSys = Data
      End If

    Case 8: 'Adapter

      If Data = -1 Then
         ColorListNum = ColorList.Adapter
       Else
         ColorList.Adapter = Data
      End If

    Case 9: 'Driver

      If Data = -1 Then
         ColorListNum = ColorList.Driver
       Else
         ColorList.Driver = Data
      End If

    Case 10: 'Automatic

      If Data = -1 Then
         ColorListNum = ColorList.Automatic
       Else
         ColorList.Automatic = Data
      End If

    Case 11: 'Boot

      If Data = -1 Then
         ColorListNum = ColorList.Boot
       Else
         ColorList.Boot = Data
      End If

    Case 12: 'System

      If Data = -1 Then
         ColorListNum = ColorList.System
       Else
         ColorList.System = Data
      End If

    Case 13: 'Manual

      If Data = -1 Then
         ColorListNum = ColorList.Manual
       Else
         ColorList.Manual = Data
      End If

    Case 14: 'Disabled

      If Data = -1 Then
         ColorListNum = ColorList.Disabled
       Else
         ColorList.Disabled = Data
      End If

    Case 15: 'Win 32 Own Process

      If Data = -1 Then
         ColorListNum = ColorList.OwnProcess
       Else
         ColorList.OwnProcess = Data
      End If

    Case 16: 'Win 32 Shared Process

      If Data = -1 Then
         ColorListNum = ColorList.SharedProcess
       Else
         ColorList.SharedProcess = Data
      End If

    Case 17: 'Own Interactive Process

      If Data = -1 Then
         ColorListNum = ColorList.OwnInteractive
       Else
         ColorList.OwnInteractive = Data
      End If

    Case 18: 'Shared Interactive Process

      If Data = -1 Then
         ColorListNum = ColorList.SharedInteractive
       Else
         ColorList.SharedInteractive = Data
      End If

    Case 19: 'Text Color

      If Data = -1 Then
         ColorListNum = ColorList.TextColor
       Else
         ColorList.TextColor = Data
      End If

    Case 20: 'Back Color

      If Data = -1 Then
         ColorListNum = ColorList.BackColor
       Else
         ColorList.BackColor = Data
      End If

    Case 21: 'Selected Text Color

      If Data = -1 Then
         ColorListNum = ColorList.SelectedText
       Else
         ColorList.SelectedText = Data
      End If

    Case 22: 'Use Colored Text

      If Data = -1 Then
         ColorListNum = ColorList.ColoredText

       Else
         Debug.Print Data
         ColorList.ColoredText = Data
      End If

   End Select

End Function

Public Sub GetColors()
On Error GoTo GetColorErr
  Dim Data As Long
  Dim a As Long

   If GetSetting("SerVice Commander\" & App.Path, "Options", "Color Set", "NO") <> "YES" Then
      'Set Color Defaults if none exist
      ColorListNum 0, &HC000&   'Running
      ColorListNum 1, &HC0&     'Stopped
      ColorListNum 2, &H80FF&   'Paused
      ColorListNum 3, &HC0FFC0  'Starting
      ColorListNum 4, &HC0C0FF  'Stopping
      ColorListNum 5, GetSysColor(COLOR_HIGHLIGHT)  'Selected Back Color
      ColorListNum 6, &HFF00FF  'Kernel Mode
      ColorListNum 7, &H80C0FF  'File System
      ColorListNum 8, &H808080  'Adapter
      ColorListNum 9, &H8080&   'Driver
      ColorListNum 10, &HFF8080 'Automatic
      ColorListNum 11, &H800080 'Boot
      ColorListNum 12, &HC0C000 'System
      ColorListNum 13, &H40C0&  'Manual
      ColorListNum 14, &HC0C0C0 'Disabled
      ColorListNum 15, &H80&    'Win 32 Own Process
      ColorListNum 16, &H4080&  'Win 32 Shared Process
      ColorListNum 17, &H4040&  'Own Interactive Process
      ColorListNum 18, &H8000&  'Shared Interactive Process
      ColorListNum 19, GetSysColor(COLOR_WINDOWTEXT)    'Text Color
      ColorListNum 20, GetSysColor(COLOR_WINDOW)        'Back Color
      ColorListNum 21, GetSysColor(COLOR_HIGHLIGHTTEXT) 'Selected Text Color
      ColorListNum 22, 1        'Use Colored Text

      For a = 0 To 5
         Form1.List1(a).FontSize = 8
      Next

    Else
      'if colors already exist load them into the color list structure

      For a = 0 To 22
         Data = CLng(GetSetting("SerVice Commander\" & App.Path, "Options", "Color " & Format(a, _
            "00"), 0))
         ColorListNum a, Data
      Next

      For a = 0 To 5
         Form1.List1(a).FontSize = GetSetting("SerVice Commander\" & App.Path, "Options", _
            "ListFontSize", 8)
      Next

   End If
Exit Sub
GetColorErr:
MsgBox "Service Commander has encountered a problem setting " & vbCrLf & "default colors or is not able to access your registry.", 16, "Service Commander"
End Sub

Private Function ListBackClr(Data As String) As Long

   If ColorList.ColoredText = 0 Then

      ListBackClr = ListColorValue(Data)
      If ListBackClr = -1 Then ListBackClr = ColorList.BackColor

    Else
      ListBackClr = ColorList.BackColor
   End If

End Function

Private Function ListColorValue(Data As String) As Long

   Select Case Data
    Case "Running": ListColorValue = ColorList.Running
    Case "Stopped": ListColorValue = ColorList.Stopped
    Case "Paused": ListColorValue = ColorList.Paused
    Case "Starting": ListColorValue = ColorList.Starting
    Case "Stopping": ListColorValue = ColorList.Stopped
    Case "Kernel Mode Driver": ListColorValue = ColorList.Kernel
    Case "File System Driver": ListColorValue = ColorList.FileSys
    Case "Adapter": ListColorValue = ColorList.Adapter
    Case "Driver Service": ListColorValue = ColorList.Driver
    Case "Win32 Own Process": ListColorValue = ColorList.OwnProcess
    Case "Win32 Shared Process": ListColorValue = ColorList.SharedProcess
    Case "Own Interactive Process": ListColorValue = ColorList.OwnInteractive
    Case "Shared Interactive": ListColorValue = ColorList.SharedInteractive
    Case "Disabled": ListColorValue = ColorList.Disabled
    Case "Automatic": ListColorValue = ColorList.Automatic
    Case "Boot": ListColorValue = ColorList.Boot
    Case "Manual": ListColorValue = ColorList.Manual
    Case "System": ListColorValue = ColorList.System
    Case Else: ListColorValue = -1
   End Select

End Function

Private Function ListTextClr(Data As String) As Long

   If ColorList.ColoredText = 1 Then
      ListTextClr = ListColorValue(Data)
      If ListTextClr = -1 Then ListTextClr = ColorList.TextColor
    Else
      ListTextClr = ColorList.TextColor
   End If

End Function

Public Function SubClassedList(ByVal Hwnd As Long, _
                               ByVal Msg As Long, _
                               ByVal wParam As Long, _
                               ByVal lParam As Long) As Long

  Dim tItem As DRAWITEMSTRUCT

  Dim sBuff As String * 255
  Dim sItem As String

  Dim lBack As Long

   'Debug.Print "&h " & Hex(Msg)

   If Msg = WM_DRAWITEM Then
      'Redraw the listbox
      'This function only passes the Address of the DrawItem Structure, so we need to
      'use the CopyMemory API to Get a Copy into the Variable we setup:
      Call CopyMemory(tItem, ByVal lParam, Len(tItem))
      'Make sure we're dealing with a Listbox
      'titem.

      If tItem.CtlType = ODT_LISTBOX Then
         'Debug.Print tItem.hwndItem

         'Get the Item Text
         Call SendMessage(tItem.hwndItem, LB_GETTEXT, tItem.itemID, ByVal sBuff)
         sItem = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
         'debug.Print sItem

         If (tItem.itemState And ODS_SELECTED) Or (tItem.itemState And ODS_FOCUS) Then

            'Item has Focus, Highlight it with color stored in color list
            lBack = CreateSolidBrush(ColorList.Selected)
            Call FillRect(tItem.hdc, tItem.rcItem, lBack)
            Call SetBkColor(tItem.hdc, ColorList.Selected)
            Call SetTextColor(tItem.hdc, ColorList.SelectedText)
            TextOut tItem.hdc, tItem.rcItem.Left, tItem.rcItem.Top, ByVal sItem, Len(sItem)
            DrawFocusRect tItem.hdc, tItem.rcItem
          Else
            'Item Doesn't Have Focus, Draw it's Colored Background
            'Create a Brush using the Color we stored in the color list
            lBack = CreateSolidBrush(ListBackClr(sItem))
            'Paint the Item Area
            Call FillRect(tItem.hdc, tItem.rcItem, lBack)
            'Set the Text Colors
            Call SetBkColor(tItem.hdc, ListBackClr(sItem))
            Call SetTextColor(tItem.hdc, ListTextClr(sItem))
            'Display the Item Text
            TextOut tItem.hdc, tItem.rcItem.Left, tItem.rcItem.Top, ByVal sItem, Len(sItem)

         End If

         Call DeleteObject(lBack)
         'Don't pass the message on
         SubClassedList = 0
         Exit Function

      End If

   End If

   SubClassedList = CallWindowProc(lPrevWndProc, Hwnd, Msg, wParam, lParam)

End Function

