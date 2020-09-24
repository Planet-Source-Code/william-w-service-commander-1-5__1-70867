Attribute VB_Name = "Choose_Color"
Option Explicit
'Color Choosing Dialog
Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100

Private Type CHOOSECOLOR
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   rgbResult As Long
   lpCustColors As String
   Flags As Long
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Private Declare Function ChooseColorAPI Lib "comdlg32.dll" _
      Alias "ChooseColorA" ( _
      pChoosecolor As CHOOSECOLOR) As Long


Public Function ShowColor(Color As Long, Flags As Long, Optional Hwnd As Long = 0) As Long
On Error GoTo ShowColorErr
  Dim cc As CHOOSECOLOR
  Dim lReturn As Long
  Dim CustomColors() As Byte
  Dim a As Long

   ReDim CustomColors(0 To 63) As Byte

   For a = 0 To 63 'load custom colors from registry
      CustomColors(a) = CByte(GetSetting("SerVice Commander\" & App.Path, "Options", "Custom Color" & _
         " " & Format(a, "00"), 0))
   Next

   cc.Flags = Flags
   cc.rgbResult = Color
   cc.lStructSize = Len(cc)
   cc.hwndOwner = Hwnd
   cc.hInstance = 0
   cc.lpCustColors = StrConv(CustomColors, vbUnicode)

   lReturn = ChooseColorAPI(cc)

   If lReturn <> 0 Then
      ShowColor = cc.rgbResult
      CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)

    Else
      ShowColor = -1
   End If

   For a = 0 To 63 'save custom colors to registry
      SaveSetting "SerVice Commander\" & App.Path, "Options", "Custom Color " & Format(a, "00"), _
         CustomColors(a)
   Next
Exit Function
ShowColorErr:
MsgBox "Service Commander has encountered a problem calling the color dialog" & vbCrLf & " or is not able to save or retrieve your custom colors from your registry", 0, "Service Commander"
Resume Next
End Function

