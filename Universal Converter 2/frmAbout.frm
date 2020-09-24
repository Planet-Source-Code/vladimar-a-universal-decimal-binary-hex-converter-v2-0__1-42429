VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   5055
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   6960
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3490
   ScaleMode       =   0  'User
   ScaleWidth      =   6536
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInfo 
      Height          =   4335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   -87
      Width           =   6975
      Begin VB.Label lblSplash 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Splash Screen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   160
         Width           =   1335
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6360
         TabIndex        =   5
         Top             =   160
         Width           =   495
      End
      Begin VB.Label lblSysInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "System Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5040
         TabIndex        =   4
         Top             =   160
         Width           =   1095
      End
      Begin VB.Label lblInstruct 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Instructions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   160
         Width           =   975
      End
      Begin VB.Label lblDevNotes 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Developer Notes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   160
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   6975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
' the above 2 lines are nessesary to allow the form to be dragged from any point
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim I As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub lblClose_Click()
  Unload Me
End Sub


Private Sub lblDevNotes_Click()
Dim info As String

info = _
"Developer Notes:" & vbCrLf & vbCrLf _
& "    I have tried to comment my code as fully as possible, I hope you find it useful." & vbCrLf _
& "    Please give credit to those who's work goes into your own projects, including me ; )" & vbCrLf _
& "    If you have any comments, feedback, suggestions etc. please Email me:" & vbCrLf _
& "         <vladimar@eircom.net>" & vbCrLf _
& "    If you are a regular visitor to Planet-Source-Code.com, and found this code/program" & vbCrLf _
& "         useful, please vote for me." & vbCrLf & vbCrLf _
& "    This is my first fully-working program, a previous 'beta' version already resides in" & vbCrLf _
& "         Planet-Source-Code.com's database." & vbCrLf _
& "    Planet-Source-Code.com has been a huge source of information / inspiration so far, I" & vbCrLf _
& "         thank those who have contributed to this program, whether they know it or not." & vbCrLf & vbCrLf _
& "    I give permission to use the VB source code only, the splash-screen image is original" & vbCrLf _
& "         work and may not be reproduced without my say-so."


txtInfo.Text = info
End Sub

Private Sub lblInstruct_Click()
Dim info, info2 As String

info = "Instructions:" & vbCrLf & vbCrLf _
    & "1 - Entering numbers:" & vbCrLf & vbCrLf _
    & "     To enter a number, just click on the name of the type of number you wish to enter" & vbCrLf _
    & "         [decimal, binary, hex]. The text-box for this type will then be selected and ready." & vbCrLf _
    & "     Then select a type of binary from the drop-down list, otherwise it will default to" & vbCrLf _
    & "         'Variable Length'." & vbCrLf _
    & "     Hit Enter, or click on 'Convert'." & vbCrLf & vbCrLf _
    & "2 - Number Types:" & vbCrLf & vbCrLf _
    & "     Decimal numbers can be positive, negative, decimal, whatever." & vbCrLf _
    & "     Binary numbers come in 5 different flavours:" & vbCrLf _
    & "             8bit / 2's compliment: 8 characters long, can show both positive and negative" & vbCrLf _
    & "                 numbers using the 2's Compliment method." & vbCrLf _
    & "             16bit: 16 characters long, treats all numbers as being positive." & vbCrLf _
    & "             Fixed-Point: 16 characters long, uses a sign-bit at the start [1 = -, 0 = +], gives" & vbCrLf _
    & "                 7bits above the decimal point, 8 below." & vbCrLf _
    & "             Floating-Point: 32 characters long, uses the same sign-bit and an 8bit exponent," & vbCrLf _
    & "                 then a varying proportion of the space is given to the numbers above and below" & vbCrLf _
    & "                 the decimal point [numbers above will take up as much as is needed]." & vbCrLf _
    & "             Variable Length: gives the simple binary at upto 31 bits." & vbCrLf _
    & "     Hex numbers are numbers in base 16, can show + or -, uses letters A, B, C, D, E, F."
info2 = "3 - Limitations" & vbCrLf & vbCrLf _
    & "     The Highest Decimal number you can enter is D(2147483647), any number greater" & vbCrLf _
    & "         than this will trigger an error message." & vbCrLf _
    & "     The highest Fixed-Point number that can be diplayed is D(127), any number" & vbCrLf _
    & "         over this will not convert to Fixed-Point accurately." & vbCrLf _
    & "     The Highest Floating-Point number that can be displayed accurately is D(16777215)." & vbCrLf _
    & "     The longest binary number that can be converted is 31bits, any number longer" & vbCrLf _
    & "         will trigger an error message." & vbCrLf _
    & "     The largest hex number that can be handled is H(7FFFFFFF), D(2147483647)."

txtInfo.Text = info & vbCrLf & vbCrLf & info2
End Sub

Private Sub lblSplash_Click()
frmSplash.Show
End Sub

Private Sub lblSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
