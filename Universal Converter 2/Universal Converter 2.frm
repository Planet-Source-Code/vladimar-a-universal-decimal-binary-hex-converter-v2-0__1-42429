VERSION 5.00
Begin VB.Form frmUniversalConverter2 
   BorderStyle     =   0  'None
   Caption         =   "Universal Decimal / Binary / Hex Converter - 2.0"
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   Icon            =   "Universal Converter 2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   0
      TabIndex        =   13
      Top             =   -90
      Width           =   7575
      Begin VB.Label Label3 
         Caption         =   "_"
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
         Left            =   7080
         TabIndex        =   17
         Top             =   165
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Universal Decimal / Binary / Hex Converter v2.0"
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
         Left            =   1673
         TabIndex        =   16
         Top             =   165
         Width           =   4215
      End
      Begin VB.Label lblClose 
         Caption         =   "X"
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
         Left            =   7320
         TabIndex        =   15
         Top             =   165
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "Help"
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
         TabIndex        =   14
         Top             =   165
         Width           =   495
      End
   End
   Begin VB.TextBox txtHex 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
   End
   Begin VB.OptionButton optDecimal 
      Caption         =   "Decimal"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton optBinary 
      Caption         =   "Binary"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton optHex 
      Caption         =   "Hex"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.ComboBox cmbBinType 
      Height          =   315
      ItemData        =   "Universal Converter 2.frx":030A
      Left            =   5640
      List            =   "Universal Converter 2.frx":031D
      TabIndex        =   4
      Text            =   "Type of Binary"
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton btnConvert 
      Caption         =   "Convert"
      Default         =   -1  'True
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtDecimal 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox txtBinary 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      MaxLength       =   32
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
   Begin VB.Frame fmFrame 
      Height          =   1335
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   7575
      Begin VB.Label lblDecChars 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   5355
         TabIndex        =   12
         Top             =   265
         Width           =   210
      End
      Begin VB.Label lblBinChars 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   5355
         TabIndex        =   11
         Top             =   625
         Width           =   210
      End
      Begin VB.Label lblHexChars 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   5355
         TabIndex        =   10
         Top             =   985
         Width           =   210
      End
   End
End
Attribute VB_Name = "frmUniversalConverter2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim selection As String
Dim abHidden As Boolean
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


Private Sub btnConvert_Click()
Dim decimalVal, binaryVal, hexVal, wk As String
Dim T, L As Integer
T = 0       'flag specifying target length of relevant text when used in lengthFix function
L = Empty   'used for the optional flag specifying whether if the number is to be padded from the left

On Error GoTo error ' skips straight to the specified title when there is an error
If cmbBinType.Text = "Type of Binary" Then cmbBinType.Text = "Variable Length"

Select Case selection
    Case "decimal"
        decimalVal = Val(txtDecimal.Text)
        binaryVal = DoConvert(decimalVal, selection, cmbBinType.Text)   'checks status of comboBox
        wk = SignSet(decimalVal, "-", "", 0)
        hexVal = SignSet(decimalVal, "-", "", 1) & Hex(wk)
    Case "binary"
        Select Case cmbBinType.Text
            Case "Fixed-Point": T = 16
            Case "Floating-Point": T = 32
            Case "8bit / 2's comp": T = 8: L = 1
            Case "16bit": T = 16: L = 1
            Case Else: T = Len(txtBinary.Text)
        End Select
        txtBinary.Text = LengthFix(txtBinary.Text, T, L)    'when a number is entered in txtbinary, it's length is adjusted to make sure the correct operation is performaed
        binaryVal = txtBinary.Text
        decimalVal = DoConvert(binaryVal, cmbBinType.Text, "decimal")
        wk = SignSet(decimalVal, "-", "", 0)
        hexVal = SignSet(decimalVal, "-", "", 1) & Hex(wk)
    Case "hex"
        hexVal = txtHex.Text
        wk = SignSet(hexVal, "-", "", 4)
        decimalVal = wk & DoConvert(hexVal, selection, "decimal")
        binaryVal = DoConvert(decimalVal, "decimal", cmbBinType.Text)
End Select

txtBinary.Text = binaryVal
txtDecimal.Text = decimalVal
txtHex.Text = hexVal

Exit Sub

error:
If Err.Number = 6 Then
MsgBox "There has been an overflow, the numbers you are using are too large", vbExclamation, "Overflow: Operation Aborted"
Exit Sub
End If


End Sub

Private Function DoConvert(ByVal inputVal As String, inputType As String, outputType As String) As String

Select Case inputType
            Case "decimal"
                Select Case outputType
                        Case "8bit / 2's comp"
                            DoConvert = BinOfDec(inputVal, 8)
                        Case "16bit"
                            DoConvert = BinOfDec(inputVal, 16)
                        Case "Fixed-Point"
                            DoConvert = FixedOfDec(inputVal)
                        Case "Floating-Point"
                            DoConvert = FloatOfDec(inputVal)
                        Case "Variable Length"
                            DoConvert = BinOfDec(inputVal)
                End Select
            Case "8bit / 2's comp"
                DoConvert = DecOfBin(inputVal)
            Case "16bit"
                DoConvert = DecOfBin(inputVal)
            Case "Variable Length"
                DoConvert = DecOfBin(inputVal)
            Case "Fixed-Point"
                'txtBinary.Text = LengthFix(txtBinary.Text, 16, 1)
                inputVal = LengthFix(inputVal, 16, 1)
                DoConvert = DecOfFixed(inputVal)
            Case "Floating-Point"
                txtBinary.Text = LengthFix(txtBinary.Text, 32)
                inputVal = LengthFix(inputVal, 32)
                DoConvert = DecOfFloat(inputVal)
            Case "hex"
                DoConvert = DecOfHex(inputVal)
End Select

End Function

Private Function LengthFix(ByVal subject As String, ByVal target As Integer, Optional ByVal leftPad As Integer) As String
'if there is a third parameter given to the function, regardless of it's value, the number will be padded from the left.
Dim L As Integer
L = Len(subject)


    If L < target Then
        If leftPad = Empty Then
            subject = subject & String((target - L), "0")
        Else
            subject = String((target - L), "0") & subject
        End If
    ElseIf L > target Then
        subject = Left(subject, target)
End If

LengthFix = subject

End Function


Private Function SignSet(ByVal Number As String, signMinus As String, signPlus As String, toReturn As Integer) As String
Dim L, K As Integer
Dim sign As String

' toReturn: 0 to return the number in positive form
' toReturn: 1 to return the specified sign symbol
' toReturn: 2 to swap signs if either is present
' toReturn: 3 to inverse sign of fixed / floating point number
' toReturn: 4 to check for the specified signs and return the sign it finds
    
    L = Len(Number) - 1

Select Case toReturn
    Case 0
        If Number <= 0 Then SignSet = (Number * -1) Else SignSet = Number
    Case 1
        If Number < 0 Then SignSet = signMinus Else SignSet = signPlus
    Case 2
        If Left(Number, 1) = signMinus Then
            SignSet = signPlus & Right(Number, L)
        ElseIf Left(Number, 1) = signPlus Then
            SignSet = signMinus & Right(Number, L)
        End If
    Case 3
        If cmbBinType.Text = "Fixed-Point" Or cmbBinType.Text = "Floating-Point" Then
            If Left(Number, 1) = signMinus Then
                SignSet = signPlus & Right(Number, L)
            ElseIf Left(Number, 1) = signPlus Then
                SignSet = signMinus & Right(Number, L)
            End If
        End If
    Case 4
        If Left(Number, 1) = signMinus Then
            SignSet = signMinus
        ElseIf Left(Number, 1) = signPlus Then
            SignSet = signPlus
        End If
        
End Select

End Function


Private Function decimalPoint(Number As String, toReturn As Integer) As String
'splits a decimal number into the parts before and after the decimal point

Dim top, bottom, point As String
point = InStr(1, Number, ".")

If point >= 1 Then

    bottom = Right(Number, (Len(Number) - point))

    top = Int(Number)
    
Else

    top = Int(Number)
    
    bottom = 0
    
End If

Select Case toReturn
    Case 0  'Returns top half (xx.)
        decimalPoint = top
    Case 1  'Returns bottom half (.xx) as integer e.g. 256 for 0.256
        decimalPoint = bottom
    Case 2  'Returns bottom half (.xx) as point number e.g. 0.256
        Do While bottom > 1
            bottom = bottom / 10
        Loop
        decimalPoint = bottom
End Select
End Function


Private Sub Form_Load()
selection = "decimal"
txtBinary.Text = "Please Vote For Me"

txtBinary.BackColor = &H8000000F
txtBinary.Locked = True
txtHex.BackColor = &H8000000F
txtHex.Locked = True

End Sub

Private Sub Form_Resize()
    If frmUniversalConverter2.WindowState = 0 And abHidden = True Then
        frmAbout.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'clears all memory used by the program on exit
    Dim I As Integer
    
    'Remove Graphics for All Forms


    For I = Forms.Count - 1 To 0 Step -1
        Unload Forms(I)
    Next
    'Remove Binary Code for All Forms
    Set frmLast = Nothing
    Set frmPrint = Nothing
    Set frmDetail = Nothing
    Set frmMain = Nothing
End Sub

Private Sub fmFrame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Label2_Click()
frmAbout.Show
frmUniversalConverter2.SetFocus
End Sub

Private Sub Label3_Click()
Me.WindowState = 1
    If frmAbout.Visible = True Then
        abHidden = True
        frmAbout.Visible = False
    Else
        abHidden = False
    End If
End Sub

Private Sub lblClose_Click()
End
End Sub

Private Sub txtBinary_Change()
lblBinChars.Caption = Len(txtBinary.Text)

End Sub

Private Sub txtBinary_KeyPress(KeyAscii As Integer)
    Const Number$ = "01" ' only allow these characters

    If KeyAscii <> 8 Then
    If InStr(Number$, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

End Sub

Private Sub txtDecimal_Change()
' changes the number on the right of the text-boxes to the number of digits used
lblDecChars.Caption = Len(txtDecimal.Text)
End Sub

Private Sub txtDecimal_KeyPress(KeyAscii As Integer)

    Const Number$ = "0123456789.-" ' only allow these characters

    If KeyAscii <> 8 Then
        If InStr(Number$, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    

    
End Sub

Private Sub txtHex_Change()
lblHexChars.Caption = Len(txtHex.Text)
End Sub

Private Sub txtHex_KeyPress(KeyAscii As Integer)
    Const Number$ = "0123456789-ABCDEFabcdef" ' only allow these characters
    Dim char As String
    
    char = UCase(Chr(KeyAscii))         'makes all characters entered uppercase
    KeyAscii = AscW(char)
    
    If KeyAscii <> 8 Then
    If InStr(Number$, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If



End Sub

Private Sub optBinary_Click()
selection = "binary"                    'notifies the button of position
txtBinary.Locked = False
txtDecimal.Locked = True                'locks the non-input text boxes
txtHex.Locked = True
txtBinary.BackColor = &HFFFFFF
txtDecimal.BackColor = &H8000000F
txtHex.BackColor = &H8000000F
txtBinary.SetFocus
txtBinary.SelStart = 0
txtBinary.SelLength = Len(txtBinary)
End Sub
Private Sub optDecimal_Click()
selection = "decimal"
txtDecimal.Locked = False
txtBinary.Locked = True
txtHex.Locked = True
txtBinary.BackColor = &H8000000F
txtDecimal.BackColor = &HFFFFFF
txtHex.BackColor = &H8000000F
txtDecimal.SetFocus
txtDecimal.SelStart = 0
txtDecimal.SelLength = Len(txtDecimal)
End Sub
Private Sub optHex_Click()
selection = "hex"
txtHex.Locked = False
txtDecimal.Locked = True
txtBinary.Locked = True
txtBinary.BackColor = &H8000000F
txtHex.BackColor = &HFFFFFF
txtDecimal.BackColor = &H8000000F
txtHex.SetFocus
txtHex.SelStart = 0
txtHex.SelLength = Len(txtHex)
End Sub
Private Sub btnReset_Click()
selection = "decimal"
txtDecimal.Text = ""
txtBinary.Text = ""
txtHex.Text = ""
optDecimal = True
Call optDecimal_Click
End Sub

Private Function FixedOfDec(ByVal decimalNumber As String) As String
Dim sign, mantissa1, mantissa2 As String

'0  set sign (0 / 1)
sign = SignSet(decimalNumber, "1", "0", 1)
'1  make number positive if not
decimalNumber = SignSet(decimalNumber, "1", "0", 0)
'2  split into xx. and .xx
mantissa1 = decimalPoint(decimalNumber, 0)
mantissa2 = decimalPoint(decimalNumber, 2)
'3  get 7bit binary of xx.
mantissa1 = BinOfDec(mantissa1, 7)
'4  get 8bit binary of .xx
mantissa2 = BinOfPointDec(mantissa2, 8)
'5  construct in form sign/mantissa1/mantissa2
FixedOfDec = sign & mantissa1 & mantissa2
'6 ensure correct length of output
FixedOfDec = LengthFix(FixedOfDec, 16)

End Function

Private Function DecOfFixed(ByVal Fixed As String) As String
Dim sign, mantissa1, mantissa2 As String

'1 set sign
If Left(Fixed, 1) = 1 Then sign = "-" Else sign = ""
'2 split into relevant parts
mantissa1 = Mid(Fixed, 2, 7)
mantissa2 = Mid(Fixed, 9, 8)
'3 get decimal of parts
mantissa1 = DecOfBin(mantissa1)
mantissa2 = DecOfPointBin(mantissa2)
'4 construct
DecOfFixed = sign & (Val(mantissa1) + Val(mantissa2))


End Function

Private Function DecOfFloat(ByVal Float As String) As String
Dim sign, mantissa1, mantissa2, exponent As String

If Left(Float, 1) = 1 Then sign = "-" Else sign = ""

exponent = Mid(Float, 2, 8)

exponent = DecOfBin(exponent)

exponent = (exponent - 127)

mantissa1 = 1 & Mid(Float, 10, exponent)

mantissa2 = Mid(Float, (10 + exponent), 32)

mantissa1 = DecOfBin(mantissa1)
mantissa2 = DecOfPointBin(mantissa2)

DecOfFloat = sign & (Val(mantissa1) + Val(mantissa2))


End Function


Private Function DecOfBin(ByVal Number As String) As String

    Dim K%
    Dim L%
    Dim D&
    Dim B$

    B = CStr(Number)
    L = Len(B)

    For K = 1 To L
        If Mid(B, K, 1) = "1" Then D = D + (2 ^ (L - K))
    Next

    DecOfBin = D

End Function
Private Function DecOfPointBin(ByVal Number As String) As String

    Dim K%
    Dim L%
    Dim D As Double
    Dim B$

    B = CStr(Number)
    L = Len(B)

    For K = 1 To L
        If Mid(B, K, 1) = "1" Then D = D + ((0.5) ^ K)
    Next

    DecOfPointBin = D

End Function

Private Function FloatOfDec(ByVal decimalNumber As String)
Dim sign, mantissa1, mantissa2, exponent As String
    '0  set sign (0 / 1)
sign = SignSet(decimalNumber, "1", "0", 1)
    '1  make number positive if not
decimalNumber = SignSet(decimalNumber, "1", "0", 0)
    '2  split into xx. and .xx
mantissa1 = decimalPoint(decimalNumber, 0)
mantissa2 = decimalPoint(decimalNumber, 2)
    '3  take xx. and get binary
mantissa1 = BinOfDec(mantissa1)
    '4  get exponent with len(binary of xx.)
mantissa1 = Right(mantissa1, (Len(mantissa1) - 1))
    '5  cut leading 1 from binary of xx.
exponent = Len(mantissa1)
    '6  add 127 to exponent
exponent = Val(exponent) + 127
    '7  get binary of exponent
exponent = BinOfDec(exponent)
exponent = LengthFix(exponent, 8, 1)
    '8  get binary of .xx
mantissa2 = BinOfPointDec(mantissa2)
    '9  construct as output, sign/exponent/binxx./bin.xx
FloatOfDec = sign & exponent & mantissa1 & mantissa2
    '10 set length to 32 bits
FloatOfDec = LengthFix(FloatOfDec, 32)

End Function

Private Function BinOfDec(ByVal Number As String, Optional length As Integer) As String

Dim D, B, L, wk, C

D = Number
L = 0


If length = Empty Then
    Do
        If D Mod 2 Then B = "1" & B Else B = "0" & B
        D = D \ 2
    Loop Until D = 0
Else
    Do
        If D Mod 2 Then B = "1" & B Else B = "0" & B
        D = D \ 2
        L = L + 1
    Loop Until L = length
End If

If Number < 0 And length = 8 Then GoTo TwosCompliment

GoTo BinAns

TwosCompliment:     'the binary is inverted and 1 is added, this is how minus numbers are represented in binary
    L = Len(B)
    D = 0
    C = 1
        For D = L To 1 Step -1
            wk = Mid(B, D, 1)
            
            If wk = 1 Then wk = 0 Else wk = 1   'inverse
            
            If wk = 1 And C = 1 Then            'add 1
                wk = 0
                C = 1
            ElseIf wk = 0 And C = 1 Then
                wk = 1
                C = 0
            ElseIf wk = 1 And C = 0 Then
                wk = 1
                C = 0
            ElseIf wk = 0 And C = 0 Then
                wk = 0
                C = 0
            End If
                        
            BinOfDec = BinOfDec & wk
                        
        Next D
  
Exit Function
BinAns:
BinOfDec = B

End Function

Private Function BinOfPointDec(ByVal Number As String, Optional length As Integer) As String
Dim B, wk As String
Dim L, I As Integer
wk = Number
    'to get the binary of a decimal number (0.xxx), the number is multiplied by 2, if the result is greater than 1
    'then a 1 is returned, and the number has reduced by 1, repeat...
If length = Empty Then length = 32

For I = 1 To length
If wk = 0 Then GoTo stoploop    'avoids unnecessary loops when the number is reduced completely
    If wk * 2 >= 1 Then
        L = 1
        wk = wk * 2
        wk = wk - 1
    Else
        L = 0
        wk = wk * 2
    End If
    
B = B & L

Next I

stoploop:

BinOfPointDec = B

End Function

Private Function DecOfHex(ByVal hexNumber As String) As String
    Dim K%
    Dim L%
    Dim V&
    Dim D&
    
    hexNumber = UCase(hexNumber)
    
    L = Len(hexNumber)

    For K = 1 To L
        Select Case Mid(hexNumber, K, 1)
               Case "A": V = 10
               Case "B": V = 11
               Case "C": V = 12
               Case "D": V = 13
               Case "E": V = 14
               Case "F": V = 15
               Case Else
                    V = Val(Mid(hexNumber, K, 1))
        End Select
        D = D + V * 16 ^ (L - K)
    Next
    
DecOfHex = D

End Function
