VERSION 5.00
Begin VB.Form Alpha 
   BackColor       =   &H00404000&
   Caption         =   "Alpha Calculator"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3645
   FillColor       =   &H00404040&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Alpha Beta BRK"
      Size            =   36
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton devide 
      Caption         =   "\"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2280
      TabIndex        =   18
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton multiply 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   17
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton minus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   16
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton plus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   15
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton equals 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton dot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   13
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton clear 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton digits 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   840
      TabIndex        =   10
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton digits 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1440
      TabIndex        =   9
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton digits 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   840
      TabIndex        =   8
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton digits 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton digits 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1440
      TabIndex        =   6
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton digits 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   840
      TabIndex        =   5
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton digits 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton digits 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton digits 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton digits 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      MaskColor       =   &H008080FF&
      TabIndex        =   1
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000A&
      Height          =   2895
      Left            =   2160
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000A&
      Height          =   2895
      Left            =   120
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "CASIO"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Alpha Beta BRK"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000A&
      Height          =   735
      Left            =   120
      Top             =   240
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      X1              =   120
      X2              =   3480
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lbldisplay 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Alpha Beta BRK"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   735
      Left            =   120
      MouseIcon       =   "Form1.frx":2A281
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Alpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oper1 As Double, Oper2 As Double, Result As Double
Dim Operator As String
Dim ClearDisp As Boolean

Private Sub clear_Click()
lbldisplay.Caption = ""
End Sub

Private Sub devide_Click()
oper1 = Val(lbldisplay.Caption)
Operator = "/"
lbldisplay.Caption = ""
equals.Enabled = True
End Sub

Private Sub dot_Click()
If ClearDisp = True Then
lbldisplay.Caption = ""
ClearDisp = False
End If
If InStr(lbldisplay.Caption, ".") Then
Exit Sub
Else
lbldisplay.Caption = lbldisplay.Caption + "."
End If
End Sub

Private Sub equals_Click()
On Error GoTo ErrorMsg
Oper2 = Val(lbldisplay.Caption)
If Operator = "+" Then Result = sum(oper1, Oper2)
If Operator = "-" Then Result = Diff(oper1, Oper2)
If Operator = "*" Then Result = Product(oper1, Oper2)
If Operator = "/" Then Result = Division(oper1, Oper2)
lbldisplay.Caption = Result
ClearDisp = True
equals.Enabled = False
Exit Sub
ErrorMsg:
MsgBox "The Operation resulted in the Error:" & Err.Description
lbl.display.Caption = "ERROR"
ClearDisp = True
End Sub

Private Sub Form_Load()
equals.Enabled = False
End Sub

Private Sub digits_Click(Index As Integer)
If ClearDisp = True Then
lbldisplay.Caption = ""
ClearDisp = False
End If
lbldisplay.Caption = lbldisplay.Caption + digits(Index).Caption
End Sub


Private Sub minus_Click()
oper1 = Val(lbldisplay.Caption)
Operator = "-"
lbldisplay.Caption = ""
equals.Enabled = True
End Sub

Private Sub multiply_Click()
oper1 = Val(lbldisplay.Caption)
Operator = "*"
lbldisplay.Caption = ""
equals.Enabled = True
End Sub

Private Sub plus_Click()
oper1 = Val(lbldisplay.Caption)
Operator = "+"
lbldisplay.Caption = ""
equals.Enabled = True
End Sub

Private Function sum(x As Double, y As Double)
sum = x + y
End Function

Private Function Diff(x As Double, y As Double)
Diff = x - y
End Function

Private Function Product(x As Double, y As Double)
Product = x * y
End Function

Private Function Division(x As Double, y As Double)
Division = x / y
End Function
