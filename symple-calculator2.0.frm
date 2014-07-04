VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "vb简易计算器2.0-by20106190唐路平"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   12945
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option3 
      Caption         =   "Gra"
      Height          =   375
      Left            =   3000
      TabIndex        =   41
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Rad"
      Height          =   375
      Left            =   1800
      TabIndex        =   40
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Deg"
      Height          =   375
      Left            =   600
      TabIndex        =   39
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command38 
      Caption         =   "1/x"
      Height          =   615
      Left            =   2760
      TabIndex        =   38
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command37 
      Caption         =   "n!"
      Height          =   615
      Left            =   2760
      TabIndex        =   37
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Log"
      Height          =   615
      Left            =   2760
      TabIndex        =   36
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Ln"
      Height          =   615
      Left            =   2760
      TabIndex        =   35
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command34 
      Caption         =   "x^2"
      Height          =   615
      Left            =   1680
      TabIndex        =   34
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command33 
      Caption         =   "x^3"
      Height          =   615
      Left            =   1680
      TabIndex        =   33
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command32 
      Caption         =   "x^y"
      Height          =   615
      Left            =   1680
      TabIndex        =   32
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Exp"
      Height          =   615
      Left            =   1680
      TabIndex        =   31
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Cot"
      Height          =   615
      Left            =   600
      TabIndex        =   30
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Tan"
      Height          =   615
      Left            =   600
      TabIndex        =   29
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Cos"
      Height          =   615
      Left            =   600
      TabIndex        =   28
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Sin"
      Height          =   615
      Left            =   600
      TabIndex        =   27
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Int"
      Height          =   615
      Left            =   9960
      TabIndex        =   26
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Not"
      Height          =   615
      Left            =   9960
      TabIndex        =   25
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Xor"
      Height          =   615
      Left            =   9960
      TabIndex        =   24
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command23 
      Caption         =   "And"
      Height          =   615
      Left            =   9960
      TabIndex        =   23
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Lsh"
      Height          =   615
      Left            =   9000
      TabIndex        =   22
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Or"
      Height          =   615
      Left            =   9000
      TabIndex        =   21
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Mod"
      Height          =   615
      Left            =   9000
      TabIndex        =   20
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command19 
      Caption         =   "."
      Height          =   615
      Left            =   6600
      TabIndex        =   19
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command18 
      Caption         =   "+/-"
      Height          =   615
      Left            =   5520
      TabIndex        =   18
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      Caption         =   "backspace"
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command12 
      Caption         =   "clear"
      Height          =   375
      Left            =   8760
      TabIndex        =   16
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command11 
      Caption         =   "="
      Height          =   615
      Left            =   9000
      TabIndex        =   15
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      Height          =   615
      Left            =   4440
      TabIndex        =   14
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   615
      Left            =   6600
      TabIndex        =   13
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   615
      Left            =   5520
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   615
      Left            =   4440
      TabIndex        =   11
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   615
      Left            =   6600
      TabIndex        =   10
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   615
      Left            =   5520
      TabIndex        =   9
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   615
      Left            =   6600
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      Caption         =   "/"
      Height          =   615
      Left            =   7920
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "*"
      Height          =   615
      Left            =   7920
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Caption         =   "-"
      Height          =   615
      Left            =   7920
      TabIndex        =   2
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "+"
      Height          =   615
      Left            =   7920
      TabIndex        =   1
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0 '默认数组下标从0开始
Option Explicit
Const PI = 3.14159265358979
Const e = 2.71828182845905

Dim Operation As String '操作符、运算字符串
Dim data1, data2, result As Double '操作数1，操作数2,运算结果

Private Sub Command1_Click() '下面是0~9十个按钮的点击事件
   Text1.Text = Text1.Text + "1"
End Sub
Private Sub Command2_Click()
   Text1.Text = Text1.Text + "2"
End Sub
Private Sub Command3_Click()
   Text1.Text = Text1.Text + "3"
End Sub
Private Sub Command4_Click()
   Text1.Text = Text1.Text + "4"
End Sub
Private Sub Command5_Click()
   Text1.Text = Text1.Text + "5"
End Sub
Private Sub Command6_Click()
   Text1.Text = Text1.Text + "6"
End Sub
Private Sub Command7_Click()
   Text1.Text = Text1.Text + "7"
End Sub
Private Sub Command8_Click()
   Text1.Text = Text1.Text + "8"
End Sub
Private Sub Command9_Click()
   Text1.Text = Text1.Text + "9"
End Sub
Private Sub Command10_Click()
  Text1.Text = Text1.Text + "0"
End Sub
Private Sub Command11_Click() '等号
  
  data2 = Val(Text1.Text)
  
  Select Case Operation
  Case "+": result = data1 + data2
  Case "-": result = data1 - data2
  Case "*": result = data1 * data2
  
  Case "/":
   If (data2 <> 0) Then
     result = data1 / data2
   Else
     Text1.Text = "除数不能为零"
     MsgBox "Error:cannot divided by 0" '除数不能为零
     Exit Sub
   End If
  
  Case "Mod": result = data1 Mod data2
  Case "Or":  result = data1 Or data2
  Case "Shl": result = data1 * 2 ^ data2
  Case "And": result = data1 And data2
  Case "Xor": result = data1 Xor data2
  Case "Not": result = Not data1
  Case "Int": result = Int(data1)
  
  Case "Sin":
    If (Option1.Value = True) Then
      result = Sin(data1 * PI / 180)
    ElseIf (Option2.Value = True) Then
      result = Sin(data1)
    ElseIf (Option2.Value = True) Then
      result = Sin(data1 * 10 / 9)
    End If
    
  Case "Cos":
    If (Option1.Value = True) Then
      result = Cos(data1 * PI / 180)
    ElseIf (Option2.Value = True) Then
      result = Cos(data1)
    ElseIf (Option3.Value = True) Then
      result = Cos(data1 * 10 / 9)
    End If
    
  Case "Tan":
    If (Option1.Value = True) Then
      result = Tan(data1 * PI / 180)
    ElseIf (Option2.Value = True) Then
      result = Tan(data1)
    ElseIf (Option3.Value = True) Then
      result = Tan(data1 * 10 / 9)
      
    End If
  Case "Cot":
    If (Option1.Value = True) Then
      data1 = data1 * PI / 180
    ElseIf (Option2.Value = True) Then
      data1 = data1
    ElseIf (Option3.Value = True) Then
      data1 = data1 * 10 / 9
    End If
    
    If (Tan(data1) <> 0) Then
      result = 1 / Tan(data1)
    Else
      Text1.Text = "除数不能为零"
      MsgBox "Error:cannot divided by 0" '除数不能为零
     Exit Sub
    End If
  
  Case "Exp":    result = data1 * 10 ^ data2
  Case "Power":  result = data1 ^ data2
  Case "Cube":   result = data1 ^ 3
  Case "Square": result = data1 * data1
  Case "Ln":     result = Log(data1)
  Case "Log":    result = Log(data1) / Log(10)
  
  Case "Fac":
    result = 1
    For data2 = Int(data1) To 1 Step -1
      result = result * data2
    Next
  
  Case "Rec":
    If (data1 <> 0) Then
      result = 1 / data1
    Else
      Text1.Text = "除数不能为零"
      MsgBox "Error:cannot divided by 0" '除数不能为零
     Exit Sub
    End If
  
  Case Else
    Text1.Text = "有错误发生"
    MsgBox "有错误发生": End
  
  End Select
  
  Text1.Text = Str(result)
End Sub
Private Sub Command12_Click() '清空
   Text1.Text = "" 'clear
End Sub
Private Sub Command13_Click() '加法
   data1 = Val(Text1.Text): Operation = "+": Text1.Text = ""
End Sub

Private Sub Command14_Click() '减法
   data1 = Val(Text1.Text): Operation = "-": Text1.Text = ""
End Sub

Private Sub Command15_Click() '乘法
   data1 = Val(Text1.Text): Operation = "*": Text1.Text = ""
End Sub
Private Sub Command16_Click() '除法
   data1 = Val(Text1.Text): Operation = "/": Text1.Text = ""
End Sub
Private Sub Command17_Click() '退格
   If (Len(Text1.Text) <> 0) Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
End Sub
Private Sub Command18_Click() '正负号
   If Mid(Text1.Text, 1, 1) = "-" Then
     Text1.Text = Right(Text1.Text, Len(Text1.Text) - 1)
   Else
     Text1.Text = "-" & Text1.Text
   End If
End Sub
Private Sub Command19_Click() '小数点
   If (InStr(1, Text1.Text, ".", 1) = 0) Then
     Text1.Text = Text1.Text + "."
   End If
End Sub
Private Sub Command20_Click() '模运算
   data1 = Val(Text1.Text): Operation = "Mod": Text1.Text = ""
End Sub
Private Sub Command21_Click() '或运算
   data1 = Val(Text1.Text): Operation = "Or": Text1.Text = ""
End Sub
Private Sub Command22_Click() '左移运算
   data1 = Val(Text1.Text): Operation = "Shl": Text1.Text = ""
End Sub
Private Sub Command23_Click() '与运算
   data1 = Val(Text1.Text): Operation = "And": Text1.Text = ""
End Sub
Private Sub Command24_Click() '异或运算
   data1 = Val(Text1.Text): Operation = "Xor": Text1.Text = ""
End Sub
Private Sub Command25_Click() '非运算
   data1 = Val(Text1.Text): Operation = "Not": Call Command11_Click
End Sub
Private Sub Command26_Click() '取整运算
   data1 = Val(Text1.Text): Operation = "Int": Call Command11_Click
End Sub
Private Sub Command27_Click() '正弦
   data1 = Val(Text1.Text): Operation = "Sin": Call Command11_Click
End Sub
Private Sub Command28_Click() '余弦
   data1 = Val(Text1.Text): Operation = "Cos": Call Command11_Click
End Sub
Private Sub Command29_Click() '正切
   data1 = Val(Text1.Text): Operation = "Tan": Call Command11_Click
End Sub
Private Sub Command30_Click() '余切
   data1 = Val(Text1.Text): Operation = "Cot": Call Command11_Click
End Sub
Private Sub Command31_Click() '幂指数函数运算
   data1 = Val(Text1.Text): Operation = "Exp": Text1.Text = ""
End Sub
Private Sub Command32_Click() '乘方
   data1 = Val(Text1.Text): Operation = "Power": Text1.Text = ""
End Sub
Private Sub Command33_Click() '立方
   data1 = Val(Text1.Text): Operation = "Cube": Call Command11_Click
End Sub
Private Sub Command34_Click() '平方
   data1 = Val(Text1.Text): Operation = "Square": Call Command11_Click
End Sub
Private Sub Command35_Click() '自然对数
   data1 = Val(Text1.Text): Operation = "Ln": Call Command11_Click
End Sub
Private Sub Command36_Click() '常用对数
   data1 = Val(Text1.Text): Operation = "Log": Call Command11_Click
End Sub
Private Sub Command37_Click() '阶乘
   data1 = Val(Text1.Text): Operation = "Fac": Call Command11_Click 'Factorial
End Sub
Private Sub Command38_Click() '取倒数
   data1 = Val(Text1.Text): Operation = "Rec": Call Command11_Click 'Reciprocal
End Sub
Private Sub Form_Load()
  Text1.Text = "" '文本区初始为空
  Option1.Value = True '默认为角度
  Form1.Picture = LoadPicture(".\forest.jpg") '背景图片
End Sub
