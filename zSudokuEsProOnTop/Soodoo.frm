VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4020
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   3360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   240
      Index           =   0
      Left            =   2640
      TabIndex        =   101
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ReForm"
      Height          =   255
      Left            =   2640
      TabIndex        =   99
      Top             =   2400
      Width           =   960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "zCheck"
      Height          =   255
      Left            =   2640
      TabIndex        =   98
      Top             =   2160
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Index           =   14
      Left            =   3360
      TabIndex        =   97
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ReStart"
      Height          =   255
      Left            =   2640
      TabIndex        =   96
      Top             =   1920
      Width           =   960
   End
   Begin VB.CommandButton Command6 
      Caption         =   "HelpSol"
      Height          =   255
      Left            =   2640
      TabIndex        =   95
      Top             =   2640
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   2280
      TabIndex        =   94
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   80
      Left            =   2040
      TabIndex        =   93
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   79
      Left            =   1800
      TabIndex        =   92
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   78
      Left            =   1440
      TabIndex        =   91
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   77
      Left            =   1200
      TabIndex        =   90
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   76
      Left            =   960
      TabIndex        =   89
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   75
      Left            =   600
      TabIndex        =   88
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   74
      Left            =   360
      TabIndex        =   87
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   73
      Left            =   120
      TabIndex        =   86
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   72
      Left            =   2280
      TabIndex        =   85
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   71
      Left            =   2040
      TabIndex        =   84
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   70
      Left            =   1800
      TabIndex        =   83
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   69
      Left            =   1440
      TabIndex        =   82
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   68
      Left            =   1200
      TabIndex        =   81
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   67
      Left            =   960
      TabIndex        =   80
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   66
      Left            =   600
      TabIndex        =   79
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   65
      Left            =   360
      TabIndex        =   78
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   64
      Left            =   120
      TabIndex        =   77
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   63
      Left            =   2280
      TabIndex        =   76
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   62
      Left            =   2040
      TabIndex        =   75
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   61
      Left            =   1800
      TabIndex        =   74
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   60
      Left            =   1440
      TabIndex        =   73
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   59
      Left            =   1200
      TabIndex        =   72
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   58
      Left            =   960
      TabIndex        =   71
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   57
      Left            =   600
      TabIndex        =   70
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   56
      Left            =   360
      TabIndex        =   69
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   55
      Left            =   120
      TabIndex        =   68
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   54
      Left            =   2280
      TabIndex        =   67
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   53
      Left            =   2040
      TabIndex        =   66
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   52
      Left            =   1800
      TabIndex        =   65
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   51
      Left            =   1440
      TabIndex        =   64
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   50
      Left            =   1200
      TabIndex        =   63
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   49
      Left            =   960
      TabIndex        =   62
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   48
      Left            =   600
      TabIndex        =   61
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   47
      Left            =   360
      TabIndex        =   60
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   46
      Left            =   120
      TabIndex        =   59
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   45
      Left            =   2280
      TabIndex        =   58
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   44
      Left            =   2040
      TabIndex        =   57
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   43
      Left            =   1800
      TabIndex        =   56
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   42
      Left            =   1440
      TabIndex        =   55
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   41
      Left            =   1200
      TabIndex        =   54
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   40
      Left            =   960
      TabIndex        =   53
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   39
      Left            =   600
      TabIndex        =   52
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   38
      Left            =   360
      TabIndex        =   51
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   37
      Left            =   120
      TabIndex        =   50
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   36
      Left            =   2280
      TabIndex        =   49
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   35
      Left            =   2040
      TabIndex        =   48
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   34
      Left            =   1800
      TabIndex        =   47
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   33
      Left            =   1440
      TabIndex        =   46
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   32
      Left            =   1200
      TabIndex        =   45
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   31
      Left            =   960
      TabIndex        =   44
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   30
      Left            =   600
      TabIndex        =   43
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   29
      Left            =   360
      TabIndex        =   42
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   28
      Left            =   120
      TabIndex        =   41
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   27
      Left            =   2280
      TabIndex        =   40
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   26
      Left            =   2040
      TabIndex        =   39
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   25
      Left            =   1800
      TabIndex        =   38
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   24
      Left            =   1440
      TabIndex        =   37
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   23
      Left            =   1200
      TabIndex        =   36
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   22
      Left            =   960
      TabIndex        =   35
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   21
      Left            =   600
      TabIndex        =   34
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   20
      Left            =   360
      TabIndex        =   33
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   19
      Left            =   120
      TabIndex        =   32
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   18
      Left            =   2280
      TabIndex        =   31
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   17
      Left            =   2040
      TabIndex        =   30
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   16
      Left            =   1800
      TabIndex        =   29
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   15
      Left            =   1440
      TabIndex        =   28
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   14
      Left            =   1200
      TabIndex        =   27
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   13
      Left            =   960
      TabIndex        =   26
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   12
      Left            =   600
      TabIndex        =   25
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   11
      Left            =   360
      TabIndex        =   24
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   10
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   9
      Left            =   2280
      TabIndex        =   22
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   8
      Left            =   2040
      TabIndex        =   21
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   7
      Left            =   1800
      TabIndex        =   20
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   6
      Left            =   1440
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   1200
      TabIndex        =   18
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000009&
      Height          =   270
      Index           =   4
      Left            =   960
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   600
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   360
      TabIndex        =   15
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Index           =   13
      Left            =   3840
      TabIndex        =   13
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Index           =   12
      Left            =   3600
      TabIndex        =   12
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Index           =   11
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   225
      Index           =   10
      Left            =   3600
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Index           =   9
      Left            =   3120
      TabIndex        =   8
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Index           =   8
      Left            =   2880
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Index           =   7
      Left            =   2640
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Index           =   6
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Index           =   5
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Index           =   4
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Index           =   3
      Left            =   3105
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Index           =   2
      Left            =   2865
      TabIndex        =   1
      Top             =   480
      Width           =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   240
      Index           =   1
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   102
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   100
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "等待中..."
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置

Dim idx%, znum%
Dim S(9, 9) As Integer
Dim SudokuY As Integer
Dim SudokuNum As Integer
Dim temp(9) As String
Dim t
Dim is1 As Boolean
Dim useTime As Integer

Private Sub Command1_Click(Index As Integer)
    If Index > 0 And Index < 10 Then
        
        If Text1(idx).ForeColor <> &HC0C0FF Then Text1(idx).Text = Index
        Text1(idx).SetFocus
        Text1(idx).SelStart = 1
    Else
        Select Case Index
            Case Is = 0
                Text1((idx + 1) Mod 81).SetFocus
                Text1((idx + 1) Mod 81).SelStart = 1
            Case Is = 10
                Text1((idx + 72) Mod 81).SetFocus
                Text1((idx + 72) Mod 81).SelStart = 1
            Case Is = 11
                If (idx + 8) Mod 9 = 0 Then idx = idx + 9
                Text1((idx + 80) Mod 81).SetFocus
                Text1((idx + 80) Mod 81).SelStart = 1
            Case Is = 12
                Text1((idx + 9) Mod 81).SetFocus
                Text1((idx + 9) Mod 81).SelStart = 1
            Case Is = 13
                If idx Mod 9 = 0 Then idx = idx - 9
                Text1((idx + 82) Mod 81).SetFocus
                Text1((idx + 82) Mod 81).SelStart = 1
            Case Is = 14
                If Text1(idx).ForeColor <> &HC0C0FF Then Text1(idx).Text = ""
                    
                Text1(idx).SetFocus
                Text1(idx).SelStart = 1
               
            Case Else
        End Select
    End If
End Sub

Private Sub Command3_Click()
    Dim i%
    For i = 0 To 80
        If Text1(i).ForeColor <> &HC0C0FF Then Text1(i).Text = Empty: Text1(i).ForeColor = vbRed
    Next
    useTime = 0
End Sub

Private Sub Command4_Click()
    Dim zError As Boolean
    Dim i As Integer
    For i = 1 To 81
        zError = zCheck(i)
        If zError = False Then MsgBox "Fasle": Exit Sub
    Next
    MsgBox "True"
End Sub

Private Sub Command5_Click()
    Dim zmax$, i%, rnum%
        Label1 = "等待输入数据..."
        If is1 Then
            zmax = 40
            is1 = False
        Else
            zmax = InputBox("", "输入剩余数字量")
            
            If zmax = Empty Then Label1 = "": Exit Sub
            If Val(zmax) < 18 Then
                MsgBox "剩余数字这么少貌似不行哦..." & vbCrLf & vbCrLf & "介于18跟63之间可行", vbInformation, "提示"
                Label1 = "等待操作..."
                Exit Sub
            End If
            If Val(zmax) > 63 Then
                MsgBox "剩余字数太多了也不太好吧..." & vbCrLf & vbCrLf & "介于18跟63之间可行", vbInformation, "提示"
                Label1 = "等待操作..."
                Exit Sub
            End If
        End If
        Label1 = "正在生成中..."
cmd5:   RandomNum
        Call Command2_Click
        If Label1 = "NoSolution" Then GoTo cmd5
       
        znum = 81 - Val(zmax)
        For i = 1 To znum
            Randomize
cmd5_2:     rnum = Int(80 * Rnd)
            If Text1(rnum).Text <> "" Then Text1(rnum).Text = Empty Else GoTo cmd5_2
        Next
        
        For i = 0 To 80
            If Text1(i).Text <> Empty Then Text1(i).ForeColor = &HC0C0FF
        Next
        Label1 = "...完整生成"
        useTime = 0
End Sub

Private Sub Command6_Click()
    Dim i%, j%
    For i = 0 To 8
        For j = 0 To 8
            If S(j, i) <> Empty And S(j, i) <> 0 Then
                Text1((j + (i * 9) + 82) Mod 81).Text = S(j, i)
                If Text1((j + (i * 9) + 82) Mod 81).ForeColor <> &HC0C0FF Then Text1((j + (i * 9) + 82) Mod 81).ForeColor = vbBlack
            End If
        Next
    Next
End Sub



Private Sub Form_DblClick()
    Me.Height = 3540
    Me.Width = 3540
    Label1.Enabled = True
    Label1 = ""
    Text1(idx).SetFocus
End Sub

Private Sub Form_Load()
    Dim i%
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    For i = 0 To 80
        Text1(i).BackColor = vbWhite
        Text1(i).ForeColor = vbRed
        Text1(i).Alignment = 2
    Next
    For i = 1 To 9
        Command1(i).Caption = i
    Next
    idx = 1
    
    Timer1.Interval = 1000
    useTime = 0
    
    Command1(0).Caption = "Next"
    Command1(10).Caption = "↑"
    Command1(11).Caption = "←"
    Command1(12).Caption = "↓"
    Command1(13).Caption = "→"
    Command1(14).Caption = "SetNull"
    
    Command3.ToolTipText = "放弃当前已经输入的，重新开始本局"
    Command4.ToolTipText = "测试当前输入是否有误"
    Command5.ToolTipText = "随机算法生成可解的一局。不排除有多解的可能性"
    Command6.ToolTipText = "将预设的情况填入。当然解法并不一定是只有这一种"
    
    Command1(1).Left = Command1(0).Left: Command1(1).Top = Command1(0).Top + Command1(0).Height + 30
    Command1(0).Width = 780
    For i = 1 To 13
        With Command1(i)
            .Width = 240
            .Height = 240
        End With
    Next
    For i = 2 To 9
        With Command1(i)
            .Left = Command1(1).Left + 270 * ((i - 1) Mod 3)
            .Top = Command1(1).Top + 270 * Int((i - 1) / 3)
        End With
    Next
    Command1(14).Left = Command1(0).Left
    Command1(14).Top = Command1(9).Top + 270
    Command1(14).Width = Command1(0).Width
    Command1(14).Height = Command1(0).Height
    Command1(10).Top = Command1(14).Top + 270
    Command1(11).Top = Command1(10).Top + 270
    Command1(12).Top = Command1(10).Top + 270
    Command1(13).Top = Command1(10).Top + 270
    Command1(10).Left = Command1(2).Left
    Command1(11).Left = Command1(1).Left
    Command1(12).Left = Command1(2).Left
    Command1(13).Left = Command1(3).Left
    
    With Command3
        .Left = Command1(0).Left
        .Top = Command1(13).Top + 270
        .Height = 250
        .Width = Command1(0).Width
    End With
    With Command4
        .Left = Command1(0).Left
        .Top = Command3.Top + 270
        .Height = 250
        .Width = Command1(0).Width
    End With
    With Command5
        .Left = Command1(0).Left
        .Top = Command4.Top + 270
        .Height = 240
        .Width = Command1(0).Width
    End With
    With Command6
        .Left = Command1(0).Left
        .Top = Command5.Top + 270
        .Height = 240
        .Width = Command1(0).Width
    End With
    
    
    Me.Height = 3540
    Me.Width = 2715
    Me.Caption = "zSudoku Es"
    
    Label2 = "Let's Sudoku!! by ArrowBro"
    'Label2.ForeColor = vbWhite
    Label2.Enabled = False
    is1 = True
    Call Command5_Click
    Label1 = "双击显示操作区"
    Label1.Enabled = False
    
    
    
    
End Sub

Private Sub Text1_Change(Index As Integer)
    If Text1(Index).ForeColor = &HC0C0FF Then
    Else
        If zCheck(Index) Then Text1(Index).ForeColor = vbRed Else Text1(Index).ForeColor = vbBlue
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).ForeColor = &HC0C0FF Then
    Else
        If zCheck(Index) Then Text1(Index).ForeColor = vbRed Else Text1(Index).ForeColor = vbBlue
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 32: If Index < 80 Then Index = Index + 1 Else Index = 0
        Case 37: If Index > 0 Then Index = Index - 1 Else Index = 80
        Case 38: If (Index + 80) Mod 81 > 8 Then Index = Index - 9
        Case 39: If Index < 80 Then Index = Index + 1 Else Index = 0
        Case 40: If Index < 73 Then Index = (Index + 9) Mod 81
        Case 46: Text1(Index) = Empty
        Case 8: Text1(Index) = Empty
        Case 27:
            If MsgBox("重来？！", vbYesNo + vbQuestion, "提示") = vbYes Then Call Command3_Click
    End Select
    Text1((Index + 81) Mod 81).SetFocus
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Text1(Index).ForeColor = &HC0C0FF Then
        KeyAscii = 0
    Else
'        If KeyAscii <= 57 And KeyAscii >= 48 Then
'            Text1(Index).Text = ""
'            KeyAscii = KeyAscii
'        Else
'            KeyAscii = 0
'        End If
'    End If
    
  '  If KeyAscii = 32 Then KeyAscii = 0: GoTo 1
  '  If KeyAscii <= 48 Or KeyAscii > 57 Then KeyAscii = 0: GoTo 2
    If KeyAscii <= 57 And KeyAscii > 48 Then Text1(Index).Text = "": KeyAscii = KeyAscii Else KeyAscii = 0
'1     If Index < 80 Then Index = Index + 1 Else Index = 0
'2     Text1(Index).SetFocus
    End If
End Sub

Public Function zCheck(Index As Integer) As Boolean
    Dim i%, j%, k%
    i = Index
    zCheck = True
   ' For i = 1 To 81
        k = i Mod 81 '记录本身位置
        For j = 1 To 9 '行检查
            If k Mod 9 = 0 Then k = k - 9
            k = (k + 82) Mod 81
            If i Mod 81 = k Or Text1(i Mod 81).Text = "" Or Text1(k).Text = "" Then
            Else
                If Text1(i Mod 81).Text = Text1(k).Text Then
                    zCheck = False
                    'Text1(k).BackColor = vbGreen
                End If
            End If
        Next
        k = i Mod 81 '记录本身位置
        For j = 1 To 9 '列检查
            k = (k + 72) Mod 81
            If i Mod 81 = k Or Text1(i Mod 81).Text = "" Or Text1(k).Text = "" Then
            Else
                If Text1(i Mod 81).Text = Text1(k).Text Then
                    zCheck = False
                    'Text1(k).BackColor = vbGreen
                End If
            End If
        Next
        k = i Mod 81 '记录本身位置
        For j = 1 To 9 '块检查
            'k = 1
            If k = 0 Then k = 81
            k = k + 1 + IIf(k Mod 3 = 0, 6, 0)
            k = k - IIf(Int((k + 26) / 27) > Int((i + 26) / 27), 27, 0)
            k = (k + 81) Mod 81

            If i Mod 81 = k Or Text1(i Mod 81).Text = "" Or Text1(k).Text = "" Then
            Else
                If Text1(i Mod 81).Text = Text1(k).Text Then
                    zCheck = False
                    'Text1(k).BackColor = vbGreen
                End If
            End If
        Next
   ' Next
End Function



Public Function RandomNum() As Boolean
    '随机  Int(9 * Rnd + 1)
    Dim i%, tmp%, j%
    For i = 0 To 80
        Text1(i).Text = Empty
        Text1(i).BackColor = vbWhite
        Text1(i).ForeColor = vbRed
    Next
    Randomize   ' 对随机数生成器做初始化的动作。
    Text1(1).Text = Int(9 * Rnd + 1)
    For i = 2 To 9
s1:     tmp = Int(9 * Rnd + 1)

        For j = 1 To i - 1
            If tmp = Text1(j).Text Then GoTo s1
        Next
        Text1(i).Text = tmp
    Next
    
s2: tmp = Int(9 * Rnd + 1)
    If tmp = Text1(1).Text Or tmp = Text1(2).Text Or tmp = Text1(3).Text Then GoTo s2
    Text1(10).Text = tmp
s3: tmp = Int(9 * Rnd + 1)
    If tmp = Text1(1).Text Or tmp = Text1(2).Text Or tmp = Text1(3).Text Or tmp = Text1(10).Text Then GoTo s3
    Text1(19).Text = tmp

    For i = 4 To 9
s4:     tmp = Int(9 * Rnd + 1)

        For j = 1 To i - 1
            If tmp = Text1((j - 1) * 9 + 1).Text Then GoTo s4
            
        Next
     Text1((i - 1) * 9 + 1).Text = tmp
    Next
End Function

Private Sub Text1_LostFocus(Index As Integer)
    idx = Index
'    If Text1(Index).ForeColor = &HC0C0FF Then
'    Else
'        If zCheck(Index) Then Text1(Index).ForeColor = vbRed Else Text1(Index).ForeColor = vbBlue
'    End If
    If Text1(Index).ForeColor <> &HC0C0FF Then Text1(Index).ForeColor = vbRed
End Sub



Private Sub Command2_Click()
    Dim i%, k%, j%
    t = Timer
    '先将资料填入阵列中
    
    For i = 0 To 8
        For j = 0 To 8
            
            k = (j + (i * 9) + 82) Mod 81
            If Text1(k) = "" Then
                S(j, i) = 0
            Else
                S(j, i) = Text1(k)
            End If
        Next
        
    Next
    '开始解题
    SudokuY = 0
    SudokuNum = 1
    If Sudoku(SudokuY, SudokuNum) = False Then
        Label1 = "NoSolution"
    Else
        '将资料填到Text1中
        For i = 0 To 8
            For j = 0 To 8
                Text1((j + (i * 9) + 82) Mod 81) = S(j, i)
            Next
        Next
        Label1 = "Success"
    End If
 End Sub

Function Sudoku(y As Integer, Num As Integer)
    Dim i As Integer, j As Integer
    For i = 0 To 8
        If S(i, y) = Num Then GoTo ex
    Next
    '判断交叉格是否为空白
    i = 0
    Do While i <= 8
        Dim NextCol As Boolean
        NextCol = False
        '换判断直行
        For j = 0 To 8
            If S(i, j) = Num Then NextCol = True '此行已有Num，换跳下一行
        Next
        If NextCol = False Then
            If S(i, y) = 0 Then '判断交叉格是否为空白
            '判断九宫格中是否有Num
            Dim a As Integer, b As Integer
            For a = (i \ 3) * 3 To (i \ 3) * 3 + 2
                For b = (y \ 3) * 3 To (y \ 3) * 3 + 2
                    If S(a, b) = Num Then '已经有Num，换跳下一直行
                        NextCol = True
                    End If
                Next
            Next
            If NextCol = False Then '准备填入并且跳下一横列
                S(i, y) = Num
ex:             If Timer - t > 0.2 Then Sudoku = False: Exit Function
               '跳下一横列
                If SudokuY = 8 Then '如果第8列了
                    If SudokuNum = 9 Then '号码为9号了
                        Sudoku = True '成功跳出
                        Exit Function
                    Else '换下一号码
                        SudokuY = 0
                        SudokuNum = SudokuNum + 1
                    End If
                Else '换下一横列
                    SudokuY = SudokuY + 1
                End If
            
                If Sudoku(SudokuY, SudokuNum) = False Then
                   If Form1.Text1(i + (y * 9)) = "" Or Form1.Text1(i + (y * 9)) = "0" Then S(i, y) = 0
                   Else
                       GoTo ex
                   End If
                End If
            End If
        End If
        i = i + 1
    Loop
    '解题失败，回到上一步骤
    Sudoku = False
    If y = 0 Then
        SudokuNum = SudokuNum - 1
        SudokuY = 8
    Else
        SudokuY = SudokuY - 1
    End If
    If SudokuNum = 0 Then
        Sudoku = False
        Exit Function
    End If
End Function

Private Sub Timer1_Timer()
    Label3 = "耗时:" & useTime & "秒"
    useTime = useTime + 1
'    Dim zcou%
'    zcou = 0
'    For i = 0 To 80
'        If Text1(i).Text <> Empty And Text1(i).ForeColor = vbRed Then zcou = zcou + 1
'    Next
'    If zcou >= znum Then Label3 = ""
End Sub
