VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tetris"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   210
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'本游戏使用枚举法，如果希望使用矩阵变换来进行方块旋转，可以参考下面文章
'《tetris游戏关键技术探讨》 高凌琴
Option Explicit
'API
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
'常量
Const Cell = 30 'me.scaleMode = 3
Const Action_Speed = 80 '这个是响应速度 单位是ms 也就是说Action_Speed毫秒后操作会进行到下一帧
Const fps = 60 '60帧
Const Game_Speed = 500
'窗体
Dim form_Width As Integer
Dim form_Height As Integer
Dim form_Top As Integer
Dim form_Left As Integer
'框架frame
Dim frame_Width As Integer
Dim frame_Height As Integer
Dim frame_Top As Integer
Dim frame_Left As Integer
Dim TwipsPerPixelX As Long '像素和缇转换变量 不同显示器不一样
Dim TwipsPerPixelY As Long
'计分器
Dim Score As Long '给出长整型
Dim User_Action As String 'left right down change
Dim Game_State As String ' running /pause /stop /dead
'记录为位置数据
Dim currentCubes_X(3) As Integer
Dim currentCubes_Y(3) As Integer
Dim currentCubes_Mode As CubesMode
Dim currentCubes_Direction As CubesDirection
'下一个方块
Dim nextCubes_X(3) As Integer
Dim nextCubes_Y(3) As Integer
Dim nextCubes_Mode As CubesMode '记录形态
Dim nextCubes_Direction As CubesDirection
'改变前的形态
Dim oldCubes_X(3) As Integer
Dim oldCubes_Y(3) As Integer
Dim oldCubes_Mode As CubesMode
Dim oldCubes_Direction As CubesDirection
'新cubes
Dim newCubes_X(3) As Integer
Dim newCubes_Y(3) As Integer
Dim newCubes_Mode As CubesMode
Dim newCubes_Direction As CubesDirection
'blocks
Const Blocks_MaxIndex = 199 ' 从0开始
Dim Blocks_X(Blocks_MaxIndex) As Integer
Dim Blocks_Y(Blocks_MaxIndex) As Integer
Dim Blocks_Status(Blocks_MaxIndex)  As Integer
Dim Blocks_Color(Blocks_MaxIndex) As Long
'形态枚举
Private Enum CubesMode
    LineMode = 1
    CubeMode = 0
    LeftSevenMode = 2
    RightSevenMode = 3
    TMode = 4
    LeftZMode = 5
    RightZMode = 6
End Enum
'方块方向枚举
Private Enum CubesDirection
    UpDirection = 0
    DownDirection = 1
    RightDirection = 2
    LeftDirection = 3
End Enum

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    User_Action = "hold"
End Sub

'开始函数
Private Sub Form_Load()
    '获得当前显示器参数
    TwipsPerPixelX = Screen.TwipsPerPixelX
    TwipsPerPixelY = Screen.TwipsPerPixelY
    'Debug.Print TwipsPerPixelX, TwipsPerPixelY '显然本机为15
    'game frame'15/24 ≈ 1/0.618 为黄金分割比例
    frame_Top = 1
    frame_Left = 1
    frame_Width = 10
    frame_Height = 20
    '初始化窗体
    form_Width = (frame_Width + 6) * Cell * TwipsPerPixelX
    form_Height = (frame_Height + 3) * Cell * TwipsPerPixelY
    form_Top = 0
    form_Left = 0
    Me.Move Screen.Width / 3, form_Top, form_Width, form_Height
    Me.ForeColor = vbBlack
    Me.DrawWidth = 2
    '初始化当前方块
    Call NewRndCubes
    Call ShowCurrentCubes
    '画下一个方块
    Call NewRndCubes
    Call ShowNextCubes
    '重置Blocks
    Call ResetBlocks
    '画
    Call ReDrawUI
End Sub
'游戏循环
Private Sub Game_Loop()
    Dim Game_NowTime As Long
    Dim Game_NewTime As Long
    Dim Action_NowTime As Long
    Dim Action_NewTime As Long
    Dim Draw_NowTime As Long
    Dim Draw_NewTime As Long
    While DoEvents
        If Game_State = "running" Then
            '操作响应
            Action_NewTime = timeGetTime()
            If Action_NewTime - Action_NowTime >= Action_Speed Then
                Action_NowTime = Action_NewTime
                Call saveCubes
                Call switchCubes(User_Action)
            End If
            '画面响应
            Game_NewTime = timeGetTime()
            If Game_NewTime - Game_NowTime >= Game_Speed Then
                Game_NowTime = Game_NewTime
                Call saveCubes
                Call switchCubes("down")
            End If
            '画面刷新
            Draw_NewTime = timeGetTime()
            If Draw_NewTime - Draw_NowTime >= 1000 / fps Then
                Draw_NowTime = Draw_NewTime
                Call ReDrawUI
            End If
        ElseIf Game_State = "stop" Then
            Exit Sub
        End If
        Sleep 1
    Wend
End Sub
'重画界面
Private Sub ReDrawUI()
    'Me.Cls
    Call DrawWhiteBackColor
    Call DrawWall
    Call DrawBlocks
    Call DrawNextCubes
    Call DrawCurrentCubes
End Sub
'达底判定
Private Function HitButtom() As Boolean
    Dim i As Integer
    '触底判断
    For i = 0 To 3
        If currentCubes_Y(i) > frame_Height Then
            HitButtom = True
            Exit Function
        End If
    Next i
End Function
'碰撞Blocks,'返回碰撞的blocks的id
Private Function HitBlocks() As Boolean
    Dim i As Integer, j As Integer
    For i = 0 To Blocks_MaxIndex
        If Blocks_Status(i) = 1 Then
            For j = 0 To 3
                If currentCubes_X(j) = Blocks_X(i) And currentCubes_Y(j) = Blocks_Y(i) Then
                    HitBlocks = True
                    Exit Function
                End If
            Next
        End If
    Next
End Function
'重置Blocks
Private Sub ResetBlocks()
    Dim i As Integer, j As Integer
    Dim X As Integer, Y As Integer
    For i = 0 To Blocks_MaxIndex
        X = i Mod frame_Width + frame_Left
        Y = i \ frame_Width + frame_Top
        Blocks_X(i) = X
        Blocks_Y(i) = Y
        Blocks_Status(i) = 0
    Next
End Sub
'消除判断 返回可消除的ROW
Private Function CheckBlocks() As Integer
    Dim i As Integer
    Dim Y As Integer, X As Integer
    Dim C As Integer
    For Y = frame_Height To frame_Top Step -1
        For X = frame_Width To frame_Left Step -1
            i = X + (Y - 1) * frame_Width - 1
            If Blocks_Status(i) = 1 Then
                C = C + 1
            End If
            
        Next
        If C >= 10 Then
            CheckBlocks = Y
            Exit Function
        Else
            '重新计算
            C = 0
        End If
    Next
End Function
'下移函数，即上方的方块状态复制到下方
Private Sub MoveBlocksStatus(ByVal Row As Integer)
    Dim X As Integer, Y As Integer
    Dim i As Integer, j As Integer
    '从第ROW行开始整体复制下移
    For Y = Row To frame_Top Step -1
        For X = frame_Width To frame_Left Step -1
            If Y > 1 Then
                i = X + (Y - 1) * frame_Width - 1
                j = X + (Y - 2) * frame_Width - 1 '上一行
                Blocks_Status(i) = 0 '清空当前行
                Blocks_Status(i) = Blocks_Status(j)
                Blocks_Color(i) = Blocks_Color(j) '把颜色也拉过来
                Blocks_Status(j) = 0 '清空上一行
            Else
                Blocks_Status(i) = 0 '最高行直接消除
            End If
        Next X
    Next Y
    '积分1分
    Score = Score + 1
End Sub
'复制当前方块到Blocks中
'在这里作死亡判断？
Private Sub CopyToBlocks()
    Dim i As Integer, j As Integer
    For i = 0 To Blocks_MaxIndex
            For j = 0 To 3
                If currentCubes_X(j) = Blocks_X(i) And currentCubes_Y(j) = Blocks_Y(i) Then
                    Blocks_Status(i) = 1
                    Blocks_Color(i) = CurrentCubesColor
                End If
            Next
    Next
End Sub
'操作函数
Private Sub switchCubes(ByVal moveDirection As String)
    Dim i As Integer
    '开始对比中
    If moveDirection = "left" Then
        For i = 0 To 3
            currentCubes_X(i) = currentCubes_X(i) - 1
        Next i
    ElseIf moveDirection = "right" Then
        For i = 0 To 3
            currentCubes_X(i) = currentCubes_X(i) + 1
        Next i
    ElseIf moveDirection = "up" Then
        For i = 0 To 3
            currentCubes_Y(i) = currentCubes_Y(i) - 1
        Next i
    ElseIf moveDirection = "down" Then
        For i = 0 To 3
            currentCubes_Y(i) = currentCubes_Y(i) + 1
        Next i
    ElseIf moveDirection = "rotate" Then
        
    End If
    '触底或者时遇到blocks
    If HitButtom = True Or HitBlocks = True Then
        Dim Row As Integer
        Call BackCubes
        Call CopyToBlocks
        '产生新的方块
        Call NextToCurrent
        '产生新的下一位方块
        Call ClsNextCubes
        Call NewRndCubes
        Call ShowNextCubes
        Call DrawNextCubes
        '重复检查直到没有
        While CheckBlocks > 0
            Row = CheckBlocks
            Call MoveBlocksStatus(Row)
            '如果有消除那么就可以重新画了
        Wend
        Exit Sub
    End If
    '判断是否回退
    If CubeInFrame = False Then
        Call BackCubes
        Exit Sub
    End If
End Sub
'判断是否在框架内
Private Function CubeInFrame() As Boolean
    Dim i As Integer
    For i = 0 To 3
        If currentCubes_X(i) >= frame_Left And currentCubes_X(i) <= frame_Width Then
            CubeInFrame = True
        Else
            CubeInFrame = False
            Exit Function
        End If
    Next
End Function
'控制台
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Game_State = "running" Then
        If KeyCode = vbKeySpace Then
            Call saveCubes
            Call rotateCubes '改变方向
            Call ReDrawUI
        ElseIf KeyCode = vbKeyA Or KeyCode = vbKeyLeft Then
            User_Action = "left"
        ElseIf KeyCode = vbKeyD Or KeyCode = vbKeyRight Then
            User_Action = "right"
        ElseIf KeyCode = vbKeyS Or KeyCode = vbKeyDown Then
            User_Action = "down"
        End If
    End If
End Sub
'随机产生Cubes
Private Sub NewRndCubes()
    Dim i As Integer
    Dim Cube0_X As Integer, Cube0_Y As Integer
    Dim mDirection As CubesDirection
    '初始方向
    newCubes_Direction = UpDirection
    '得到形状
    Randomize
    newCubes_Mode = Int(Rnd * 7)
    '确定方块坐标及方向
    Randomize
    mDirection = Int(Rnd * 4)
    '确定初始坐标
    Cube0_X = 5
    Cube0_Y = 0
    Select Case newCubes_Mode
        Case CubeMode
            newCubes_X(0) = Cube0_X
            newCubes_Y(0) = Cube0_Y
            newCubes_X(1) = Cube0_X + 1
            newCubes_Y(1) = Cube0_Y
            newCubes_X(2) = Cube0_X
            newCubes_Y(2) = Cube0_Y + 1
            newCubes_X(3) = Cube0_X + 1
            newCubes_Y(3) = Cube0_Y + 1
        Case LineMode
            '确定up型方块
                newCubes_X(0) = Cube0_X
                newCubes_Y(0) = Cube0_Y
                newCubes_X(1) = Cube0_X + 1
                newCubes_Y(1) = Cube0_Y
                newCubes_X(2) = Cube0_X + 2
                newCubes_Y(2) = Cube0_Y
                newCubes_X(3) = Cube0_X + 3
                newCubes_Y(3) = Cube0_Y
            If mDirection = UpDirection Or mDirection = DownDirection Then
                '无需旋转
            ElseIf mDirection = LeftDirection Or mDirection = RightDirection Then
                '一次旋转，变换形态
                Call rotateNewCubes
            Else
                MsgBox "随机数可能算错了！", vbCritical, "错误"
            End If
        Case LeftZMode
                newCubes_X(0) = Cube0_X
                newCubes_Y(0) = Cube0_Y
                newCubes_X(1) = Cube0_X + 1
                newCubes_Y(1) = Cube0_Y
                newCubes_X(2) = Cube0_X + 1
                newCubes_Y(2) = Cube0_Y + 1
                newCubes_X(3) = Cube0_X + 2
                newCubes_Y(3) = Cube0_Y + 1
            If mDirection = UpDirection Or mDirection = DownDirection Then
                'up
            ElseIf mDirection = LeftDirection Or mDirection = RightDirection Then
                Call rotateNewCubes
            Else
                MsgBox "随机数可能算错了！", vbCritical, "错误"
            End If
        Case RightZMode
                newCubes_X(0) = Cube0_X
                newCubes_Y(0) = Cube0_Y
                newCubes_X(1) = Cube0_X - 1
                newCubes_Y(1) = Cube0_Y
                newCubes_X(2) = Cube0_X - 1
                newCubes_Y(2) = Cube0_Y + 1
                newCubes_X(3) = Cube0_X - 2
                newCubes_Y(3) = Cube0_Y + 1
            If mDirection = UpDirection Or mDirection = DownDirection Then
                'up
            ElseIf mDirection = LeftDirection Or mDirection = RightDirection Then
                Call rotateNewCubes
            Else
                MsgBox "随机数可能算错了！", vbCritical, "错误"
            End If
        Case TMode
                newCubes_X(0) = Cube0_X
                newCubes_Y(0) = Cube0_Y
                newCubes_X(1) = Cube0_X
                newCubes_Y(1) = Cube0_Y + 1
                newCubes_X(2) = Cube0_X - 1
                newCubes_Y(2) = Cube0_Y + 1
                newCubes_X(3) = Cube0_X + 1
                newCubes_Y(3) = Cube0_Y + 1
            If mDirection = UpDirection Then
                'up
            ElseIf mDirection = RightDirection Then
                Call rotateNewCubes
            ElseIf mDirection = DownDirection Then
                Call rotateNewCubes
                Call rotateNewCubes
            ElseIf mDirection = LeftDirection Then
                Call rotateNewCubes
                Call rotateNewCubes
                Call rotateNewCubes
            End If
            'Debug.Print "TMode direction : mode", newCubes_Direction, newCubes_Mode
        Case LeftSevenMode
                newCubes_X(0) = Cube0_X
                newCubes_Y(0) = Cube0_Y
                newCubes_X(1) = Cube0_X
                newCubes_Y(1) = Cube0_Y + 1
                newCubes_X(2) = Cube0_X - 1
                newCubes_Y(2) = Cube0_Y + 1
                newCubes_X(3) = Cube0_X - 2
                newCubes_Y(3) = Cube0_Y + 1
                If mDirection = UpDirection Then
                    'up
                ElseIf mDirection = RightDirection Then
                    Call rotateNewCubes
                ElseIf mDirection = DownDirection Then
                    Call rotateNewCubes
                    Call rotateNewCubes 'down 由up 经过两次变换得到
                ElseIf mDirection = LeftDirection Then
                    Call rotateNewCubes
                    Call rotateNewCubes
                    Call rotateNewCubes 'left三次变换得到
                End If
        Case RightSevenMode
                newCubes_X(0) = Cube0_X
                newCubes_Y(0) = Cube0_Y
                newCubes_X(1) = Cube0_X
                newCubes_Y(1) = Cube0_Y + 1
                newCubes_X(2) = Cube0_X + 1
                newCubes_Y(2) = Cube0_Y + 1
                newCubes_X(3) = Cube0_X + 2
                newCubes_Y(3) = Cube0_Y + 1
                If mDirection = UpDirection Then
                    'up
                ElseIf mDirection = RightDirection Then
                    Call rotateNewCubes
                ElseIf mDirection = DownDirection Then
                    Call rotateNewCubes
                    Call rotateNewCubes 'down 由up 经过两次变换得到
                ElseIf mDirection = LeftDirection Then
                    Call rotateNewCubes
                    Call rotateNewCubes
                    Call rotateNewCubes 'left三次变换得到
                End If
    End Select
End Sub

'保存currentCubes到oldcubes
Private Sub saveCubes()
    Dim i As Integer
    For i = 0 To 3
        oldCubes_X(i) = currentCubes_X(i)
        oldCubes_Y(i) = currentCubes_Y(i)
    Next
    oldCubes_Mode = currentCubes_Mode
    oldCubes_Direction = currentCubes_Direction
End Sub
'把随机方块转到当前
Private Sub ShowCurrentCubes()
    Dim i As Integer
    For i = 0 To 3
        currentCubes_X(i) = newCubes_X(i)
        currentCubes_Y(i) = newCubes_Y(i)
    Next
    currentCubes_Mode = newCubes_Mode
    currentCubes_Direction = newCubes_Direction
End Sub
'BackCubes()
Private Sub BackCubes()
    Dim i As Integer
    For i = 0 To 3
        currentCubes_X(i) = oldCubes_X(i)
        currentCubes_Y(i) = oldCubes_Y(i)
    Next i
    currentCubes_Mode = oldCubes_Mode
    currentCubes_Direction = oldCubes_Direction
End Sub
'展示下一个方块
Private Sub ShowNextCubes()
    Dim i As Integer
    For i = 0 To 3
        nextCubes_X(i) = newCubes_X(i) + 7
        nextCubes_Y(i) = newCubes_Y(i) + 2
    Next
    nextCubes_Mode = newCubes_Mode
    nextCubes_Direction = newCubes_Direction
End Sub
'把下一个方块放到当前
Private Sub NextToCurrent()
    Dim X As Integer, Y As Integer
    Dim i As Integer
    currentCubes_Mode = nextCubes_Mode
    currentCubes_Direction = nextCubes_Direction
    '位置的话需要移动至最顶部和最中央
    For i = 0 To 3
        currentCubes_X(i) = nextCubes_X(i) - 9
        currentCubes_Y(i) = nextCubes_Y(i)
    Next i
End Sub
'展示下一个方块
Private Sub DrawNextCubes()
    Dim i As Integer
    Dim ModeStr As String
    Dim ModeColor As Long
    '根据不同的方块模型选择颜色
    Select Case nextCubes_Mode
        Case CubeMode
            ModeColor = RGB(255, 174, 0) 'yellow
        Case LineMode
            ModeColor = RGB(47, 155, 255) 'skyblue
        Case LeftZMode
            ModeColor = RGB(222, 41, 44) 'red
        Case RightZMode
            ModeColor = RGB(11, 171, 20) 'green
        Case TMode
            ModeColor = RGB(160, 32, 240) 'purple
        Case LeftSevenMode
            ModeColor = RGB(238, 154, 0) 'orange
        Case RightSevenMode
            ModeColor = RGB(43, 83, 173) 'blue
    End Select
    For i = 0 To 3
        Call DrawCell(nextCubes_X(i), nextCubes_Y(i), ModeColor)
    Next
End Sub
'画出当前的Blocks，可视化就是这么简单
Private Sub DrawBlocks()
    Dim i As Integer
    For i = 0 To Blocks_MaxIndex
        If Blocks_Status(i) = 1 Then  '1就是有了
            Call DrawCell(Blocks_X(i), Blocks_Y(i), Blocks_Color(i))
        End If
    Next
End Sub
'画白色背景
Private Sub DrawWhiteBackColor()
    Me.Line (form_Left, form_Top)-(form_Width, form_Height), vbWhite, BF
End Sub
'画边框及其它
Private Sub DrawWall()
    Dim X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer
    X1 = frame_Left * Cell
    Y1 = frame_Top * Cell
    '因为需要其外墙，所以需要加1
    X2 = frame_Width * Cell + Cell
    Y2 = frame_Height * Cell + Cell
    '墙体
    Me.Line (X1, Y1)-(X2, Y2), vbBlack, B
End Sub
'旋转方块
Private Function rotateCubes() As Boolean
    Select Case currentCubes_Direction
    Case UpDirection
        Select Case currentCubes_Mode
        Case LineMode 'linemode只有两种 up left
            currentCubes_Direction = LeftDirection
            currentCubes_X(0) = oldCubes_X(0) + 2
            currentCubes_Y(0) = oldCubes_Y(0) - 2
            currentCubes_X(1) = oldCubes_X(1) + 1
            currentCubes_Y(1) = oldCubes_Y(1) - 1
            currentCubes_X(3) = oldCubes_X(3) - 1
            currentCubes_Y(3) = oldCubes_Y(3) + 1
        Case LeftSevenMode
            currentCubes_Direction = RightDirection
            currentCubes_X(0) = oldCubes_X(0) - 1
            currentCubes_Y(0) = oldCubes_Y(0) + 1
            currentCubes_X(2) = oldCubes_X(2) + 1
            currentCubes_Y(2) = oldCubes_Y(2) + 1
            currentCubes_X(3) = oldCubes_X(3) + 2
            currentCubes_Y(3) = oldCubes_Y(3) + 2
        Case RightSevenMode '7字
            currentCubes_Direction = RightDirection
            currentCubes_X(0) = oldCubes_X(0) - 1
            currentCubes_Y(0) = oldCubes_Y(0) + 1
            currentCubes_X(2) = oldCubes_X(2) - 1
            currentCubes_Y(2) = oldCubes_Y(2) - 1
            currentCubes_X(3) = oldCubes_X(3) - 2
            currentCubes_Y(3) = oldCubes_Y(3) - 2
        Case LeftZMode
            currentCubes_Direction = LeftDirection
            currentCubes_X(0) = oldCubes_X(0) + 1
            currentCubes_Y(0) = oldCubes_Y(0) - 1
            currentCubes_X(2) = oldCubes_X(2) - 1
            currentCubes_Y(2) = oldCubes_Y(2) - 1
            currentCubes_X(3) = oldCubes_X(3) - 2
            currentCubes_Y(3) = oldCubes_Y(3)
        Case RightZMode
            currentCubes_Direction = LeftDirection
            currentCubes_X(0) = oldCubes_X(0) - 1
            currentCubes_Y(0) = oldCubes_Y(0) - 1
            currentCubes_X(2) = oldCubes_X(2) + 1
            currentCubes_Y(2) = oldCubes_Y(2) - 1
            currentCubes_X(3) = oldCubes_X(3) + 2
            currentCubes_Y(3) = oldCubes_Y(3)
        Case TMode 'T字型
            currentCubes_Direction = RightDirection
            currentCubes_X(0) = oldCubes_X(0) - 1
            currentCubes_Y(0) = oldCubes_Y(0) + 1
            currentCubes_X(2) = oldCubes_X(2) + 1
            currentCubes_Y(2) = oldCubes_Y(2) + 1
            currentCubes_X(3) = oldCubes_X(3) - 1
            currentCubes_Y(3) = oldCubes_Y(3) - 1
        End Select
    Case DownDirection
        Select Case currentCubes_Mode
            Case LeftSevenMode
                currentCubes_Direction = LeftDirection
                currentCubes_X(0) = oldCubes_X(0) + 1
                currentCubes_Y(0) = oldCubes_Y(0) - 1
                currentCubes_X(2) = oldCubes_X(2) - 1
                currentCubes_Y(2) = oldCubes_Y(2) - 1
                currentCubes_X(3) = oldCubes_X(3) - 2
                currentCubes_Y(3) = oldCubes_Y(3) - 2
            Case RightSevenMode '7字
                currentCubes_Direction = LeftDirection
                currentCubes_X(0) = oldCubes_X(0) + 1
                currentCubes_Y(0) = oldCubes_Y(0) - 1
                currentCubes_X(2) = oldCubes_X(2) + 1
                currentCubes_Y(2) = oldCubes_Y(2) + 1
                currentCubes_X(3) = oldCubes_X(3) + 2
                currentCubes_Y(3) = oldCubes_Y(3) + 2
            Case TMode 'T字型
                currentCubes_Direction = LeftDirection
                currentCubes_X(0) = oldCubes_X(0) + 1
                currentCubes_Y(0) = oldCubes_Y(0) - 1
                currentCubes_X(2) = oldCubes_X(2) - 1
                currentCubes_Y(2) = oldCubes_Y(2) - 1
                currentCubes_X(3) = oldCubes_X(3) + 1
                currentCubes_Y(3) = oldCubes_Y(3) + 1
        End Select
    Case LeftDirection
        Select Case currentCubes_Mode
            Case LineMode
                currentCubes_Direction = UpDirection  '上 右 下 左
                currentCubes_X(0) = oldCubes_X(0) - 2
                currentCubes_Y(0) = oldCubes_Y(0) + 2
                currentCubes_X(1) = oldCubes_X(1) - 1
                currentCubes_Y(1) = oldCubes_Y(1) + 1
                currentCubes_X(3) = oldCubes_X(3) + 1
                currentCubes_Y(3) = oldCubes_Y(3) - 1
            Case LeftSevenMode '7字
                currentCubes_Direction = UpDirection
                currentCubes_X(0) = oldCubes_X(0) - 1
                currentCubes_Y(0) = oldCubes_Y(0) - 1
                currentCubes_X(2) = oldCubes_X(2) - 1
                currentCubes_Y(2) = oldCubes_Y(2) + 1
                currentCubes_X(3) = oldCubes_X(3) - 2
                currentCubes_Y(3) = oldCubes_Y(3) + 2
            Case RightSevenMode '7字
                currentCubes_Direction = UpDirection
                currentCubes_X(0) = oldCubes_X(0) - 1
                currentCubes_Y(0) = oldCubes_Y(0) - 1
                currentCubes_X(2) = oldCubes_X(2) + 1
                currentCubes_Y(2) = oldCubes_Y(2) - 1
                currentCubes_X(3) = oldCubes_X(3) + 2
                currentCubes_Y(3) = oldCubes_Y(3) - 2
            Case LeftZMode '左Z型
                currentCubes_Direction = UpDirection
                currentCubes_X(0) = oldCubes_X(0) - 1
                currentCubes_Y(0) = oldCubes_Y(0) + 1
                currentCubes_X(2) = oldCubes_X(2) + 1
                currentCubes_Y(2) = oldCubes_Y(2) + 1
                currentCubes_X(3) = oldCubes_X(3) + 2
                currentCubes_Y(3) = oldCubes_Y(3)
            Case RightZMode
                currentCubes_Direction = UpDirection
                currentCubes_X(0) = oldCubes_X(0) + 1
                currentCubes_Y(0) = oldCubes_Y(0) + 1
                currentCubes_X(2) = oldCubes_X(2) - 1
                currentCubes_Y(2) = oldCubes_Y(2) + 1
                currentCubes_X(3) = oldCubes_X(3) - 2
                currentCubes_Y(3) = oldCubes_Y(3)
            Case TMode 'T字型
                currentCubes_Direction = UpDirection
                currentCubes_X(0) = oldCubes_X(0) - 1
                currentCubes_Y(0) = oldCubes_Y(0) - 1
                currentCubes_X(2) = oldCubes_X(2) - 1
                currentCubes_Y(2) = oldCubes_Y(2) + 1
                currentCubes_X(3) = oldCubes_X(3) + 1
                currentCubes_Y(3) = oldCubes_Y(3) - 1
        End Select
    Case RightDirection
        Select Case currentCubes_Mode
            Case LeftSevenMode
                currentCubes_Direction = DownDirection
                currentCubes_X(0) = oldCubes_X(0) + 1
                currentCubes_Y(0) = oldCubes_Y(0) + 1
                currentCubes_X(2) = oldCubes_X(2) + 1
                currentCubes_Y(2) = oldCubes_Y(2) - 1
                currentCubes_X(3) = oldCubes_X(3) + 2
                currentCubes_Y(3) = oldCubes_Y(3) - 2
            Case RightSevenMode '7字
                currentCubes_Direction = DownDirection
                currentCubes_X(0) = oldCubes_X(0) + 1
                currentCubes_Y(0) = oldCubes_Y(0) + 1
                currentCubes_X(2) = oldCubes_X(2) - 1
                currentCubes_Y(2) = oldCubes_Y(2) + 1
                currentCubes_X(3) = oldCubes_X(3) - 2
                currentCubes_Y(3) = oldCubes_Y(3) + 2
        Case TMode 'T字型
            currentCubes_Direction = DownDirection
            currentCubes_X(0) = oldCubes_X(0) + 1
            currentCubes_Y(0) = oldCubes_Y(0) + 1
            currentCubes_X(2) = oldCubes_X(2) + 1
            currentCubes_Y(2) = oldCubes_Y(2) - 1
            currentCubes_X(3) = oldCubes_X(3) - 1
            currentCubes_Y(3) = oldCubes_Y(3) + 1
        End Select
    End Select
End Function

'旋转方块
Private Function rotateNewCubes() As Boolean
    Select Case newCubes_Direction
    Case UpDirection
        Select Case newCubes_Mode
        Case LineMode 'linemode只有两种 up left
            newCubes_Direction = LeftDirection
            newCubes_X(0) = newCubes_X(0) + 2
            newCubes_Y(0) = newCubes_Y(0) - 2
            newCubes_X(1) = newCubes_X(1) + 1
            newCubes_Y(1) = newCubes_Y(1) - 1
            newCubes_X(3) = newCubes_X(3) - 1
            newCubes_Y(3) = newCubes_Y(3) + 1
        Case LeftSevenMode
            newCubes_Direction = RightDirection
            newCubes_X(0) = newCubes_X(0) - 1
            newCubes_Y(0) = newCubes_Y(0) + 1
            newCubes_X(2) = newCubes_X(2) + 1
            newCubes_Y(2) = newCubes_Y(2) + 1
            newCubes_X(3) = newCubes_X(3) + 2
            newCubes_Y(3) = newCubes_Y(3) + 2
        Case RightSevenMode '7字
            newCubes_Direction = RightDirection
            newCubes_X(0) = newCubes_X(0) - 1
            newCubes_Y(0) = newCubes_Y(0) + 1
            newCubes_X(2) = newCubes_X(2) - 1
            newCubes_Y(2) = newCubes_Y(2) - 1
            newCubes_X(3) = newCubes_X(3) - 2
            newCubes_Y(3) = newCubes_Y(3) - 2
        Case LeftZMode
            newCubes_Direction = LeftDirection
            newCubes_X(0) = newCubes_X(0) + 1
            newCubes_Y(0) = newCubes_Y(0) - 1
            newCubes_X(2) = newCubes_X(2) - 1
            newCubes_Y(2) = newCubes_Y(2) - 1
            newCubes_X(3) = newCubes_X(3) - 2
            newCubes_Y(3) = newCubes_Y(3)
        Case RightZMode
            newCubes_Direction = LeftDirection
            newCubes_X(0) = newCubes_X(0) - 1
            newCubes_Y(0) = newCubes_Y(0) - 1
            newCubes_X(2) = newCubes_X(2) + 1
            newCubes_Y(2) = newCubes_Y(2) - 1
            newCubes_X(3) = newCubes_X(3) + 2
            newCubes_Y(3) = newCubes_Y(3)
        Case TMode 'T字型
            newCubes_Direction = RightDirection
            newCubes_X(0) = newCubes_X(0) - 1
            newCubes_Y(0) = newCubes_Y(0) + 1
            newCubes_X(2) = newCubes_X(2) + 1
            newCubes_Y(2) = newCubes_Y(2) + 1
            newCubes_X(3) = newCubes_X(3) - 1
            newCubes_Y(3) = newCubes_Y(3) - 1
        End Select
    Case DownDirection
        Select Case newCubes_Mode
            Case LeftSevenMode
                newCubes_Direction = LeftDirection
                newCubes_X(0) = newCubes_X(0) + 1
                newCubes_Y(0) = newCubes_Y(0) - 1
                newCubes_X(2) = newCubes_X(2) - 1
                newCubes_Y(2) = newCubes_Y(2) - 1
                newCubes_X(3) = newCubes_X(3) - 2
                newCubes_Y(3) = newCubes_Y(3) - 2
            Case RightSevenMode '7字
                newCubes_Direction = LeftDirection
                newCubes_X(0) = newCubes_X(0) + 1
                newCubes_Y(0) = newCubes_Y(0) - 1
                newCubes_X(2) = newCubes_X(2) + 1
                newCubes_Y(2) = newCubes_Y(2) + 1
                newCubes_X(3) = newCubes_X(3) + 2
                newCubes_Y(3) = newCubes_Y(3) + 2
            Case TMode 'T字型
                newCubes_Direction = LeftDirection
                newCubes_X(0) = newCubes_X(0) + 1
                newCubes_Y(0) = newCubes_Y(0) - 1
                newCubes_X(2) = newCubes_X(2) - 1
                newCubes_Y(2) = newCubes_Y(2) - 1
                newCubes_X(3) = newCubes_X(3) + 1
                newCubes_Y(3) = newCubes_Y(3) + 1
        End Select
    Case LeftDirection
        Select Case newCubes_Mode
            Case LineMode
                newCubes_Direction = UpDirection  '上 右 下 左
                newCubes_X(0) = newCubes_X(0) - 2
                newCubes_Y(0) = newCubes_Y(0) + 2
                newCubes_X(1) = newCubes_X(1) - 1
                newCubes_Y(1) = newCubes_Y(1) + 1
                newCubes_X(3) = newCubes_X(3) + 1
                newCubes_Y(3) = newCubes_Y(3) - 1
            Case LeftSevenMode '7字
                newCubes_Direction = UpDirection
                newCubes_X(0) = newCubes_X(0) - 1
                newCubes_Y(0) = newCubes_Y(0) - 1
                newCubes_X(2) = newCubes_X(2) - 1
                newCubes_Y(2) = newCubes_Y(2) + 1
                newCubes_X(3) = newCubes_X(3) - 2
                newCubes_Y(3) = newCubes_Y(3) + 2
            Case RightSevenMode '7字
                newCubes_Direction = UpDirection
                newCubes_X(0) = newCubes_X(0) - 1
                newCubes_Y(0) = newCubes_Y(0) - 1
                newCubes_X(2) = newCubes_X(2) + 1
                newCubes_Y(2) = newCubes_Y(2) - 1
                newCubes_X(3) = newCubes_X(3) + 2
                newCubes_Y(3) = newCubes_Y(3) - 2
            Case LeftZMode '左Z型
                newCubes_Direction = UpDirection
                newCubes_X(0) = newCubes_X(0) - 1
                newCubes_Y(0) = newCubes_Y(0) + 1
                newCubes_X(2) = newCubes_X(2) + 1
                newCubes_Y(2) = newCubes_Y(2) + 1
                newCubes_X(3) = newCubes_X(3) + 2
                newCubes_Y(3) = newCubes_Y(3)
            Case RightZMode
                newCubes_Direction = UpDirection
                newCubes_X(0) = newCubes_X(0) + 1
                newCubes_Y(0) = newCubes_Y(0) + 1
                newCubes_X(2) = newCubes_X(2) - 1
                newCubes_Y(2) = newCubes_Y(2) + 1
                newCubes_X(3) = newCubes_X(3) - 2
                newCubes_Y(3) = newCubes_Y(3)
            Case TMode 'T字型
                newCubes_Direction = UpDirection
                newCubes_X(0) = newCubes_X(0) - 1
                newCubes_Y(0) = newCubes_Y(0) - 1
                newCubes_X(2) = newCubes_X(2) - 1
                newCubes_Y(2) = newCubes_Y(2) + 1
                newCubes_X(3) = newCubes_X(3) + 1
                newCubes_Y(3) = newCubes_Y(3) - 1
        End Select
    Case RightDirection
        Select Case newCubes_Mode
            Case LeftSevenMode
                newCubes_Direction = DownDirection
                newCubes_X(0) = newCubes_X(0) + 1
                newCubes_Y(0) = newCubes_Y(0) + 1
                newCubes_X(2) = newCubes_X(2) + 1
                newCubes_Y(2) = newCubes_Y(2) - 1
                newCubes_X(3) = newCubes_X(3) + 2
                newCubes_Y(3) = newCubes_Y(3) - 2
            Case RightSevenMode '7字
                newCubes_Direction = DownDirection
                newCubes_X(0) = newCubes_X(0) + 1
                newCubes_Y(0) = newCubes_Y(0) + 1
                newCubes_X(2) = newCubes_X(2) - 1
                newCubes_Y(2) = newCubes_Y(2) + 1
                newCubes_X(3) = newCubes_X(3) - 2
                newCubes_Y(3) = newCubes_Y(3) + 2
        Case TMode 'T字型
            newCubes_Direction = DownDirection
            newCubes_X(0) = newCubes_X(0) + 1
            newCubes_Y(0) = newCubes_Y(0) + 1
            newCubes_X(2) = newCubes_X(2) + 1
            newCubes_Y(2) = newCubes_Y(2) - 1
            newCubes_X(3) = newCubes_X(3) - 1
            newCubes_Y(3) = newCubes_Y(3) + 1
        End Select
    End Select
End Function
'画方块
Private Function DrawCurrentCubes() As Boolean
    Dim i As Integer
    Dim ModeStr As String
    For i = 0 To 3
        Call DrawCell(currentCubes_X(i), currentCubes_Y(i), CurrentCubesColor) '0
    Next
    'Debug.Print "Mode:Direction", currentCubes_Mode, currentCubes_Direction
End Function
'获得当前方块的颜色
Private Function CurrentCubesColor() As Long
    Dim ModeColor  As Long
    Select Case currentCubes_Mode
        Case CubeMode
            ModeColor = RGB(255, 174, 0) 'yellow
        Case LineMode
            ModeColor = RGB(47, 155, 255) 'skyblue
        Case LeftZMode
            ModeColor = RGB(222, 41, 44) 'red
        Case RightZMode
            ModeColor = RGB(11, 171, 20) 'green
        Case TMode
            ModeColor = RGB(160, 32, 240) 'purple
        Case LeftSevenMode
            ModeColor = RGB(238, 154, 0) 'orange
        Case RightSevenMode
            ModeColor = RGB(43, 83, 173) 'blue
    End Select
    CurrentCubesColor = ModeColor
End Function
'删除方块
Private Function ClsOldCubes()
    Dim i As Integer
    For i = 0 To 3
        Call ClsCell(oldCubes_X(i), oldCubes_Y(i))
    Next
End Function
'删除当前方块
Private Function ClsCurrentCubes()
    Dim i As Integer
    For i = 0 To 3
        Call ClsCell(currentCubes_X(i), currentCubes_Y(i))
    Next
End Function
'删除当前的ClsNextCubes
Private Sub ClsNextCubes()
    Dim i As Integer
    For i = 0 To 3
        Call ClsCell(nextCubes_X(i), nextCubes_Y(i))
    Next
End Sub
'删除细胞
Private Function ClsCell(ByVal Cell_X As Integer, ByVal Cell_Y As Integer)
    Dim X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer
    X1 = Cell_X * Cell
    X2 = X1 + Cell
    Y1 = Cell_Y * Cell
    Y2 = Y1 + Cell
    Me.Line (X1, Y1)-(X2, Y2), vbWhite, BF
End Function
'画细胞
Private Function DrawCell(ByVal Cell_X As Integer, ByVal Cell_Y As Integer, cellColor As Long)
    Dim X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
    X1 = Cell_X * Cell
    X2 = X1 + 1 * Cell
    Y1 = Cell_Y * Cell
    Y2 = Y1 + 1 * Cell
    'Debug.Print x1, x2, y1, y2, frame_Width, frame_Height
    '画之前判断是否在frame内
    Me.Line (X1, Y1)-(X2, Y2), cellColor, BF
    Me.Line (X1, Y1)-(X2, Y2), vbBlack, B
End Function
'把新Cubes展示出来
Private Sub DrawNewCubes()
    Call saveCubes
    Call ShowCurrentCubes
    Call ClsOldCubes
    Call DrawCurrentCubes
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Game_State = "running"
        Call Game_Loop
    Else
        Game_State = "stop"
        Call DrawWhiteBackColor
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Game_State = "stop"
End Sub
