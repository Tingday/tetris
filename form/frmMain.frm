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
'����Ϸʹ��ö�ٷ������ϣ��ʹ�þ���任�����з�����ת�����Բο���������
'��tetris��Ϸ�ؼ�����̽�֡� ������
Option Explicit
'API
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
'����
Const Cell = 30 'me.scaleMode = 3
Const Action_Speed = 80 '�������Ӧ�ٶ� ��λ��ms Ҳ����˵Action_Speed������������е���һ֡
Const fps = 60 '60֡
Const Game_Speed = 500
'����
Dim form_Width As Integer
Dim form_Height As Integer
Dim form_Top As Integer
Dim form_Left As Integer
'���frame
Dim frame_Width As Integer
Dim frame_Height As Integer
Dim frame_Top As Integer
Dim frame_Left As Integer
Dim TwipsPerPixelX As Long '���غ��ת������ ��ͬ��ʾ����һ��
Dim TwipsPerPixelY As Long
'�Ʒ���
Dim Score As Long '����������
Dim User_Action As String 'left right down change
Dim Game_State As String ' running /pause /stop /dead
'��¼Ϊλ������
Dim currentCubes_X(3) As Integer
Dim currentCubes_Y(3) As Integer
Dim currentCubes_Mode As CubesMode
Dim currentCubes_Direction As CubesDirection
'��һ������
Dim nextCubes_X(3) As Integer
Dim nextCubes_Y(3) As Integer
Dim nextCubes_Mode As CubesMode '��¼��̬
Dim nextCubes_Direction As CubesDirection
'�ı�ǰ����̬
Dim oldCubes_X(3) As Integer
Dim oldCubes_Y(3) As Integer
Dim oldCubes_Mode As CubesMode
Dim oldCubes_Direction As CubesDirection
'��cubes
Dim newCubes_X(3) As Integer
Dim newCubes_Y(3) As Integer
Dim newCubes_Mode As CubesMode
Dim newCubes_Direction As CubesDirection
'blocks
Const Blocks_MaxIndex = 199 ' ��0��ʼ
Dim Blocks_X(Blocks_MaxIndex) As Integer
Dim Blocks_Y(Blocks_MaxIndex) As Integer
Dim Blocks_Status(Blocks_MaxIndex)  As Integer
Dim Blocks_Color(Blocks_MaxIndex) As Long
'��̬ö��
Private Enum CubesMode
    LineMode = 1
    CubeMode = 0
    LeftSevenMode = 2
    RightSevenMode = 3
    TMode = 4
    LeftZMode = 5
    RightZMode = 6
End Enum
'���鷽��ö��
Private Enum CubesDirection
    UpDirection = 0
    DownDirection = 1
    RightDirection = 2
    LeftDirection = 3
End Enum

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    User_Action = "hold"
End Sub

'��ʼ����
Private Sub Form_Load()
    '��õ�ǰ��ʾ������
    TwipsPerPixelX = Screen.TwipsPerPixelX
    TwipsPerPixelY = Screen.TwipsPerPixelY
    'Debug.Print TwipsPerPixelX, TwipsPerPixelY '��Ȼ����Ϊ15
    'game frame'15/24 �� 1/0.618 Ϊ�ƽ�ָ����
    frame_Top = 1
    frame_Left = 1
    frame_Width = 10
    frame_Height = 20
    '��ʼ������
    form_Width = (frame_Width + 6) * Cell * TwipsPerPixelX
    form_Height = (frame_Height + 3) * Cell * TwipsPerPixelY
    form_Top = 0
    form_Left = 0
    Me.Move Screen.Width / 3, form_Top, form_Width, form_Height
    Me.ForeColor = vbBlack
    Me.DrawWidth = 2
    '��ʼ����ǰ����
    Call NewRndCubes
    Call ShowCurrentCubes
    '����һ������
    Call NewRndCubes
    Call ShowNextCubes
    '����Blocks
    Call ResetBlocks
    '��
    Call ReDrawUI
End Sub
'��Ϸѭ��
Private Sub Game_Loop()
    Dim Game_NowTime As Long
    Dim Game_NewTime As Long
    Dim Action_NowTime As Long
    Dim Action_NewTime As Long
    Dim Draw_NowTime As Long
    Dim Draw_NewTime As Long
    While DoEvents
        If Game_State = "running" Then
            '������Ӧ
            Action_NewTime = timeGetTime()
            If Action_NewTime - Action_NowTime >= Action_Speed Then
                Action_NowTime = Action_NewTime
                Call saveCubes
                Call switchCubes(User_Action)
            End If
            '������Ӧ
            Game_NewTime = timeGetTime()
            If Game_NewTime - Game_NowTime >= Game_Speed Then
                Game_NowTime = Game_NewTime
                Call saveCubes
                Call switchCubes("down")
            End If
            '����ˢ��
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
'�ػ�����
Private Sub ReDrawUI()
    'Me.Cls
    Call DrawWhiteBackColor
    Call DrawWall
    Call DrawBlocks
    Call DrawNextCubes
    Call DrawCurrentCubes
End Sub
'����ж�
Private Function HitButtom() As Boolean
    Dim i As Integer
    '�����ж�
    For i = 0 To 3
        If currentCubes_Y(i) > frame_Height Then
            HitButtom = True
            Exit Function
        End If
    Next i
End Function
'��ײBlocks,'������ײ��blocks��id
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
'����Blocks
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
'�����ж� ���ؿ�������ROW
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
            '���¼���
            C = 0
        End If
    Next
End Function
'���ƺ��������Ϸ��ķ���״̬���Ƶ��·�
Private Sub MoveBlocksStatus(ByVal Row As Integer)
    Dim X As Integer, Y As Integer
    Dim i As Integer, j As Integer
    '�ӵ�ROW�п�ʼ���帴������
    For Y = Row To frame_Top Step -1
        For X = frame_Width To frame_Left Step -1
            If Y > 1 Then
                i = X + (Y - 1) * frame_Width - 1
                j = X + (Y - 2) * frame_Width - 1 '��һ��
                Blocks_Status(i) = 0 '��յ�ǰ��
                Blocks_Status(i) = Blocks_Status(j)
                Blocks_Color(i) = Blocks_Color(j) '����ɫҲ������
                Blocks_Status(j) = 0 '�����һ��
            Else
                Blocks_Status(i) = 0 '�����ֱ������
            End If
        Next X
    Next Y
    '����1��
    Score = Score + 1
End Sub
'���Ƶ�ǰ���鵽Blocks��
'�������������жϣ�
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
'��������
Private Sub switchCubes(ByVal moveDirection As String)
    Dim i As Integer
    '��ʼ�Ա���
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
    '���׻���ʱ����blocks
    If HitButtom = True Or HitBlocks = True Then
        Dim Row As Integer
        Call BackCubes
        Call CopyToBlocks
        '�����µķ���
        Call NextToCurrent
        '�����µ���һλ����
        Call ClsNextCubes
        Call NewRndCubes
        Call ShowNextCubes
        Call DrawNextCubes
        '�ظ����ֱ��û��
        While CheckBlocks > 0
            Row = CheckBlocks
            Call MoveBlocksStatus(Row)
            '�����������ô�Ϳ������»���
        Wend
        Exit Sub
    End If
    '�ж��Ƿ����
    If CubeInFrame = False Then
        Call BackCubes
        Exit Sub
    End If
End Sub
'�ж��Ƿ��ڿ����
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
'����̨
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Game_State = "running" Then
        If KeyCode = vbKeySpace Then
            Call saveCubes
            Call rotateCubes '�ı䷽��
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
'�������Cubes
Private Sub NewRndCubes()
    Dim i As Integer
    Dim Cube0_X As Integer, Cube0_Y As Integer
    Dim mDirection As CubesDirection
    '��ʼ����
    newCubes_Direction = UpDirection
    '�õ���״
    Randomize
    newCubes_Mode = Int(Rnd * 7)
    'ȷ���������꼰����
    Randomize
    mDirection = Int(Rnd * 4)
    'ȷ����ʼ����
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
            'ȷ��up�ͷ���
                newCubes_X(0) = Cube0_X
                newCubes_Y(0) = Cube0_Y
                newCubes_X(1) = Cube0_X + 1
                newCubes_Y(1) = Cube0_Y
                newCubes_X(2) = Cube0_X + 2
                newCubes_Y(2) = Cube0_Y
                newCubes_X(3) = Cube0_X + 3
                newCubes_Y(3) = Cube0_Y
            If mDirection = UpDirection Or mDirection = DownDirection Then
                '������ת
            ElseIf mDirection = LeftDirection Or mDirection = RightDirection Then
                'һ����ת���任��̬
                Call rotateNewCubes
            Else
                MsgBox "�������������ˣ�", vbCritical, "����"
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
                MsgBox "�������������ˣ�", vbCritical, "����"
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
                MsgBox "�������������ˣ�", vbCritical, "����"
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
                    Call rotateNewCubes 'down ��up �������α任�õ�
                ElseIf mDirection = LeftDirection Then
                    Call rotateNewCubes
                    Call rotateNewCubes
                    Call rotateNewCubes 'left���α任�õ�
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
                    Call rotateNewCubes 'down ��up �������α任�õ�
                ElseIf mDirection = LeftDirection Then
                    Call rotateNewCubes
                    Call rotateNewCubes
                    Call rotateNewCubes 'left���α任�õ�
                End If
    End Select
End Sub

'����currentCubes��oldcubes
Private Sub saveCubes()
    Dim i As Integer
    For i = 0 To 3
        oldCubes_X(i) = currentCubes_X(i)
        oldCubes_Y(i) = currentCubes_Y(i)
    Next
    oldCubes_Mode = currentCubes_Mode
    oldCubes_Direction = currentCubes_Direction
End Sub
'���������ת����ǰ
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
'չʾ��һ������
Private Sub ShowNextCubes()
    Dim i As Integer
    For i = 0 To 3
        nextCubes_X(i) = newCubes_X(i) + 7
        nextCubes_Y(i) = newCubes_Y(i) + 2
    Next
    nextCubes_Mode = newCubes_Mode
    nextCubes_Direction = newCubes_Direction
End Sub
'����һ������ŵ���ǰ
Private Sub NextToCurrent()
    Dim X As Integer, Y As Integer
    Dim i As Integer
    currentCubes_Mode = nextCubes_Mode
    currentCubes_Direction = nextCubes_Direction
    'λ�õĻ���Ҫ�ƶ��������������
    For i = 0 To 3
        currentCubes_X(i) = nextCubes_X(i) - 9
        currentCubes_Y(i) = nextCubes_Y(i)
    Next i
End Sub
'չʾ��һ������
Private Sub DrawNextCubes()
    Dim i As Integer
    Dim ModeStr As String
    Dim ModeColor As Long
    '���ݲ�ͬ�ķ���ģ��ѡ����ɫ
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
'������ǰ��Blocks�����ӻ�������ô��
Private Sub DrawBlocks()
    Dim i As Integer
    For i = 0 To Blocks_MaxIndex
        If Blocks_Status(i) = 1 Then  '1��������
            Call DrawCell(Blocks_X(i), Blocks_Y(i), Blocks_Color(i))
        End If
    Next
End Sub
'����ɫ����
Private Sub DrawWhiteBackColor()
    Me.Line (form_Left, form_Top)-(form_Width, form_Height), vbWhite, BF
End Sub
'���߿�����
Private Sub DrawWall()
    Dim X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer
    X1 = frame_Left * Cell
    Y1 = frame_Top * Cell
    '��Ϊ��Ҫ����ǽ��������Ҫ��1
    X2 = frame_Width * Cell + Cell
    Y2 = frame_Height * Cell + Cell
    'ǽ��
    Me.Line (X1, Y1)-(X2, Y2), vbBlack, B
End Sub
'��ת����
Private Function rotateCubes() As Boolean
    Select Case currentCubes_Direction
    Case UpDirection
        Select Case currentCubes_Mode
        Case LineMode 'linemodeֻ������ up left
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
        Case RightSevenMode '7��
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
        Case TMode 'T����
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
            Case RightSevenMode '7��
                currentCubes_Direction = LeftDirection
                currentCubes_X(0) = oldCubes_X(0) + 1
                currentCubes_Y(0) = oldCubes_Y(0) - 1
                currentCubes_X(2) = oldCubes_X(2) + 1
                currentCubes_Y(2) = oldCubes_Y(2) + 1
                currentCubes_X(3) = oldCubes_X(3) + 2
                currentCubes_Y(3) = oldCubes_Y(3) + 2
            Case TMode 'T����
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
                currentCubes_Direction = UpDirection  '�� �� �� ��
                currentCubes_X(0) = oldCubes_X(0) - 2
                currentCubes_Y(0) = oldCubes_Y(0) + 2
                currentCubes_X(1) = oldCubes_X(1) - 1
                currentCubes_Y(1) = oldCubes_Y(1) + 1
                currentCubes_X(3) = oldCubes_X(3) + 1
                currentCubes_Y(3) = oldCubes_Y(3) - 1
            Case LeftSevenMode '7��
                currentCubes_Direction = UpDirection
                currentCubes_X(0) = oldCubes_X(0) - 1
                currentCubes_Y(0) = oldCubes_Y(0) - 1
                currentCubes_X(2) = oldCubes_X(2) - 1
                currentCubes_Y(2) = oldCubes_Y(2) + 1
                currentCubes_X(3) = oldCubes_X(3) - 2
                currentCubes_Y(3) = oldCubes_Y(3) + 2
            Case RightSevenMode '7��
                currentCubes_Direction = UpDirection
                currentCubes_X(0) = oldCubes_X(0) - 1
                currentCubes_Y(0) = oldCubes_Y(0) - 1
                currentCubes_X(2) = oldCubes_X(2) + 1
                currentCubes_Y(2) = oldCubes_Y(2) - 1
                currentCubes_X(3) = oldCubes_X(3) + 2
                currentCubes_Y(3) = oldCubes_Y(3) - 2
            Case LeftZMode '��Z��
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
            Case TMode 'T����
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
            Case RightSevenMode '7��
                currentCubes_Direction = DownDirection
                currentCubes_X(0) = oldCubes_X(0) + 1
                currentCubes_Y(0) = oldCubes_Y(0) + 1
                currentCubes_X(2) = oldCubes_X(2) - 1
                currentCubes_Y(2) = oldCubes_Y(2) + 1
                currentCubes_X(3) = oldCubes_X(3) - 2
                currentCubes_Y(3) = oldCubes_Y(3) + 2
        Case TMode 'T����
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

'��ת����
Private Function rotateNewCubes() As Boolean
    Select Case newCubes_Direction
    Case UpDirection
        Select Case newCubes_Mode
        Case LineMode 'linemodeֻ������ up left
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
        Case RightSevenMode '7��
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
        Case TMode 'T����
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
            Case RightSevenMode '7��
                newCubes_Direction = LeftDirection
                newCubes_X(0) = newCubes_X(0) + 1
                newCubes_Y(0) = newCubes_Y(0) - 1
                newCubes_X(2) = newCubes_X(2) + 1
                newCubes_Y(2) = newCubes_Y(2) + 1
                newCubes_X(3) = newCubes_X(3) + 2
                newCubes_Y(3) = newCubes_Y(3) + 2
            Case TMode 'T����
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
                newCubes_Direction = UpDirection  '�� �� �� ��
                newCubes_X(0) = newCubes_X(0) - 2
                newCubes_Y(0) = newCubes_Y(0) + 2
                newCubes_X(1) = newCubes_X(1) - 1
                newCubes_Y(1) = newCubes_Y(1) + 1
                newCubes_X(3) = newCubes_X(3) + 1
                newCubes_Y(3) = newCubes_Y(3) - 1
            Case LeftSevenMode '7��
                newCubes_Direction = UpDirection
                newCubes_X(0) = newCubes_X(0) - 1
                newCubes_Y(0) = newCubes_Y(0) - 1
                newCubes_X(2) = newCubes_X(2) - 1
                newCubes_Y(2) = newCubes_Y(2) + 1
                newCubes_X(3) = newCubes_X(3) - 2
                newCubes_Y(3) = newCubes_Y(3) + 2
            Case RightSevenMode '7��
                newCubes_Direction = UpDirection
                newCubes_X(0) = newCubes_X(0) - 1
                newCubes_Y(0) = newCubes_Y(0) - 1
                newCubes_X(2) = newCubes_X(2) + 1
                newCubes_Y(2) = newCubes_Y(2) - 1
                newCubes_X(3) = newCubes_X(3) + 2
                newCubes_Y(3) = newCubes_Y(3) - 2
            Case LeftZMode '��Z��
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
            Case TMode 'T����
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
            Case RightSevenMode '7��
                newCubes_Direction = DownDirection
                newCubes_X(0) = newCubes_X(0) + 1
                newCubes_Y(0) = newCubes_Y(0) + 1
                newCubes_X(2) = newCubes_X(2) - 1
                newCubes_Y(2) = newCubes_Y(2) + 1
                newCubes_X(3) = newCubes_X(3) - 2
                newCubes_Y(3) = newCubes_Y(3) + 2
        Case TMode 'T����
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
'������
Private Function DrawCurrentCubes() As Boolean
    Dim i As Integer
    Dim ModeStr As String
    For i = 0 To 3
        Call DrawCell(currentCubes_X(i), currentCubes_Y(i), CurrentCubesColor) '0
    Next
    'Debug.Print "Mode:Direction", currentCubes_Mode, currentCubes_Direction
End Function
'��õ�ǰ�������ɫ
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
'ɾ������
Private Function ClsOldCubes()
    Dim i As Integer
    For i = 0 To 3
        Call ClsCell(oldCubes_X(i), oldCubes_Y(i))
    Next
End Function
'ɾ����ǰ����
Private Function ClsCurrentCubes()
    Dim i As Integer
    For i = 0 To 3
        Call ClsCell(currentCubes_X(i), currentCubes_Y(i))
    Next
End Function
'ɾ����ǰ��ClsNextCubes
Private Sub ClsNextCubes()
    Dim i As Integer
    For i = 0 To 3
        Call ClsCell(nextCubes_X(i), nextCubes_Y(i))
    Next
End Sub
'ɾ��ϸ��
Private Function ClsCell(ByVal Cell_X As Integer, ByVal Cell_Y As Integer)
    Dim X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer
    X1 = Cell_X * Cell
    X2 = X1 + Cell
    Y1 = Cell_Y * Cell
    Y2 = Y1 + Cell
    Me.Line (X1, Y1)-(X2, Y2), vbWhite, BF
End Function
'��ϸ��
Private Function DrawCell(ByVal Cell_X As Integer, ByVal Cell_Y As Integer, cellColor As Long)
    Dim X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
    X1 = Cell_X * Cell
    X2 = X1 + 1 * Cell
    Y1 = Cell_Y * Cell
    Y2 = Y1 + 1 * Cell
    'Debug.Print x1, x2, y1, y2, frame_Width, frame_Height
    '��֮ǰ�ж��Ƿ���frame��
    Me.Line (X1, Y1)-(X2, Y2), cellColor, BF
    Me.Line (X1, Y1)-(X2, Y2), vbBlack, B
End Function
'����Cubesչʾ����
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
