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
Const Cell = 30
Const Action_Speed = 80 '�������Ӧ�ٶ� ��λ��ms Ҳ����˵Action_Speed������������е���һ֡
Const fps = 120 '60֡
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
'��ǰ����
Dim NowCubes_X(3) As Integer
Dim NowCubes_Y(3) As Integer
Dim NowCubes_Mode As CubesMode
Dim NowCubes_Direction As CubesDirection
'��һ������
Dim NextCubes_X(3) As Integer
Dim NextCubes_Y(3) As Integer
Dim NextCubes_Mode As CubesMode '��¼��̬
Dim NextCubes_Direction As CubesDirection
'�ı�ǰ����̬
Dim OldCubes_X(3) As Integer
Dim OldCubes_Y(3) As Integer
Dim OldCubes_Mode As CubesMode
Dim OldCubes_Direction As CubesDirection
'�·���
Dim NewCubes_X(3) As Integer
Dim NewCubes_Y(3) As Integer
Dim NewCubes_Mode As CubesMode
Dim NewCubes_Direction As CubesDirection
'Ӱ�ӷ���
Dim ShadowCubes_X(3) As Integer
Dim ShadowCubes_Y(3) As Integer
Dim ShadowCubes_Mode As CubesDirection
Dim ShadowCubes_Direction As CubesDirection
'���巽��
Const Blocks_MaxIndex = 199 '10 * 20
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
    Me.ScaleMode = 3
    form_Width = (frame_Width + 6) * Cell * TwipsPerPixelX
    form_Height = (frame_Height + 3) * Cell * TwipsPerPixelY
    form_Top = 0
    form_Left = 0
    Me.Move Screen.Width / 3, form_Top, form_Width, form_Height
    Me.ForeColor = vbBlack
    Me.DrawWidth = 2
    '��ʼ����ǰ����
    Call NewRndCubes
    Call ShowNowCubes
    'Ӱ�ӷ���
    Call ShowShadowCubes
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
                Call ShowShadowCubes
            End If
            '������Ӧ
            Game_NewTime = timeGetTime()
            If Game_NewTime - Game_NowTime >= Game_Speed Then
                Game_NowTime = Game_NewTime
                Call saveCubes
                Call switchCubes("down")
                Call ShowShadowCubes
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
    Call DrawNowCubes
    Call DrawShadowCubes
End Sub
'����ж�
Private Function HitButtom(ByVal CubesName As String) As Boolean
    Dim i As Integer
    '�����ж�
    If CubesName = "shadowcubes" Then
        For i = 0 To 3
            If ShadowCubes_Y(i) > frame_Height Then
                HitButtom = True
                Exit Function
            End If
        Next i
    ElseIf CubesName = "nowcubes" Then
        For i = 0 To 3
            If NowCubes_Y(i) > frame_Height Then
                HitButtom = True
                Exit Function
            End If
        Next i
    End If
End Function
'blocks��ײ����
Private Function HitBlocks(ByVal CubesName As String) As Boolean
    Dim i As Integer, j As Integer
    If CubesName = "shadowcubes" Then
            For i = 0 To Blocks_MaxIndex
                If Blocks_Status(i) = 1 Then
                    For j = 0 To 3
                        If ShadowCubes_X(j) = Blocks_X(i) And ShadowCubes_Y(j) = Blocks_Y(i) Then
                            HitBlocks = True
                            Exit Function
                        End If
                    Next
                End If
            Next
    ElseIf CubesName = "nowcubes" Then
            For i = 0 To Blocks_MaxIndex
                If Blocks_Status(i) = 1 Then
                    For j = 0 To 3
                        If NowCubes_X(j) = Blocks_X(i) And NowCubes_Y(j) = Blocks_Y(i) Then
                            HitBlocks = True
                            Exit Function
                        End If
                    Next
                End If
            Next
    End If
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
    Dim c As Integer
    For Y = frame_Height To frame_Top Step -1
        For X = frame_Width To frame_Left Step -1
            i = X + (Y - 1) * frame_Width - 1
            If Blocks_Status(i) = 1 Then
                c = c + 1
            End If
            
        Next
        If c >= 10 Then
            CheckBlocks = Y
            Exit Function
        Else
            '���¼���
            c = 0
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
                If NowCubes_X(j) = Blocks_X(i) And NowCubes_Y(j) = Blocks_Y(i) Then
                    Blocks_Status(i) = 1
                    Blocks_Color(i) = NowCubesColor
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
            NowCubes_X(i) = NowCubes_X(i) - 1
        Next i
        '�ж��Ƿ�����������ǽ��
        If HitFrame <> "" Or HitBlocks("nowcubes") = True Then
            Call BackCubes
            Exit Sub
        End If
    ElseIf moveDirection = "right" Then
        For i = 0 To 3
            NowCubes_X(i) = NowCubes_X(i) + 1
        Next i
        '�ж��Ƿ�����������ǽ��
        If HitFrame <> "" Or HitBlocks("nowcubes") = True Then
            Call BackCubes
            Exit Sub
        End If
    ElseIf moveDirection = "down" Then
        For i = 0 To 3
            NowCubes_Y(i) = NowCubes_Y(i) + 1
        Next i
        '���׻���ʱ����blocks   ֻ�����µ�ʱ����
        If HitButtom("nowcubes") = True Or HitBlocks("nowcubes") = True Then
            Call BackCubes
            Call LockCubes
            Exit Sub
        End If
    ElseIf moveDirection = "rotate" Then
    End If
End Sub
'�������飬������һ��
Private Sub LockCubes()
    Dim Row As Integer
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
End Sub
'�ж��Ƿ��ڿ����
Private Function HitFrame() As String
    Dim i As Integer
    For i = 0 To 3
        If NowCubes_X(i) < frame_Left Then
            HitFrame = "left"
            Exit Function
        ElseIf NowCubes_X(i) > frame_Width Then
            HitFrame = "right"
            Exit Function
        Else
            HitFrame = ""
        End If
    Next
End Function
'����̨
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Game_State = "running" Then
        If KeyCode = vbKeyC Then 'Ӳ��
            Call saveCubes
            Call DownIt
            Call LockCubes 'ֱ������
        ElseIf KeyCode = vbKeySpace Or KeyCode = vbKeyZ Then '��ʱ����ת
            Call saveCubes
            Call rotateCubes '�ı䷽��
        ElseIf KeyCode = vbKeyX Then '˳ʱ����ת
            Call saveCubes
            Call rotateCubes
            Call rotateCubes
            Call rotateCubes
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
    NewCubes_Direction = UpDirection
    '�õ���״
    Randomize
    NewCubes_Mode = Int(Rnd * 7)
    'ȷ���������꼰����
    Randomize
    mDirection = Int(Rnd * 4)
    'ȷ����ʼ����
    Cube0_X = 5
    Cube0_Y = 0
    Select Case NewCubes_Mode
        Case CubeMode
            NewCubes_X(0) = Cube0_X
            NewCubes_Y(0) = Cube0_Y
            NewCubes_X(1) = Cube0_X + 1
            NewCubes_Y(1) = Cube0_Y
            NewCubes_X(2) = Cube0_X
            NewCubes_Y(2) = Cube0_Y + 1
            NewCubes_X(3) = Cube0_X + 1
            NewCubes_Y(3) = Cube0_Y + 1
        Case LineMode
            'ȷ��up�ͷ���
                NewCubes_X(0) = Cube0_X
                NewCubes_Y(0) = Cube0_Y
                NewCubes_X(1) = Cube0_X + 1
                NewCubes_Y(1) = Cube0_Y
                NewCubes_X(2) = Cube0_X + 2
                NewCubes_Y(2) = Cube0_Y
                NewCubes_X(3) = Cube0_X + 3
                NewCubes_Y(3) = Cube0_Y
            If mDirection = UpDirection Or mDirection = DownDirection Then
                '������ת
            ElseIf mDirection = LeftDirection Or mDirection = RightDirection Then
                'һ����ת���任��̬
                Call rotateNewCubes
            Else
                MsgBox "�������������ˣ�", vbCritical, "����"
            End If
        Case LeftZMode
                NewCubes_X(0) = Cube0_X
                NewCubes_Y(0) = Cube0_Y
                NewCubes_X(1) = Cube0_X + 1
                NewCubes_Y(1) = Cube0_Y
                NewCubes_X(2) = Cube0_X + 1
                NewCubes_Y(2) = Cube0_Y + 1
                NewCubes_X(3) = Cube0_X + 2
                NewCubes_Y(3) = Cube0_Y + 1
            If mDirection = UpDirection Or mDirection = DownDirection Then
                'up
            ElseIf mDirection = LeftDirection Or mDirection = RightDirection Then
                Call rotateNewCubes
            Else
                MsgBox "�������������ˣ�", vbCritical, "����"
            End If
        Case RightZMode
                NewCubes_X(0) = Cube0_X
                NewCubes_Y(0) = Cube0_Y
                NewCubes_X(1) = Cube0_X - 1
                NewCubes_Y(1) = Cube0_Y
                NewCubes_X(2) = Cube0_X - 1
                NewCubes_Y(2) = Cube0_Y + 1
                NewCubes_X(3) = Cube0_X - 2
                NewCubes_Y(3) = Cube0_Y + 1
            If mDirection = UpDirection Or mDirection = DownDirection Then
                'up
            ElseIf mDirection = LeftDirection Or mDirection = RightDirection Then
                Call rotateNewCubes
            Else
                MsgBox "�������������ˣ�", vbCritical, "����"
            End If
        Case TMode
                NewCubes_X(0) = Cube0_X
                NewCubes_Y(0) = Cube0_Y
                NewCubes_X(1) = Cube0_X
                NewCubes_Y(1) = Cube0_Y + 1
                NewCubes_X(2) = Cube0_X - 1
                NewCubes_Y(2) = Cube0_Y + 1
                NewCubes_X(3) = Cube0_X + 1
                NewCubes_Y(3) = Cube0_Y + 1
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
                NewCubes_X(0) = Cube0_X
                NewCubes_Y(0) = Cube0_Y
                NewCubes_X(1) = Cube0_X
                NewCubes_Y(1) = Cube0_Y + 1
                NewCubes_X(2) = Cube0_X - 1
                NewCubes_Y(2) = Cube0_Y + 1
                NewCubes_X(3) = Cube0_X - 2
                NewCubes_Y(3) = Cube0_Y + 1
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
                NewCubes_X(0) = Cube0_X
                NewCubes_Y(0) = Cube0_Y
                NewCubes_X(1) = Cube0_X
                NewCubes_Y(1) = Cube0_Y + 1
                NewCubes_X(2) = Cube0_X + 1
                NewCubes_Y(2) = Cube0_Y + 1
                NewCubes_X(3) = Cube0_X + 2
                NewCubes_Y(3) = Cube0_Y + 1
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

'����NowCubes��oldcubes
Private Sub saveCubes()
    Dim i As Integer
    For i = 0 To 3
        OldCubes_X(i) = NowCubes_X(i)
        OldCubes_Y(i) = NowCubes_Y(i)
    Next
    OldCubes_Mode = NowCubes_Mode
    OldCubes_Direction = NowCubes_Direction
End Sub
'���������ת����ǰ
Private Sub ShowNowCubes()
    Dim i As Integer
    For i = 0 To 3
        NowCubes_X(i) = NewCubes_X(i)
        NowCubes_Y(i) = NewCubes_Y(i)
    Next
    NowCubes_Mode = NewCubes_Mode
    NowCubes_Direction = NewCubes_Direction
End Sub
'չʾ�ҵ�Ӱ��
Private Sub ShowShadowCubes()
    Dim i As Integer
    Dim c As Boolean
    c = True
    For i = 0 To 3
        ShadowCubes_X(i) = NowCubes_X(i)
        ShadowCubes_Y(i) = NowCubes_Y(i)
    Next
    ShadowCubes_Mode = NowCubes_Mode
    ShadowCubes_Direction = NowCubes_Direction
    '����
    While c
        For i = 0 To 3
            ShadowCubes_Y(i) = ShadowCubes_Y(i) + 1 '����
        Next
        If HitButtom("shadowcubes") = True Or HitBlocks("shadowcubes") = True Then
            For i = 0 To 3
                ShadowCubes_Y(i) = ShadowCubes_Y(i) - 1
            Next i
            c = False
        End If
    Wend
End Sub
'����
Private Sub BackCubes()
    Dim i As Integer
    For i = 0 To 3
        NowCubes_X(i) = OldCubes_X(i)
        NowCubes_Y(i) = OldCubes_Y(i)
    Next i
    NowCubes_Mode = OldCubes_Mode
    NowCubes_Direction = OldCubes_Direction
End Sub
'չʾ��һ������
Private Sub ShowNextCubes()
    Dim i As Integer
    For i = 0 To 3
        NextCubes_X(i) = NewCubes_X(i) + 7
        NextCubes_Y(i) = NewCubes_Y(i) + 2
    Next
    NextCubes_Mode = NewCubes_Mode
    NextCubes_Direction = NewCubes_Direction
End Sub
'����һ������ŵ���ǰ
Private Sub NextToCurrent()
    Dim i As Integer
    NowCubes_Mode = NextCubes_Mode
    NowCubes_Direction = NextCubes_Direction
    'λ�õĻ���Ҫ�ƶ��������������
    For i = 0 To 3
        NowCubes_X(i) = NextCubes_X(i) - 9
        NowCubes_Y(i) = NextCubes_Y(i)
    Next i
End Sub
'Ӳ��ʵ��
Private Sub DownIt()
    Dim i As Integer
    For i = 0 To 3
        NowCubes_X(i) = ShadowCubes_X(i)
        NowCubes_Y(i) = ShadowCubes_Y(i)
    Next
End Sub
'չʾ��һ������
Private Sub DrawNextCubes()
    Dim i As Integer
    Dim ModeStr As String
    Dim ModeColor As Long
    '���ݲ�ͬ�ķ���ģ��ѡ����ɫ
    Select Case NextCubes_Mode
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
        Call DrawCell(NextCubes_X(i), NextCubes_Y(i), ModeColor)
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
    Call saveCubes
    Select Case NowCubes_Direction
    Case UpDirection
        Select Case NowCubes_Mode
        Case LineMode 'linemode
            NowCubes_Direction = LeftDirection
            NowCubes_X(0) = OldCubes_X(0) + 2
            NowCubes_Y(0) = OldCubes_Y(0) - 2
            NowCubes_X(1) = OldCubes_X(1) + 1
            NowCubes_Y(1) = OldCubes_Y(1) - 1
            NowCubes_X(3) = OldCubes_X(3) - 1
            NowCubes_Y(3) = OldCubes_Y(3) + 1
        Case LeftSevenMode 'LeftSevenMode
            NowCubes_Direction = RightDirection
            NowCubes_X(0) = OldCubes_X(0) - 1
            NowCubes_Y(0) = OldCubes_Y(0) + 1
            NowCubes_X(2) = OldCubes_X(2) + 1
            NowCubes_Y(2) = OldCubes_Y(2) + 1
            NowCubes_X(3) = OldCubes_X(3) + 2
            NowCubes_Y(3) = OldCubes_Y(3) + 2
        Case RightSevenMode 'RightSevenMode
            NowCubes_Direction = RightDirection
            NowCubes_X(0) = OldCubes_X(0) - 1
            NowCubes_Y(0) = OldCubes_Y(0) + 1
            NowCubes_X(2) = OldCubes_X(2) - 1
            NowCubes_Y(2) = OldCubes_Y(2) - 1
            NowCubes_X(3) = OldCubes_X(3) - 2
            NowCubes_Y(3) = OldCubes_Y(3) - 2
        Case LeftZMode 'LeftZMode
            NowCubes_Direction = LeftDirection
            NowCubes_X(0) = OldCubes_X(0) + 1
            NowCubes_Y(0) = OldCubes_Y(0) - 1
            NowCubes_X(2) = OldCubes_X(2) - 1
            NowCubes_Y(2) = OldCubes_Y(2) - 1
            NowCubes_X(3) = OldCubes_X(3) - 2
            NowCubes_Y(3) = OldCubes_Y(3)
        Case RightZMode 'RightZMode
            NowCubes_Direction = LeftDirection
            NowCubes_X(0) = OldCubes_X(0) - 1
            NowCubes_Y(0) = OldCubes_Y(0) - 1
            NowCubes_X(2) = OldCubes_X(2) + 1
            NowCubes_Y(2) = OldCubes_Y(2) - 1
            NowCubes_X(3) = OldCubes_X(3) + 2
            NowCubes_Y(3) = OldCubes_Y(3)
        Case TMode 'T
            NowCubes_Direction = RightDirection
            NowCubes_X(0) = OldCubes_X(0) - 1
            NowCubes_Y(0) = OldCubes_Y(0) + 1
            NowCubes_X(2) = OldCubes_X(2) + 1
            NowCubes_Y(2) = OldCubes_Y(2) + 1
            NowCubes_X(3) = OldCubes_X(3) - 1
            NowCubes_Y(3) = OldCubes_Y(3) - 1
        End Select
    Case DownDirection
        Select Case NowCubes_Mode
            Case LeftSevenMode
                NowCubes_Direction = LeftDirection
                NowCubes_X(0) = OldCubes_X(0) + 1
                NowCubes_Y(0) = OldCubes_Y(0) - 1
                NowCubes_X(2) = OldCubes_X(2) - 1
                NowCubes_Y(2) = OldCubes_Y(2) - 1
                NowCubes_X(3) = OldCubes_X(3) - 2
                NowCubes_Y(3) = OldCubes_Y(3) - 2
            Case RightSevenMode '7��
                NowCubes_Direction = LeftDirection
                NowCubes_X(0) = OldCubes_X(0) + 1
                NowCubes_Y(0) = OldCubes_Y(0) - 1
                NowCubes_X(2) = OldCubes_X(2) + 1
                NowCubes_Y(2) = OldCubes_Y(2) + 1
                NowCubes_X(3) = OldCubes_X(3) + 2
                NowCubes_Y(3) = OldCubes_Y(3) + 2
            Case TMode 'T����
                NowCubes_Direction = LeftDirection
                NowCubes_X(0) = OldCubes_X(0) + 1
                NowCubes_Y(0) = OldCubes_Y(0) - 1
                NowCubes_X(2) = OldCubes_X(2) - 1
                NowCubes_Y(2) = OldCubes_Y(2) - 1
                NowCubes_X(3) = OldCubes_X(3) + 1
                NowCubes_Y(3) = OldCubes_Y(3) + 1
        End Select
    Case LeftDirection
        Select Case NowCubes_Mode
            Case LineMode
                NowCubes_Direction = UpDirection
                NowCubes_X(0) = OldCubes_X(0) - 2
                NowCubes_Y(0) = OldCubes_Y(0) + 2
                NowCubes_X(1) = OldCubes_X(1) - 1
                NowCubes_Y(1) = OldCubes_Y(1) + 1
                NowCubes_X(3) = OldCubes_X(3) + 1
                NowCubes_Y(3) = OldCubes_Y(3) - 1
            Case LeftSevenMode '7��
                NowCubes_Direction = UpDirection
                NowCubes_X(0) = OldCubes_X(0) - 1
                NowCubes_Y(0) = OldCubes_Y(0) - 1
                NowCubes_X(2) = OldCubes_X(2) - 1
                NowCubes_Y(2) = OldCubes_Y(2) + 1
                NowCubes_X(3) = OldCubes_X(3) - 2
                NowCubes_Y(3) = OldCubes_Y(3) + 2
            Case RightSevenMode '7��
                NowCubes_Direction = UpDirection
                NowCubes_X(0) = OldCubes_X(0) - 1
                NowCubes_Y(0) = OldCubes_Y(0) - 1
                NowCubes_X(2) = OldCubes_X(2) + 1
                NowCubes_Y(2) = OldCubes_Y(2) - 1
                NowCubes_X(3) = OldCubes_X(3) + 2
                NowCubes_Y(3) = OldCubes_Y(3) - 2
            Case LeftZMode '��Z��
                NowCubes_Direction = UpDirection
                NowCubes_X(0) = OldCubes_X(0) - 1
                NowCubes_Y(0) = OldCubes_Y(0) + 1
                NowCubes_X(2) = OldCubes_X(2) + 1
                NowCubes_Y(2) = OldCubes_Y(2) + 1
                NowCubes_X(3) = OldCubes_X(3) + 2
                NowCubes_Y(3) = OldCubes_Y(3)
            Case RightZMode
                NowCubes_Direction = UpDirection
                NowCubes_X(0) = OldCubes_X(0) + 1
                NowCubes_Y(0) = OldCubes_Y(0) + 1
                NowCubes_X(2) = OldCubes_X(2) - 1
                NowCubes_Y(2) = OldCubes_Y(2) + 1
                NowCubes_X(3) = OldCubes_X(3) - 2
                NowCubes_Y(3) = OldCubes_Y(3)
            Case TMode 'T����
                NowCubes_Direction = UpDirection
                NowCubes_X(0) = OldCubes_X(0) - 1
                NowCubes_Y(0) = OldCubes_Y(0) - 1
                NowCubes_X(2) = OldCubes_X(2) - 1
                NowCubes_Y(2) = OldCubes_Y(2) + 1
                NowCubes_X(3) = OldCubes_X(3) + 1
                NowCubes_Y(3) = OldCubes_Y(3) - 1
        End Select
    Case RightDirection
        Select Case NowCubes_Mode
            Case LeftSevenMode
                NowCubes_Direction = DownDirection
                NowCubes_X(0) = OldCubes_X(0) + 1
                NowCubes_Y(0) = OldCubes_Y(0) + 1
                NowCubes_X(2) = OldCubes_X(2) + 1
                NowCubes_Y(2) = OldCubes_Y(2) - 1
                NowCubes_X(3) = OldCubes_X(3) + 2
                NowCubes_Y(3) = OldCubes_Y(3) - 2
            Case RightSevenMode '7��
                NowCubes_Direction = DownDirection
                NowCubes_X(0) = OldCubes_X(0) + 1
                NowCubes_Y(0) = OldCubes_Y(0) + 1
                NowCubes_X(2) = OldCubes_X(2) - 1
                NowCubes_Y(2) = OldCubes_Y(2) + 1
                NowCubes_X(3) = OldCubes_X(3) - 2
                NowCubes_Y(3) = OldCubes_Y(3) + 2
        Case TMode 'T����
            NowCubes_Direction = DownDirection
            NowCubes_X(0) = OldCubes_X(0) + 1
            NowCubes_Y(0) = OldCubes_Y(0) + 1
            NowCubes_X(2) = OldCubes_X(2) + 1
            NowCubes_Y(2) = OldCubes_Y(2) - 1
            NowCubes_X(3) = OldCubes_X(3) - 1
            NowCubes_Y(3) = OldCubes_Y(3) + 1
        End Select
    End Select
    '��ת��Ϻ��ж�������Ҫ���ƻ����� '��ǽϵͳ
    Dim i As Integer
    While HitFrame = "left"
        '��������
        For i = 0 To 3
            NowCubes_X(i) = NowCubes_X(i) + 1
        Next i
    Wend
    While HitFrame = "right"
        '��������
        For i = 0 To 3
            NowCubes_X(i) = NowCubes_X(i) - 1
        Next i
    Wend
    '�ж���ǽ���Ƿ�ײ���������ǽʧ��
    If HitBlocks("nowcubes") = True Then
        Call BackCubes
    End If
End Function

'��ת�·���
Private Function rotateNewCubes() As Boolean
    Select Case NewCubes_Direction
    Case UpDirection
        Select Case NewCubes_Mode
        Case LineMode 'linemodeֻ������ up left
            NewCubes_Direction = LeftDirection
            NewCubes_X(0) = NewCubes_X(0) + 2
            NewCubes_Y(0) = NewCubes_Y(0) - 2
            NewCubes_X(1) = NewCubes_X(1) + 1
            NewCubes_Y(1) = NewCubes_Y(1) - 1
            NewCubes_X(3) = NewCubes_X(3) - 1
            NewCubes_Y(3) = NewCubes_Y(3) + 1
        Case LeftSevenMode
            NewCubes_Direction = RightDirection
            NewCubes_X(0) = NewCubes_X(0) - 1
            NewCubes_Y(0) = NewCubes_Y(0) + 1
            NewCubes_X(2) = NewCubes_X(2) + 1
            NewCubes_Y(2) = NewCubes_Y(2) + 1
            NewCubes_X(3) = NewCubes_X(3) + 2
            NewCubes_Y(3) = NewCubes_Y(3) + 2
        Case RightSevenMode '7��
            NewCubes_Direction = RightDirection
            NewCubes_X(0) = NewCubes_X(0) - 1
            NewCubes_Y(0) = NewCubes_Y(0) + 1
            NewCubes_X(2) = NewCubes_X(2) - 1
            NewCubes_Y(2) = NewCubes_Y(2) - 1
            NewCubes_X(3) = NewCubes_X(3) - 2
            NewCubes_Y(3) = NewCubes_Y(3) - 2
        Case LeftZMode
            NewCubes_Direction = LeftDirection
            NewCubes_X(0) = NewCubes_X(0) + 1
            NewCubes_Y(0) = NewCubes_Y(0) - 1
            NewCubes_X(2) = NewCubes_X(2) - 1
            NewCubes_Y(2) = NewCubes_Y(2) - 1
            NewCubes_X(3) = NewCubes_X(3) - 2
            NewCubes_Y(3) = NewCubes_Y(3)
        Case RightZMode
            NewCubes_Direction = LeftDirection
            NewCubes_X(0) = NewCubes_X(0) - 1
            NewCubes_Y(0) = NewCubes_Y(0) - 1
            NewCubes_X(2) = NewCubes_X(2) + 1
            NewCubes_Y(2) = NewCubes_Y(2) - 1
            NewCubes_X(3) = NewCubes_X(3) + 2
            NewCubes_Y(3) = NewCubes_Y(3)
        Case TMode 'T����
            NewCubes_Direction = RightDirection
            NewCubes_X(0) = NewCubes_X(0) - 1
            NewCubes_Y(0) = NewCubes_Y(0) + 1
            NewCubes_X(2) = NewCubes_X(2) + 1
            NewCubes_Y(2) = NewCubes_Y(2) + 1
            NewCubes_X(3) = NewCubes_X(3) - 1
            NewCubes_Y(3) = NewCubes_Y(3) - 1
        End Select
    Case DownDirection
        Select Case NewCubes_Mode
            Case LeftSevenMode
                NewCubes_Direction = LeftDirection
                NewCubes_X(0) = NewCubes_X(0) + 1
                NewCubes_Y(0) = NewCubes_Y(0) - 1
                NewCubes_X(2) = NewCubes_X(2) - 1
                NewCubes_Y(2) = NewCubes_Y(2) - 1
                NewCubes_X(3) = NewCubes_X(3) - 2
                NewCubes_Y(3) = NewCubes_Y(3) - 2
            Case RightSevenMode '7��
                NewCubes_Direction = LeftDirection
                NewCubes_X(0) = NewCubes_X(0) + 1
                NewCubes_Y(0) = NewCubes_Y(0) - 1
                NewCubes_X(2) = NewCubes_X(2) + 1
                NewCubes_Y(2) = NewCubes_Y(2) + 1
                NewCubes_X(3) = NewCubes_X(3) + 2
                NewCubes_Y(3) = NewCubes_Y(3) + 2
            Case TMode 'T����
                NewCubes_Direction = LeftDirection
                NewCubes_X(0) = NewCubes_X(0) + 1
                NewCubes_Y(0) = NewCubes_Y(0) - 1
                NewCubes_X(2) = NewCubes_X(2) - 1
                NewCubes_Y(2) = NewCubes_Y(2) - 1
                NewCubes_X(3) = NewCubes_X(3) + 1
                NewCubes_Y(3) = NewCubes_Y(3) + 1
        End Select
    Case LeftDirection
        Select Case NewCubes_Mode
            Case LineMode
                NewCubes_Direction = UpDirection  '�� �� �� ��
                NewCubes_X(0) = NewCubes_X(0) - 2
                NewCubes_Y(0) = NewCubes_Y(0) + 2
                NewCubes_X(1) = NewCubes_X(1) - 1
                NewCubes_Y(1) = NewCubes_Y(1) + 1
                NewCubes_X(3) = NewCubes_X(3) + 1
                NewCubes_Y(3) = NewCubes_Y(3) - 1
            Case LeftSevenMode '7��
                NewCubes_Direction = UpDirection
                NewCubes_X(0) = NewCubes_X(0) - 1
                NewCubes_Y(0) = NewCubes_Y(0) - 1
                NewCubes_X(2) = NewCubes_X(2) - 1
                NewCubes_Y(2) = NewCubes_Y(2) + 1
                NewCubes_X(3) = NewCubes_X(3) - 2
                NewCubes_Y(3) = NewCubes_Y(3) + 2
            Case RightSevenMode '7��
                NewCubes_Direction = UpDirection
                NewCubes_X(0) = NewCubes_X(0) - 1
                NewCubes_Y(0) = NewCubes_Y(0) - 1
                NewCubes_X(2) = NewCubes_X(2) + 1
                NewCubes_Y(2) = NewCubes_Y(2) - 1
                NewCubes_X(3) = NewCubes_X(3) + 2
                NewCubes_Y(3) = NewCubes_Y(3) - 2
            Case LeftZMode '��Z��
                NewCubes_Direction = UpDirection
                NewCubes_X(0) = NewCubes_X(0) - 1
                NewCubes_Y(0) = NewCubes_Y(0) + 1
                NewCubes_X(2) = NewCubes_X(2) + 1
                NewCubes_Y(2) = NewCubes_Y(2) + 1
                NewCubes_X(3) = NewCubes_X(3) + 2
                NewCubes_Y(3) = NewCubes_Y(3)
            Case RightZMode
                NewCubes_Direction = UpDirection
                NewCubes_X(0) = NewCubes_X(0) + 1
                NewCubes_Y(0) = NewCubes_Y(0) + 1
                NewCubes_X(2) = NewCubes_X(2) - 1
                NewCubes_Y(2) = NewCubes_Y(2) + 1
                NewCubes_X(3) = NewCubes_X(3) - 2
                NewCubes_Y(3) = NewCubes_Y(3)
            Case TMode 'T����
                NewCubes_Direction = UpDirection
                NewCubes_X(0) = NewCubes_X(0) - 1
                NewCubes_Y(0) = NewCubes_Y(0) - 1
                NewCubes_X(2) = NewCubes_X(2) - 1
                NewCubes_Y(2) = NewCubes_Y(2) + 1
                NewCubes_X(3) = NewCubes_X(3) + 1
                NewCubes_Y(3) = NewCubes_Y(3) - 1
        End Select
    Case RightDirection
        Select Case NewCubes_Mode
            Case LeftSevenMode
                NewCubes_Direction = DownDirection
                NewCubes_X(0) = NewCubes_X(0) + 1
                NewCubes_Y(0) = NewCubes_Y(0) + 1
                NewCubes_X(2) = NewCubes_X(2) + 1
                NewCubes_Y(2) = NewCubes_Y(2) - 1
                NewCubes_X(3) = NewCubes_X(3) + 2
                NewCubes_Y(3) = NewCubes_Y(3) - 2
            Case RightSevenMode '7��
                NewCubes_Direction = DownDirection
                NewCubes_X(0) = NewCubes_X(0) + 1
                NewCubes_Y(0) = NewCubes_Y(0) + 1
                NewCubes_X(2) = NewCubes_X(2) - 1
                NewCubes_Y(2) = NewCubes_Y(2) + 1
                NewCubes_X(3) = NewCubes_X(3) - 2
                NewCubes_Y(3) = NewCubes_Y(3) + 2
        Case TMode 'T����
            NewCubes_Direction = DownDirection
            NewCubes_X(0) = NewCubes_X(0) + 1
            NewCubes_Y(0) = NewCubes_Y(0) + 1
            NewCubes_X(2) = NewCubes_X(2) + 1
            NewCubes_Y(2) = NewCubes_Y(2) - 1
            NewCubes_X(3) = NewCubes_X(3) - 1
            NewCubes_Y(3) = NewCubes_Y(3) + 1
        End Select
    End Select
End Function
'������
Private Function DrawNowCubes() As Boolean
    Dim i As Integer
    For i = 0 To 3
        Call DrawCell(NowCubes_X(i), NowCubes_Y(i), NowCubesColor) '0
    Next
End Function
'��Ӱ��
Private Sub DrawShadowCubes()
    Dim X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
    Dim i As Integer
    For i = 0 To 3
        X1 = ShadowCubes_X(i) * Cell
        X2 = X1 + Cell
        Y1 = ShadowCubes_Y(i) * Cell
        Y2 = Y1 + Cell
        Me.Line (X1, Y1)-(X2, Y2), RGB(0, 191, 255), B
    Next
End Sub
'��õ�ǰ�������ɫ
Private Function NowCubesColor() As Long
    Dim ModeColor  As Long
    Select Case NowCubes_Mode
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
    NowCubesColor = ModeColor
End Function
'ɾ������
Private Function ClsOldCubes()
    Dim i As Integer
    For i = 0 To 3
        Call ClsCell(OldCubes_X(i), OldCubes_Y(i))
    Next
End Function
'ɾ����ǰ����
Private Function ClsNowCubes()
    Dim i As Integer
    For i = 0 To 3
        Call ClsCell(NowCubes_X(i), NowCubes_Y(i))
    Next
End Function
'ɾ����ǰ��ClsNextCubes
Private Sub ClsNextCubes()
    Dim i As Integer
    For i = 0 To 3
        Call ClsCell(NextCubes_X(i), NextCubes_Y(i))
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
    Call ShowNowCubes
    Call ClsOldCubes
    Call DrawNowCubes
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
