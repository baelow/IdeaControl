# IdeaControl
思维混乱 想法决策工具   

VB6写的 源代码如下    
```
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'
Private Const GWL_STYLE = -16
Private Const WS_SYSMENU = &H80000
Private Const WS_MAXIMIZEBOX = &H10000
Private buttonClickLog As String
Private Const LogFileName As String = "record.txt"
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1


Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    buttonClickLog = ""
    Dim style As Long
    style = GetWindowLong(Me.hwnd, GWL_STYLE)
    style = style And Not WS_MAXIMIZEBOX
    SetWindowLong Me.hwnd, GWL_STYLE, style

End Sub

Private Sub CommandOK_Click()
    'log记录
    Dim currentTime As String
    currentTime = Format(Now, "yyyy-mm-dd hh:mm:ss")
    buttonClickLog = buttonClickLog & "[" & currentTime & "]" & " Can solve"  '& vbCrLf
    WriteToLogFile buttonClickLog
    buttonClickLog = ""
    '格式设置
    CommandOK.BackColor = RGB(0, 201, 167) ' 设置按钮为绿色背景
    CommandOK.Enabled = False ' 禁用按钮
    Timer1.Interval = 30000 ' 设置计时器间隔为30秒
    Timer1.Enabled = True ' 启动计时器
End Sub


Private Sub CommandNo_Click()
    'log设置
    Dim currentTime As String
    currentTime = Format(Now, "yyyy-mm-dd hh:mm:ss")
    buttonClickLog = buttonClickLog & "[" & currentTime & "]" & " Can't solve" '& vbCrLf
    WriteToLogFile buttonClickLog
    '清空日志变量
    buttonClickLog = ""
    '样式设置
    CommandNo.BackColor = RGB(230, 33, 0) ' 设置按钮为红色背景
    CommandNo.Enabled = False ' 禁用按钮
    Timer2.Interval = 10000 ' 设置计时器间隔为10秒
    Timer2.Enabled = True ' 启动计时器
End Sub



Private Sub Timer1_Timer()
    CommandOK.BackColor = vbButtonFace ' 恢复按钮初始背景
    CommandOK.Enabled = True ' 启用按钮
    Timer1.Enabled = False ' 关闭计时器
End Sub

Private Sub Timer2_Timer()
    CommandNo.BackColor = vbButtonFace ' 恢复按钮初始背景
    CommandNo.Enabled = True ' 启用按钮
    Timer2.Enabled = False ' 关闭计时器
End Sub



Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
    End If
End Sub

Private Sub WriteToLogFile(ByVal logData As String)
    Dim FilePath As String
    FilePath = App.Path & "\" & LogFileName
    Dim fileNum As Integer
    fileNum = FreeFile
    If Dir(FilePath) = "" Then
        Open FilePath For Output As fileNum
    Else
        Open FilePath For Append As fileNum
    End If
    If Right(logData, Len(vbCrLf)) <> vbCrLf Then
        logData = logData & vbCrLf
    End If
    Print #fileNum, logData
    Close fileNum
End Sub

Private Sub CommandRecord_Click()
    ' 检查文本框是否隐藏，如果是，则显示出来
    If Text1.Visible = False Then
        Text1.Visible = True
    Else
        SaveAndHideText
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    ' 检查用户是否按下 Enter 键
    If KeyAscii = 13 Then
        SaveAndHideText
    End If
End Sub

Private Sub SaveAndHideText()
    ' 检查文本框是否为空
    If Text1.Text <> "" Then
        Dim FilePath As String
        Dim FreeFileNum As Integer
        FilePath = App.Path & "\IdeaRecord.txt"
        FreeFileNum = FreeFile
        Open FilePath For Append As FreeFileNum
        Print #FreeFileNum, Text1.Text
        Close FreeFileNum
        Text1.Text = ""
    End If
    Text1.Visible = False
End Sub




```
