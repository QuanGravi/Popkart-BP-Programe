Attribute VB_Name = "ModTransparentTextbox"
Option Explicit

'类型声明
Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
'Download by http://down.liehuo.net
'常数声明
Private Const GWL_WNDPROC = (-4)
Private Const WM_COMMAND As Long = &H111
Private Const WM_CTLCOLOREDIT As Long = &H133
Private Const WM_DESTROY As Long = &H2
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_HSCROLL As Long = &H114
Private Const WM_VSCROLL As Long = &H115

'API函数声明
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long

Public Function MakeTransparentTextbox(aTxt As TextBox)

  '建立背景刷子(这个刷子是位图刷子)
  CreateBGBrush aTxt
  
  '如果主窗口没有子类化,那么就开始子类化
  '在此简要说明GetProp和SetProc的用法
  'GetProc是得到一个窗口的属性,SetProc是设置一个窗口的属性(没有设置属性值前,属性值为0)
  If GetProp(GetParent(aTxt.hwnd), "OrigProcAddr") = 0 Then
    SetProp GetParent(aTxt.hwnd), "OrigProcAddr", SetWindowLong(GetParent(aTxt.hwnd), GWL_WNDPROC, AddressOf NewWindowProc)
  End If
  
  '如果文本框没有子类化,那么开始子类化
  If GetProp(aTxt.hwnd, "OrigProcAddr") = 0 Then
    SetProp aTxt.hwnd, "OrigProcAddr", SetWindowLong(aTxt.hwnd, GWL_WNDPROC, AddressOf NewTxtBoxProc)
  End If
  
End Function

Private Sub CreateBGBrush(aTxtBox As TextBox)

  Dim screenDC As Long
  Dim imgLeft As Long
  Dim imgTop As Long
  Dim picDC As Long
  Dim picBmp As Long
  Dim aTempBmp As Long
  Dim aTempDC As Long
  Dim txtWid As Long
  Dim txtHgt As Long
  Dim SolidBrush As Long
  Dim aRect As RECT
  
  If aTxtBox.Parent.Picture Is Nothing Then Exit Sub
  
  '获得文本框的各种属性值,以像素为单位
  txtWid = aTxtBox.Width / Screen.TwipsPerPixelX
  txtHgt = aTxtBox.Height / Screen.TwipsPerPixelY
  imgLeft = aTxtBox.Left / Screen.TwipsPerPixelX
  imgTop = aTxtBox.Top / Screen.TwipsPerPixelY
  
  screenDC = GetDC(0)                                             '取得屏幕DC(DC即设备场景)
  picDC = CreateCompatibleDC(screenDC)                            '创建和屏幕DC一致的内存DC picDC
  picBmp = SelectObject(picDC, aTxtBox.Parent.Picture.Handle)     '将主窗口的背景图放到内存DC

  aTempDC = CreateCompatibleDC(screenDC)                          '创建临时内存DC aTempDC
  aTempBmp = CreateCompatibleBitmap(screenDC, txtWid, txtHgt)     '创建与屏幕DC兼容的位图(很抽象,个人理解为相当于创
                                                                  '建一个位图占位符,为接下来BitBlt位图做准备)
  DeleteObject SelectObject(aTempDC, aTempBmp)                    '将位图占位符选入aTempDC,并删除aTempDC原来
                                                                  '的内容,原来的内容就是纯黑色背景,
  
  '从picDC复制位图到aTempDC
  BitBlt aTempDC, 0, 0, txtWid, txtHgt, picDC, imgLeft, imgTop, vbSrcCopy
  
  '如果文本框已经设置了位图刷子,那么就重新设置位图刷子
  If GetProp(aTxtBox.hwnd, "CustomBGBrush") <> 0 Then
    DeleteObject GetProp(aTxtBox.hwnd, "CustomBGBrush")
  End If
  SetProp aTxtBox.hwnd, "CustomBGBrush", CreatePatternBrush(aTempBmp)
  
  '扫尾工作
  DeleteDC aTempDC
  DeleteObject aTempBmp
  SelectObject picDC, picBmp
  DeleteDC picDC
  DeleteObject picBmp
  ReleaseDC 0, screenDC
  
End Sub

Private Function NewWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  Dim origProc As Long
  Dim isSubclassed As Long
  
  origProc = GetProp(hwnd, "OrigProcAddr")
  
  If origProc <> 0 Then
  
    If (uMsg = WM_CTLCOLOREDIT) Then     '当文本框控件要draw时会给父窗口发这个消息,此时wParam就是文本框控件的DC,
                                         'lParam就是文本框控件的句柄
      isSubclassed = (GetProp(lParam, "OrigProcAddr") <> 0)
      If isSubclassed Then
        CallWindowProc origProc, hwnd, uMsg, wParam, lParam
        SetBkMode wParam, 1                               '设置文本框控件的背景模式为透明
        NewWindowProc = GetProp(lParam, "CustomBGBrush")  '关键的一步,改变文本框控件的背景
      Else
        NewWindowProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
      End If
      
    ElseIf uMsg = WM_COMMAND Then     '文本框控件被触发时，向父窗体发送一个WM_COMMAND消息
                                      '拥有焦点和在文本框内输入字符都会触发该消息
      isSubclassed = (GetProp(lParam, "OrigProcAddr") <> 0)
      If isSubclassed Then
        LockWindowUpdate GetParent(lParam)     '锁定主窗口,禁止其更新
        InvalidateRect lParam, 0&, 1&          '屏蔽文本框窗口的全部区域
        UpdateWindow lParam                    '强制立即更新文本框窗口
      End If
      NewWindowProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
      If isSubclassed Then LockWindowUpdate 0& '对之前上锁的窗口进行解锁
      
    ElseIf uMsg = WM_DESTROY Then
    
      SetWindowLong hwnd, GWL_WNDPROC, origProc
      NewWindowProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
      RemoveProp hwnd, "OrigProcAddr"
      
    Else
      NewWindowProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
    End If
  Else
    '如果有意外发生的话
    NewWindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
  End If
  
End Function

Private Function NewTxtBoxProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
  Dim aRect As RECT
  Dim origProc As Long
  Dim aBrush As Long
  

  origProc = GetProp(hwnd, "OrigProcAddr")
  
  If origProc <> 0 Then
    If uMsg = WM_ERASEBKGND Then         '当文本框窗口需要被擦除时触发该消息,此时wParam是文本框的DC
      aBrush = GetProp(hwnd, "CustomBGBrush")
      If aBrush <> 0 Then
        GetClientRect hwnd, aRect
        FillRect wParam, aRect, aBrush
        NewTxtBoxProc = 1                '告诉系统我们已经自己重绘过了,不这样做的话系统会自己再重绘一次引起闪烁
      Else
        NewTxtBoxProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
      End If
      
    ElseIf uMsg = WM_HSCROLL Or uMsg = WM_VSCROLL Then

      LockWindowUpdate GetParent(hwnd)
      NewTxtBoxProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
      InvalidateRect hwnd, 0&, 1&
      UpdateWindow hwnd
      LockWindowUpdate 0&
      
    ElseIf uMsg = WM_DESTROY Then
    
      aBrush = GetProp(hwnd, "CustomBGBrush")
      DeleteObject aBrush
      RemoveProp hwnd, "OrigProcAddr"
      RemoveProp hwnd, "CustomBGBrush"
      SetWindowLong hwnd, GWL_WNDPROC, origProc
      NewTxtBoxProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
      
    Else
      NewTxtBoxProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
    End If
  Else
    NewTxtBoxProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
  End If
  
End Function
