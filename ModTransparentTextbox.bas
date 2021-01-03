Attribute VB_Name = "ModTransparentTextbox"
Option Explicit

'��������
Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
'Download by http://down.liehuo.net
'��������
Private Const GWL_WNDPROC = (-4)
Private Const WM_COMMAND As Long = &H111
Private Const WM_CTLCOLOREDIT As Long = &H133
Private Const WM_DESTROY As Long = &H2
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_HSCROLL As Long = &H114
Private Const WM_VSCROLL As Long = &H115

'API��������
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

  '��������ˢ��(���ˢ����λͼˢ��)
  CreateBGBrush aTxt
  
  '���������û�����໯,��ô�Ϳ�ʼ���໯
  '�ڴ˼�Ҫ˵��GetProp��SetProc���÷�
  'GetProc�ǵõ�һ�����ڵ�����,SetProc������һ�����ڵ�����(û����������ֵǰ,����ֵΪ0)
  If GetProp(GetParent(aTxt.hwnd), "OrigProcAddr") = 0 Then
    SetProp GetParent(aTxt.hwnd), "OrigProcAddr", SetWindowLong(GetParent(aTxt.hwnd), GWL_WNDPROC, AddressOf NewWindowProc)
  End If
  
  '����ı���û�����໯,��ô��ʼ���໯
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
  
  '����ı���ĸ�������ֵ,������Ϊ��λ
  txtWid = aTxtBox.Width / Screen.TwipsPerPixelX
  txtHgt = aTxtBox.Height / Screen.TwipsPerPixelY
  imgLeft = aTxtBox.Left / Screen.TwipsPerPixelX
  imgTop = aTxtBox.Top / Screen.TwipsPerPixelY
  
  screenDC = GetDC(0)                                             'ȡ����ĻDC(DC���豸����)
  picDC = CreateCompatibleDC(screenDC)                            '��������ĻDCһ�µ��ڴ�DC picDC
  picBmp = SelectObject(picDC, aTxtBox.Parent.Picture.Handle)     '�������ڵı���ͼ�ŵ��ڴ�DC

  aTempDC = CreateCompatibleDC(screenDC)                          '������ʱ�ڴ�DC aTempDC
  aTempBmp = CreateCompatibleBitmap(screenDC, txtWid, txtHgt)     '��������ĻDC���ݵ�λͼ(�ܳ���,�������Ϊ�൱�ڴ�
                                                                  '��һ��λͼռλ��,Ϊ������BitBltλͼ��׼��)
  DeleteObject SelectObject(aTempDC, aTempBmp)                    '��λͼռλ��ѡ��aTempDC,��ɾ��aTempDCԭ��
                                                                  '������,ԭ�������ݾ��Ǵ���ɫ����,
  
  '��picDC����λͼ��aTempDC
  BitBlt aTempDC, 0, 0, txtWid, txtHgt, picDC, imgLeft, imgTop, vbSrcCopy
  
  '����ı����Ѿ�������λͼˢ��,��ô����������λͼˢ��
  If GetProp(aTxtBox.hwnd, "CustomBGBrush") <> 0 Then
    DeleteObject GetProp(aTxtBox.hwnd, "CustomBGBrush")
  End If
  SetProp aTxtBox.hwnd, "CustomBGBrush", CreatePatternBrush(aTempBmp)
  
  'ɨβ����
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
  
    If (uMsg = WM_CTLCOLOREDIT) Then     '���ı���ؼ�Ҫdrawʱ��������ڷ������Ϣ,��ʱwParam�����ı���ؼ���DC,
                                         'lParam�����ı���ؼ��ľ��
      isSubclassed = (GetProp(lParam, "OrigProcAddr") <> 0)
      If isSubclassed Then
        CallWindowProc origProc, hwnd, uMsg, wParam, lParam
        SetBkMode wParam, 1                               '�����ı���ؼ��ı���ģʽΪ͸��
        NewWindowProc = GetProp(lParam, "CustomBGBrush")  '�ؼ���һ��,�ı��ı���ؼ��ı���
      Else
        NewWindowProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
      End If
      
    ElseIf uMsg = WM_COMMAND Then     '�ı���ؼ�������ʱ���򸸴��巢��һ��WM_COMMAND��Ϣ
                                      'ӵ�н�������ı����������ַ����ᴥ������Ϣ
      isSubclassed = (GetProp(lParam, "OrigProcAddr") <> 0)
      If isSubclassed Then
        LockWindowUpdate GetParent(lParam)     '����������,��ֹ�����
        InvalidateRect lParam, 0&, 1&          '�����ı��򴰿ڵ�ȫ������
        UpdateWindow lParam                    'ǿ�����������ı��򴰿�
      End If
      NewWindowProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
      If isSubclassed Then LockWindowUpdate 0& '��֮ǰ�����Ĵ��ڽ��н���
      
    ElseIf uMsg = WM_DESTROY Then
    
      SetWindowLong hwnd, GWL_WNDPROC, origProc
      NewWindowProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
      RemoveProp hwnd, "OrigProcAddr"
      
    Else
      NewWindowProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
    End If
  Else
    '��������ⷢ���Ļ�
    NewWindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
  End If
  
End Function

Private Function NewTxtBoxProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
  Dim aRect As RECT
  Dim origProc As Long
  Dim aBrush As Long
  

  origProc = GetProp(hwnd, "OrigProcAddr")
  
  If origProc <> 0 Then
    If uMsg = WM_ERASEBKGND Then         '���ı��򴰿���Ҫ������ʱ��������Ϣ,��ʱwParam���ı����DC
      aBrush = GetProp(hwnd, "CustomBGBrush")
      If aBrush <> 0 Then
        GetClientRect hwnd, aRect
        FillRect wParam, aRect, aBrush
        NewTxtBoxProc = 1                '����ϵͳ�����Ѿ��Լ��ػ����,���������Ļ�ϵͳ���Լ����ػ�һ��������˸
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
