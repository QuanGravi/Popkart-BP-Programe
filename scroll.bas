Attribute VB_Name = "scroll"
'������1�͹�����2
Option Explicit
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = -4&

Public Const WM_MOUSEWHEEL = &H20A

Public OldWindowProc As Long '��������ϵͳĬ�ϵĴ�����Ϣ�������ĵ�ַ
Public OldWindowProc2 As Long
Public OldWindowProc_x2 As Long
Public OldWindowProc_y1 As Long

Public hwndVS As Long '��������ؼ��ľ��
Public hwndVS2 As Long
Public hwndVS_x2 As Long
Public hwndVS_y1 As Long


'������1
'�Զ������Ϣ������
Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
If Msg = WM_MOUSEWHEEL Then
'����������¼����д���
If wParam = -7864320 Then '���¹���
   Form01.VScroll1.Value = Form01.VScroll1.Value + 1
ElseIf wParam = 7864320 Then '���Ϲ���
   Form01.VScroll1.Value = Form01.VScroll1.Value - 1
End If
Else
'����Ĭ�ϴ�����Ϣ������
NewWindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
End If
End Function

'������2
Public Function NewWindowProc2(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
If Msg = WM_MOUSEWHEEL Then
If wParam = -7864320 Then '���¹���
   Form01.VScroll2.Value = Form01.VScroll2.Value + 1
ElseIf wParam = 7864320 Then '���Ϲ���
   Form01.VScroll2.Value = Form01.VScroll2.Value - 1
End If
Else
NewWindowProc2 = CallWindowProc(OldWindowProc2, hwnd, Msg, wParam, lParam)
End If
End Function

Public Function NewWindowProc_x2(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
If Msg = WM_MOUSEWHEEL Then
If wParam = -7864320 Then '���¹���
   Form05.VScroll1.Value = Form05.VScroll1.Value + 1
ElseIf wParam = 7864320 Then '���Ϲ���
   Form05.VScroll1.Value = Form05.VScroll1.Value - 1
End If
Else
NewWindowProc_x2 = CallWindowProc(OldWindowProc_x2, hwnd, Msg, wParam, lParam)
End If
End Function

'map pool creation���������3
Public Function NewWindowProc_y1(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
If Msg = WM_MOUSEWHEEL Then
If wParam = -7864320 Then '���¹���
   Form03.VScroll3.Value = Form03.VScroll3.Value + 1
ElseIf wParam = 7864320 Then '���Ϲ���
   Form03.VScroll3.Value = Form03.VScroll3.Value - 1
End If
Else
NewWindowProc_y1 = CallWindowProc(OldWindowProc_y1, hwnd, Msg, wParam, lParam)
End If
End Function
































