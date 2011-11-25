VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TTrashWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const CLASS_NAME = "w>gophertrash"

Dim mhWnd As Long
Dim mView As mfxView
Dim mImg As mfxBitmap

Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
End Function

Private Sub Class_Initialize()

    EZRegisterClass CLASS_NAME
    mhWnd = EZ4AddWindow(CLASS_NAME, Me)
    Set mImg = load_image_obj(g_MakePath(App.Path) & "trash.png")

End Sub

Private Sub Class_Terminate()

    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Public Sub Go(ByRef Position As BPoint)

    If (mImg Is Nothing) Or (Position Is Nothing) Or (mhWnd = 0) Then _
        Exit Sub

Static i As Integer

    Set mView = New mfxView
    With mView
        .SizeTo mImg.Height, mImg.Height
        For i = 0 To 4
            .Clear
            .DrawImage mImg, .Bounds, new_BRect(i * mImg.Height, 0, mImg.Height, mImg.Height)
            apply_view_to_window mView, mhWnd, , Position.x - Fix(mImg.Height / 2), Position.y - Fix(mImg.Height / 2)
            g_ShowWindow mhWnd, True, False
            Sleep 50

        Next i

    End With

    g_ShowWindow mhWnd, False

End Sub


