Attribute VB_Name = "mMain"
Option Explicit

Public Const CLASS_NAME = "w>gopherlaunch"

Public gWindow As TWindow

Public Sub Main()
Dim hWndExisting As Long

    hWndExisting = FindWindow(CLASS_NAME, CLASS_NAME)

    If Command$ = "-quit" Then
        If IsWindow(hWndExisting) <> 0 Then _
            SendMessage hWndExisting, WM_CLOSE, 0, ByVal 0&

        Exit Sub

    ElseIf IsWindow(hWndExisting) <> 0 Then
        Exit Sub

    End If

    Set gWindow = New TWindow

    With New BMsgLooper
        .Run

    End With

    Set gWindow = Nothing

End Sub
