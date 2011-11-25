Attribute VB_Name = "mMain"
Option Explicit

Public gWindow As TWindow

Public Sub Main()

    Set gWindow = New TWindow

    With New BMsgLooper
        .Run

    End With

    Set gWindow = Nothing

End Sub
