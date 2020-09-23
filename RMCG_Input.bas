Attribute VB_Name = "RMCG_Input"
Option Explicit

Public Sub CheckInput()
'Keyboard Check'
    If D3D.IsKey(DIK_ESCAPE) Then Render = False
    If D3D.IsMouseB(2) Then
        D3D.BackgroundVisible = Not D3D.BackgroundVisible
        Call D3D.Freeze(500)
    End If
End Sub
