Attribute VB_Name = "RMCG_Render"
Option Explicit

Public Sub RenderScene()
    If Not D3D.RenderMeshes Then Call SetError(ErrMsg13)
    Call D3D.WriteText(&HFFFFFFFF, 0, 0, 200, 20, "Current Framerate : " & D3D.CurrentFrameRate, DT_VCENTER Or DT_CENTER)
    Call D3D.WriteText(&HFFFFFFFF, D3D.ScreenX - 320, 0, D3D.ScreenX, 20, "X:" & D3D.MouseX & " Y:" & D3D.MouseY & " Z:" & D3D.MouseZ & " B1:" & D3D.IsMouseB(0) & " B2:" & D3D.IsMouseB(1), DT_VCENTER Or DT_CENTER)
End Sub
