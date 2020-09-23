Attribute VB_Name = "RMCG_Transforms"
Option Explicit

Public Sub TransformMeshes()
    Dim x As Integer
'Diamond'
    AngleCube = AngleCube + (60 / D3D.CurrentFrameRate)
    If AngleCube > 360 Then AngleCube = AngleCube - 360
    Call D3D.RotateMeshAroundXAxis(0, AngleCube)
    Call D3D.RotateMeshAroundYAxis(0, AngleCube)
    Call D3D.RotateMeshAroundZAxis(0, AngleCube)
    Call D3D.MoveMeshAlongXAxis(0, D3D.MouseX - (D3D.ScreenX / 2))
    Call D3D.MoveMeshAlongYAxis(0, (D3D.MouseY - (D3D.ScreenY / 2)) * -1)
    Call D3D.MoveMeshAlongZAxis(0, D3D.MouseZ)
'Asteroids'
    For x = 2 To 7
        Call D3D.RotateMeshAroundXAxis(x, -AngleCube)
        Call D3D.RotateMeshAroundYAxis(x, -AngleCube)
        Call D3D.RotateMeshAroundZAxis(x, -AngleCube)
    Next x
    Call D3D.MoveMeshAlongXAxis(2, 260)
    Call D3D.MoveMeshAlongXAxis(3, -260)
    Call D3D.MoveMeshAlongYAxis(4, 260)
    Call D3D.MoveMeshAlongYAxis(5, -260)
    Call D3D.MoveMeshAlongZAxis(6, 260)
    Call D3D.MoveMeshAlongZAxis(7, -260)
    For x = 2 To 7
        Call D3D.RotateMeshAroundXAxis(x, -AngleCube)
        Call D3D.RotateMeshAroundYAxis(x, -AngleCube)
        Call D3D.RotateMeshAroundZAxis(x, -AngleCube)
        Call D3D.MoveMeshAlongXAxis(x, D3D.MouseX - (D3D.ScreenX / 2))
        Call D3D.MoveMeshAlongYAxis(x, (D3D.MouseY - (D3D.ScreenY / 2)) * -1)
        Call D3D.MoveMeshAlongZAxis(x, D3D.MouseZ)
    Next x
'Back Effect'
    AngleBack = AngleBack + (10 / D3D.CurrentFrameRate)
    If AngleBack > 360 Then AngleBack = AngleBack - 360
    Call D3D.RotateMeshAroundZAxis(1, -AngleBack)
End Sub
