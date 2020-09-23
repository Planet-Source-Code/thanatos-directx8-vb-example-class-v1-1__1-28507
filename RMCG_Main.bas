Attribute VB_Name = "RMCG_Main"
Option Explicit

'Key Functions'
' ESC = Exit '
' Middle Mouse Button = On/Off Background Picture'
' Mouse Move = Cube Move
' Mouse Wheel = Fade Cube in/out

Sub main()
'Initializations'
    If Not D3D.InitDirectX8 Then Call SetError(ErrMsg0)
    'If Not D3D.InitD3DManual(640, 480, D3DFMT_X8R8G8B8, 100, D3DFMT_D16, RMCG_Window.hWnd) Then Call SetError(ErrMsg1)
    If Not D3D.InitD3DAutomatic(RMCG_Window.hWnd) Then Call SetError(ErrMsg1A)
    If Not D3D.InitDI(RMCG_Window.hWnd) Then Call SetError(ErrMsg2)
    If Not D3D.InitDS(RMCG_Window.hWnd) Then Call SetError(ErrMsg3)
    If Not D3D.InitRenderState Then Call SetError(ErrMsg7)
    If Not D3D.InitWorldGeometry Then Call SetError(ErrMsg9)
    If Not D3D.SetLight(1, 1, 1, &H202020, 0, 0, -1) Then Call SetError(ErrMsg8)
    If Not D3D.SetMousePointer(32, 32, "mouse.dds") Then Call SetError(ErrMsg5)
    If Not D3D.SetBackgroundPicture("cspace.jpg") Then Call SetError(ErrMsg10)
    If Not D3D.SetFont("Comic Sans MS", 12, False, False) Then Call SetError(ErrMsg6)
    If Not D3D.SetTexture(0, 0, "rock.jpg") Then Call SetError(ErrMsg11)
    If Not D3D.SetTexture(1, 0, "wood.jpg") Then Call SetError(ErrMsg11)
    If Not D3D.SetTexture(2, 0, "nebular.dds") Then Call SetError(ErrMsg11)
'Settings'
    D3D.BackgroundVisible = True
    Call SetMeshes
'Rendering'
    On Error GoTo RError
    Render = True
    While Render
        'Szene clear'
        Call D3D.RefreshData
        'Input Check'
        Call CheckInput
        'Transformations'
        Call TransformMeshes
        'Start Rendering'
        Call D3D.Device.BeginScene
        'Render Scene'
        Call RenderScene
        'End Rendering'
        Call D3D.Device.EndScene
        'Backbuffer flip to Screen'
        Call D3D.Device.Present(ByVal 0, ByVal 0, 0, ByVal 0)
        DoEvents
    Wend
    D3D.Cleanup
RError:
    Call SetError(ErrMsg4)
End Sub
