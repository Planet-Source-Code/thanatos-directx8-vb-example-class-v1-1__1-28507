Attribute VB_Name = "RMCG_Errors"
Option Explicit

'Error Constant'
Public Const ErrMsg0 As String = "Direct 3D 8 : DirectX 8 Initialization Error"
Public Const ErrMsg1 As String = "Direct 3D 8 : Interface Initialization Error"
Public Const ErrMsg1A As String = "Direct 3D 8 : Interface Initialization Error - No available ScreenMode found"
Public Const ErrMsg2 As String = "Direct Input 8 : Interface Initialization Error"
Public Const ErrMsg3 As String = "Direct Sound 8 : Interface Initialization Error"
Public Const ErrMsg4 As String = "Direct 3D 8 : Rendering Error"
Public Const ErrMsg5 As String = "Direct 3D 8 : Mouse Pointer Initialization Error"
Public Const ErrMsg6 As String = "Direct 3D 8 : Font Initialization Error"
Public Const ErrMsg7 As String = "Direct 3D 8 : Render State Initialization Error"
Public Const ErrMsg8 As String = "Direct 3D 8 : Light Initialization Error"
Public Const ErrMsg9 As String = "Direct 3D 8 : World Geometry Initialization Error"
Public Const ErrMsg10 As String = "Direct 3D 8 : Background Initialization Error"
Public Const ErrMsg11 As String = "Direct 3D 8 : Texture Initialization Error"
Public Const ErrMsg12 As String = "Direct 3D 8 : Polygon Initialization Error"
Public Const ErrMsg13 As String = "Direct 3D 8 : Mesh Render Error"

Public Sub SetError(ErrStr As String)
    Call MsgBox(ErrStr)
    D3D.Cleanup
End Sub


