VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DX8_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'W32 APIS'
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Shader'
Private Const CFVF = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1

'Math'
Private Const PI As Single = 3.14159265358979
Private Const Radiant As Single = PI / 180

'Vertex Format'
Private Type CVERTEX
    x As Single
    y As Single
    z As Single
    nx As Single
    ny As Single
    nz As Single
    tu As Single
    tv As Single
End Type

Private Type D3DMESH
    Vertices()   As CVERTEX
    Indices()    As Integer
    Vbuffer      As Direct3DVertexBuffer8
    IBuffer      As Direct3DIndexBuffer8
    Texture      As Integer
    PolyNum      As Integer
    IndexNum     As Integer
    VertexNum    As Integer
    Draw         As Boolean
End Type

Private Type D3DMOUSEMODE
    x       As Long
    y       As Long
    z       As Long
    b(2)    As Boolean
End Type

Private Type D3DFRAMES
    Last    As Long
    Count   As Long
    Current As Long
End Type

Private Type D3DFONT
    New     As D3DXFont
    Desc    As IFont
    Set     As New StdFont
End Type

'Woldcoords'
Private matWorld As D3DMATRIX
Private matView  As D3DMATRIX
Private matProj  As D3DMATRIX
'DX Interface Variables'
Private DX              As New DirectX8
Private D3D             As Direct3D8
Private D3DCaps         As D3DCAPS8
Private D3Dpp           As D3DPRESENT_PARAMETERS
Private D3DDevice       As Direct3DDevice8
Private D3DX            As New D3DX8
Private DI              As DirectInput8
Private DIDKeyboard     As DirectInputDevice8
Private DIDMouse        As DirectInputDevice8
Private KeyState        As DIKEYBOARDSTATE
Private MouseState      As DIMOUSESTATE
Private DS              As DirectSound8
Private Sound           As DirectSoundSecondaryBuffer8
Private Screen          As D3DDISPLAYMODE
Private ZBufferFormat   As CONST_D3DFORMAT
Private Mouse           As D3DMOUSEMODE
Private D3DSMouse       As Direct3DSurface8
Private D3DSBackground  As Direct3DSurface8
Private Frames          As D3DFRAMES
Private Font            As D3DFONT
Private Mesh()          As D3DMESH
Private RenderMesh()    As D3DMESH
Private Texture()       As Direct3DTexture8
Private TempMode        As D3DDISPLAYMODE
'Misc Vars'
Private BackgroundActive As Boolean
Private ModeEnum(100)     As Integer
Private ModeCount        As Integer

Public Function IsKey(Key As Byte) As Boolean
'Keyboard Check'
    Call DIDKeyboard.GetDeviceStateKeyboard(KeyState)
    If (KeyState.Key(Key) And &H80) Then
        IsKey = True
    Else
        IsKey = False
    End If
End Function

Public Function IsMouseB(Nr As Integer) As Boolean
    IsMouseB = Mouse.b(Nr)
End Function

Public Property Get ScreenX() As Long
    ScreenX = Screen.Width
End Property

Public Property Get ScreenY() As Long
    ScreenY = Screen.Height
End Property

Public Property Get ScreenFmt() As CONST_D3DFORMAT
    ScreenFmt = Screen.Format
End Property

Public Property Get MouseX() As Long
    MouseX = Mouse.x
End Property

Public Property Get MouseY() As Long
    MouseY = Mouse.y
End Property

Public Property Get MouseZ() As Long
    MouseZ = Mouse.z
End Property

Public Property Get CurrentFrameRate() As Long
    CurrentFrameRate = Frames.Current
End Property

Public Property Get Device() As Direct3DDevice8
    Set Device = D3DDevice
End Property

Public Property Get Class() As D3DX8
    Set Class = D3DX
End Property

Public Property Get BackgroundVisible() As Boolean
    BackgroundVisible = BackgroundActive
End Property

Public Property Let BackgroundVisible(State As Boolean)
    BackgroundActive = State
End Property

Private Function D3DCompatibilityCheck()
    Dim nModes  As Long
    Dim x       As Integer
    ModeCount = 0
'Search for Supported Screen Mode'
    nModes = D3D.GetAdapterModeCount(D3DADAPTER_DEFAULT) 'Max Display Modes
    For x = nModes - 1 To 0 Step -1
        Call D3D.EnumAdapterModes(D3DADAPTER_DEFAULT, x, TempMode)
        If (TempMode.Format = D3DFMT_X8R8G8B8) And _
            (TempMode.Height = 768) And _
            (TempMode.Width = 1024) Then
            ModeEnum(ModeCount) = x
            ModeCount = ModeCount + 1
        End If
    Next x
    For x = nModes - 1 To 0 Step -1
        Call D3D.EnumAdapterModes(D3DADAPTER_DEFAULT, x, TempMode)
        If (TempMode.Format = D3DFMT_X8R8G8B8) And _
            (TempMode.Height = 600) And _
            (TempMode.Width = 800) Then
            ModeEnum(ModeCount) = x
            ModeCount = ModeCount + 1
        End If
    Next x
    For x = nModes - 1 To 0 Step -1
        Call D3D.EnumAdapterModes(D3DADAPTER_DEFAULT, x, TempMode)
        If (TempMode.Format = D3DFMT_X8R8G8B8) And _
            (TempMode.Height = 480) And _
            (TempMode.Width = 640) Then
            ModeEnum(ModeCount) = x
            ModeCount = ModeCount + 1
        End If
    Next x
'Search for Supported ZBuffer'
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DFMT_X8R8G8B8, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D32) = D3D_OK Then
        ZBufferFormat = D3DFMT_D32
        GoTo ZBuffer_Select
    End If
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DFMT_X8R8G8B8, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
        ZBufferFormat = D3DFMT_D16
        GoTo ZBuffer_Select
    End If
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DFMT_X8R8G8B8, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D16_LOCKABLE) = D3D_OK Then
        ZBufferFormat = D3DFMT_D16_LOCKABLE
        GoTo ZBuffer_Select
    End If
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DFMT_X8R8G8B8, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D24S8) = D3D_OK Then
        ZBufferFormat = D3DFMT_D24S8
        GoTo ZBuffer_Select
    End If
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DFMT_X8R8G8B8, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D24X4S4) = D3D_OK Then
        ZBufferFormat = D3DFMT_D24X4S4
        GoTo ZBuffer_Select
    End If
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DFMT_X8R8G8B8, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D24X8) = D3D_OK Then
        ZBufferFormat = D3DFMT_D24X8
        GoTo ZBuffer_Select
    End If
    ZBufferFormat = D3DFMT_UNKNOWN
    Exit Function
ZBuffer_Select:
End Function

Public Function InitDirectX8()
    On Error GoTo D3DError
'D3D Interface Init'
    Set D3D = DX.Direct3DCreate()
    Call D3D.GetDeviceCaps(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DCaps)
'Misc'
    BackgroundActive = False
    ReDim Texture(0)
    ReDim Mesh(0)
    InitDirectX8 = True
    Exit Function
D3DError:
    InitDirectX8 = False
End Function

Public Function InitD3DAutomatic(HWin As Long) As Boolean
    Dim xCount As Integer
    On Error GoTo D3DError
    Call D3DCompatibilityCheck
    If ZBufferFormat = D3DFMT_UNKNOWN Then GoTo D3DError
    xCount = 0
NextMode:
'Get Selected Mode'
    If xCount = ModeCount Then GoTo D3DError
    Call D3D.EnumAdapterModes(D3DADAPTER_DEFAULT, ModeEnum(xCount), TempMode)
    xCount = xCount + 1
    If Not InitD3DManual(TempMode.Width, TempMode.Height, TempMode.Format, TempMode.RefreshRate, ZBufferFormat, HWin) Then GoTo NextMode
    InitD3DAutomatic = True
    Exit Function
D3DError:
    InitD3DAutomatic = False
End Function

Public Function InitD3DManual(Width As Long, Height As Long, Format As CONST_D3DFORMAT, RefreshRate As Long, ZBufferFmt As CONST_D3DFORMAT, HWin As Long) As Boolean
    On Error GoTo D3DError
'Screen'
    Screen.Format = Format
    Screen.Height = Height
    Screen.Width = Width
    Screen.RefreshRate = RefreshRate
'D3D Format Init'
    With D3Dpp
        .Windowed = False
        .BackBufferWidth = Screen.Width
        .BackBufferHeight = Screen.Height
        .BackBufferFormat = Screen.Format
        .SwapEffect = D3DSWAPEFFECT_FLIP
        .BackBufferCount = 1
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = ZBufferFmt
        .MultiSampleType = D3DMULTISAMPLE_NONE
        .FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
        .FullScreen_RefreshRateInHz = Screen.RefreshRate
    End With
'D3D Device Init'
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, HWin, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3Dpp)
    InitD3DManual = True
    Call ShowCursor(0)
    Exit Function
D3DError:
    InitD3DManual = False
End Function

Public Function InitDI(HWin As Long) As Boolean
    On Error GoTo DIError
'DI Interface Init'
    Set DI = DX.DirectInputCreate
'DI Device Init'
    Set DIDKeyboard = DI.CreateDevice("GUID_SysKeyboard")
    Set DIDMouse = DI.CreateDevice("GUID_SysMouse")
'DI Data Format Init'
    Call DIDKeyboard.SetCommonDataFormat(DIFORMAT_KEYBOARD)
    Call DIDMouse.SetCommonDataFormat(DIFORMAT_MOUSE)
'DI Cooperative LVL set'
    Call DIDKeyboard.SetCooperativeLevel(HWin, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
    Call DIDMouse.SetCooperativeLevel(HWin, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
'DI Acquire set'
    Call DIDKeyboard.Acquire
    Call DIDMouse.Acquire
    InitDI = True
    Exit Function
DIError:
    InitDI = False
End Function

Public Function InitDS(HWin As Long) As Boolean
    On Error GoTo DSError
'DS Interface Init'
    Set DS = DX.DirectSoundCreate("")
'DS Cooperative LVL set'
    Call DS.SetCooperativeLevel(HWin, DSSCL_NORMAL)
    InitDS = True
    Exit Function
DSError:
    InitDS = False
End Function

Public Function InitRenderState() As Boolean
    On Error GoTo RSError
'Shader Mode'
    Call D3DDevice.SetVertexShader(CFVF)
'Render State'
    Call D3DDevice.SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
    Call D3DDevice.SetRenderState(D3DRS_ZENABLE, D3DZB_TRUE)
    Call D3DDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
    Call D3DDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, True)
    Call D3DDevice.SetRenderState(D3DRS_SHADEMODE, D3DSHADE_GOURAUD)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC)
    InitRenderState = True
    Exit Function
RSError:
    InitRenderState = False
End Function

Public Function InitWorldGeometry() As Boolean
    On Error GoTo WError
'WORLD'
    Call D3DXMatrixIdentity(matWorld)
    Call D3DDevice.SetTransform(D3DTS_WORLD, matWorld)
'VIEW'
    Call D3DXMatrixLookAtRH(matView, CreateVector(0, 0, 1000), CreateVector(0, 0, 0), CreateVector(0, 1, 0))
    Call D3DDevice.SetTransform(D3DTS_VIEW, matView)
'PROJECTION'
    Call D3DXMatrixPerspectiveFovRH(matProj, PI / 4.28, 0.75, 100, 2000)
    Call D3DDevice.SetTransform(D3DTS_PROJECTION, matProj)
    InitWorldGeometry = True
    Exit Function
WError:
    InitWorldGeometry = False
End Function

Public Function SetLight(r As Single, g As Single, b As Single, Ambient As Long, x As Single, y As Single, z As Single) As Boolean
    On Error GoTo LError
    Dim Material As D3DMATERIAL8
    Dim Color    As D3DCOLORVALUE
    Dim Light    As D3DLIGHT8
    Color.a = 1: Color.b = 1: Color.g = 1: Color.r = 1
    Material.Ambient = Color
    Material.diffuse = Color
    Call D3DDevice.SetMaterial(Material)
    With Light
        .Type = D3DLIGHT_DIRECTIONAL
        .diffuse.r = r
        .diffuse.g = g
        .diffuse.b = b
        .Direction = CreateVector(x, y, z)
    End With
    Call D3DDevice.SetLight(0, Light)
'Set Light'
    Call D3DDevice.LightEnable(0, True)
    Call D3DDevice.SetRenderState(D3DRS_LIGHTING, True)
    Call D3DDevice.SetRenderState(D3DRS_AMBIENT, Ambient)
    SetLight = True
    Exit Function
LError:
    SetLight = False
End Function

Public Function SetMousePointer(Width As Long, Height As Long, Filename As String) As Boolean
    On Error GoTo MError
'Get Mouse Surface'
    Set D3DSMouse = D3DDevice.CreateImageSurface(Width, Height, D3DFMT_A8R8G8B8)
'Load Cursor'
    Call D3DX.LoadSurfaceFromFile(D3DSMouse, ByVal 0, ByVal 0, App.Path & "\" & Filename, ByVal 0, D3DX_FILTER_NONE, &HFF000000, ByVal 0)
'Show Cursor'
    Call D3DDevice.SetCursorProperties(0, 0, D3DSMouse)
    Call D3DDevice.ShowCursor(True)
'Init Cursor Position'
    Mouse.x = Screen.Width / 2
    Mouse.y = Screen.Height / 2
    Mouse.z = 0
    SetMousePointer = True
    Exit Function
MError:
    SetMousePointer = False
End Function

Public Function SetBackgroundPicture(Filename As String) As Boolean
    On Error GoTo BError1
'Get Background Surface 1 or 2'
    Set D3DSBackground = D3DDevice.CreateRenderTarget(Screen.Width, Screen.Height, Screen.Format, D3DMULTISAMPLE_NONE, True)
    GoTo BOk
BError1:
    On Error GoTo BError2
    Set D3DSBackground = D3DDevice.CreateImageSurface(Screen.Width, Screen.Height, Screen.Format)
BOk:
'Load Background'
    Call D3DX.LoadSurfaceFromFile(D3DSBackground, ByVal 0, CreateRect(0, 0, Screen.Width, Screen.Height), App.Path & "\" & Filename, ByVal 0, D3DX_FILTER_TRIANGLE, 0, ByVal 0)
    SetBackgroundPicture = True
    Exit Function
BError2:
    SetBackgroundPicture = False
End Function

Public Function SetTexture(TextureNr As Integer, TransparentColor As Long, Filename As String) As Boolean
    On Error GoTo TError
'Texture Init'
    If UBound(Texture) < TextureNr Then ReDim Preserve Texture(TextureNr)
    Set Texture(TextureNr) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\" & Filename, 0, 0, 0, 0, _
                     D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, TransparentColor, ByVal 0, ByVal 0)
    SetTexture = True
    Exit Function
TError:
    SetTexture = False
End Function

Public Function SetFont(Fontname As String, Size As Integer, Bold As Boolean, Italic As Boolean) As Boolean
    On Error GoTo FError
'Font Init'
    Font.Set.Name = Fontname
    Font.Set.Size = Size
    Font.Set.Bold = Bold
    Font.Set.Italic = Italic
    Set Font.Desc = Font.Set
    Set Font.New = D3DX.CreateFont(D3DDevice, Font.Desc.hFont)
    SetFont = True
    Exit Function
FError:
    SetFont = False
End Function

Public Function WriteText(Color As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, Text As String, Alignment As Long)
    Call D3DX.DrawText(Font.New, Color, Text, CreateRect(x1, y1, x2, y2), Alignment)
End Function

Public Function AddPolygonToMesh(MeshNr As Integer, _
x1 As Single, y1 As Single, z1 As Single, tu1 As Single, tv1 As Single, _
x2 As Single, y2 As Single, z2 As Single, tu2 As Single, tv2 As Single, _
x3 As Single, y3 As Single, z3 As Single, tu3 As Single, tv3 As Single)
'Polygon Point definition'
'1'
'|\'
'| \'
'2--3'
'Light Vector direct out of Screen'
    Dim VNorm As D3DVECTOR
    If UBound(Mesh) < MeshNr Then ReDim Preserve Mesh(MeshNr)
    With Mesh(MeshNr)
        ReDim Preserve .Vertices((.VertexNum + 3) - 1)
        .Vertices(.VertexNum + 0) = CreateVertex(x1, y1, z1, tu1, tv1)
        .Vertices(.VertexNum + 1) = CreateVertex(x2, y2, z2, tu2, tv2)
        .Vertices(.VertexNum + 2) = CreateVertex(x3, y3, z3, tu3, tv3)
        VNorm = CreateNormalVector(.Vertices(.VertexNum + 0), .Vertices(.VertexNum + 1), .Vertices(.VertexNum + 2))
        .Vertices(.VertexNum + 0).nx = VNorm.x: .Vertices(.VertexNum + 1).nx = VNorm.x: .Vertices(.VertexNum + 2).nx = VNorm.x
        .Vertices(.VertexNum + 0).ny = VNorm.y: .Vertices(.VertexNum + 1).ny = VNorm.y: .Vertices(.VertexNum + 2).ny = VNorm.y
        .Vertices(.VertexNum + 0).nz = VNorm.z: .Vertices(.VertexNum + 1).nz = VNorm.z: .Vertices(.VertexNum + 2).nz = VNorm.z
        ReDim Preserve .Indices((.IndexNum + 3) - 1)
        .Indices(.IndexNum + 0) = .VertexNum + 0
        .Indices(.IndexNum + 1) = .VertexNum + 1
        .Indices(.IndexNum + 2) = .VertexNum + 2
        .VertexNum = .VertexNum + 3
        .IndexNum = .IndexNum + 3
        .PolyNum = .PolyNum + 1
    End With
    Set Mesh(MeshNr).Vbuffer = D3DDevice.CreateVertexBuffer(Len(Mesh(MeshNr).Vertices(0)) * Mesh(MeshNr).VertexNum, 0, CFVF, D3DPOOL_DEFAULT)
    Call D3DVertexBuffer8SetData(Mesh(MeshNr).Vbuffer, 0, Len(Mesh(MeshNr).Vertices(0)) * Mesh(MeshNr).VertexNum, 0, Mesh(MeshNr).Vertices(0))
    Set Mesh(MeshNr).IBuffer = D3DDevice.CreateIndexBuffer(Mesh(MeshNr).IndexNum * Len(Mesh(MeshNr).Indices(0)), 0, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
    Call D3DIndexBuffer8SetData(Mesh(MeshNr).IBuffer, 0, Mesh(MeshNr).IndexNum * Len(Mesh(MeshNr).Indices(0)), 0, Mesh(MeshNr).Indices(0))
End Function

Public Function AddQuadratToMesh(MeshNr As Integer, _
x1 As Single, y1 As Single, z1 As Single, tu1 As Single, tv1 As Single, _
x2 As Single, y2 As Single, z2 As Single, tu2 As Single, tv2 As Single, _
x3 As Single, y3 As Single, z3 As Single, tu3 As Single, tv3 As Single, _
x4 As Single, y4 As Single, z4 As Single, tu4 As Single, tv4 As Single)
'Quadrat Point definition'
'1-----4'
'|     |'
'|     |'
'2-----3'
'Light Vector direct out of Screen'
    Dim VNorm As D3DVECTOR
    If UBound(Mesh) < MeshNr Then ReDim Preserve Mesh(MeshNr)
    With Mesh(MeshNr)
        ReDim Preserve .Vertices((.VertexNum + 4) - 1)
        .Vertices(.VertexNum + 0) = CreateVertex(x1, y1, z1, tu1, tv1)
        .Vertices(.VertexNum + 1) = CreateVertex(x2, y2, z2, tu2, tv2)
        .Vertices(.VertexNum + 2) = CreateVertex(x3, y3, z3, tu3, tv3)
        .Vertices(.VertexNum + 3) = CreateVertex(x4, y4, z4, tu4, tv4)
        VNorm = CreateNormalVector(.Vertices(.VertexNum + 0), .Vertices(.VertexNum + 1), .Vertices(.VertexNum + 2))
        .Vertices(.VertexNum + 0).nx = VNorm.x: .Vertices(.VertexNum + 1).nx = VNorm.x: .Vertices(.VertexNum + 2).nx = VNorm.x: .Vertices(.VertexNum + 3).nx = VNorm.x
        .Vertices(.VertexNum + 0).ny = VNorm.y: .Vertices(.VertexNum + 1).ny = VNorm.y: .Vertices(.VertexNum + 2).ny = VNorm.y: .Vertices(.VertexNum + 3).ny = VNorm.y
        .Vertices(.VertexNum + 0).nz = VNorm.z: .Vertices(.VertexNum + 1).nz = VNorm.z: .Vertices(.VertexNum + 2).nz = VNorm.z: .Vertices(.VertexNum + 3).nz = VNorm.z
        ReDim Preserve .Indices((.IndexNum + 6) - 1)
        .Indices(.IndexNum + 0) = .VertexNum + 0
        .Indices(.IndexNum + 1) = .VertexNum + 1
        .Indices(.IndexNum + 2) = .VertexNum + 2
        .Indices(.IndexNum + 3) = .VertexNum + 2
        .Indices(.IndexNum + 4) = .VertexNum + 3
        .Indices(.IndexNum + 5) = .VertexNum + 0
        .VertexNum = .VertexNum + 4
        .IndexNum = .IndexNum + 6
        .PolyNum = .PolyNum + 2
    End With
    Set Mesh(MeshNr).Vbuffer = D3DDevice.CreateVertexBuffer(Len(Mesh(MeshNr).Vertices(0)) * Mesh(MeshNr).VertexNum, 0, CFVF, D3DPOOL_DEFAULT)
    Call D3DVertexBuffer8SetData(Mesh(MeshNr).Vbuffer, 0, Len(Mesh(MeshNr).Vertices(0)) * Mesh(MeshNr).VertexNum, 0, Mesh(MeshNr).Vertices(0))
    Set Mesh(MeshNr).IBuffer = D3DDevice.CreateIndexBuffer(Mesh(MeshNr).IndexNum * Len(Mesh(MeshNr).Indices(0)), 0, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
    Call D3DIndexBuffer8SetData(Mesh(MeshNr).IBuffer, 0, Mesh(MeshNr).IndexNum * Len(Mesh(MeshNr).Indices(0)), 0, Mesh(MeshNr).Indices(0))
End Function

Public Function AddTextureToMesh(MeshNr As Integer, TextureNr As Integer)
    If UBound(Mesh) < MeshNr Then ReDim Preserve Mesh(MeshNr)
    Mesh(MeshNr).Texture = TextureNr
End Function

Public Function CopyMeshToMesh(TargetMeshNr As Integer, SourceMeshNr As Integer, ScaleMultiplier As Single)
    Dim x As Integer
    If UBound(Mesh) < TargetMeshNr Then ReDim Preserve Mesh(TargetMeshNr)
    ReDim Mesh(TargetMeshNr).Vertices(Mesh(SourceMeshNr).VertexNum - 1)
    For x = 0 To Mesh(SourceMeshNr).VertexNum - 1
        Mesh(TargetMeshNr).Vertices(x).nx = Mesh(SourceMeshNr).Vertices(x).nx
        Mesh(TargetMeshNr).Vertices(x).ny = Mesh(SourceMeshNr).Vertices(x).ny
        Mesh(TargetMeshNr).Vertices(x).nz = Mesh(SourceMeshNr).Vertices(x).nz
        Mesh(TargetMeshNr).Vertices(x).tu = Mesh(SourceMeshNr).Vertices(x).tu
        Mesh(TargetMeshNr).Vertices(x).tv = Mesh(SourceMeshNr).Vertices(x).tv
        Mesh(TargetMeshNr).Vertices(x).x = Mesh(SourceMeshNr).Vertices(x).x * ScaleMultiplier
        Mesh(TargetMeshNr).Vertices(x).y = Mesh(SourceMeshNr).Vertices(x).y * ScaleMultiplier
        Mesh(TargetMeshNr).Vertices(x).z = Mesh(SourceMeshNr).Vertices(x).z * ScaleMultiplier
    Next x
    ReDim Mesh(TargetMeshNr).Indices(Mesh(SourceMeshNr).IndexNum - 1)
    For x = 0 To Mesh(SourceMeshNr).IndexNum - 1
        Mesh(TargetMeshNr).Indices(x) = Mesh(SourceMeshNr).Indices(x)
    Next x
    Mesh(TargetMeshNr).IndexNum = Mesh(SourceMeshNr).IndexNum
    Mesh(TargetMeshNr).Indices = Mesh(SourceMeshNr).Indices
    Mesh(TargetMeshNr).PolyNum = Mesh(SourceMeshNr).PolyNum
    Mesh(TargetMeshNr).Texture = Mesh(SourceMeshNr).Texture
    Mesh(TargetMeshNr).VertexNum = Mesh(SourceMeshNr).VertexNum
    Mesh(TargetMeshNr).Draw = Mesh(SourceMeshNr).Draw
    Set Mesh(TargetMeshNr).Vbuffer = D3DDevice.CreateVertexBuffer(Len(Mesh(TargetMeshNr).Vertices(0)) * Mesh(TargetMeshNr).VertexNum, 0, CFVF, D3DPOOL_DEFAULT)
    Call D3DVertexBuffer8SetData(Mesh(TargetMeshNr).Vbuffer, 0, Len(Mesh(TargetMeshNr).Vertices(0)) * Mesh(TargetMeshNr).VertexNum, 0, Mesh(TargetMeshNr).Vertices(0))
    Set Mesh(TargetMeshNr).IBuffer = D3DDevice.CreateIndexBuffer(Mesh(TargetMeshNr).IndexNum * Len(Mesh(TargetMeshNr).Indices(0)), 0, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
    Call D3DIndexBuffer8SetData(Mesh(TargetMeshNr).IBuffer, 0, Mesh(TargetMeshNr).IndexNum * Len(Mesh(TargetMeshNr).Indices(0)), 0, Mesh(TargetMeshNr).Indices(0))
End Function

Public Function RenderMeshes() As Boolean
    On Error GoTo RError
    Dim x As Integer
    For x = 0 To UBound(RenderMesh) Step 1
        Call D3DDevice.SetTexture(0, Texture(RenderMesh(x).Texture))
        Call D3DDevice.SetStreamSource(0, RenderMesh(x).Vbuffer, Len(RenderMesh(x).Vertices(0)))
        Call D3DDevice.SetIndices(RenderMesh(x).IBuffer, 0)
        Call D3DDevice.DrawIndexedPrimitive(D3DPT_TRIANGLELIST, 0, RenderMesh(x).IndexNum, 0, RenderMesh(x).PolyNum)
    Next x
    RenderMeshes = True
    Exit Function
RError:
    RenderMeshes = False
End Function

Public Function Freeze(ms As Long)
    Dim Wait As Long
    Wait = GetTickCount
    While GetTickCount < (Wait + ms)
        DoEvents
    Wend
End Function

Public Function Cleanup()
    On Error Resume Next
'Direct 3D'
    Set D3DDevice = Nothing
    Set D3DX = Nothing
    Set D3D = Nothing
    Set DX = Nothing
'Direct Input'
    DIDKeyboard.Unacquire
    DIDMouse.Unacquire
    Set DIDKeyboard = Nothing
    Set DIDMouse = Nothing
    Set DI = Nothing
'Direct Sound'
    Set Sound = Nothing
    Set DS = Nothing
'Buffer leeren'
    Set D3DSBackground = Nothing
    Set D3DSMouse = Nothing
    Erase Mesh
    Erase Texture
'Exit'
    Call ShowCursor(1)
    End
End Function

Public Function RefreshData()
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFF0000FF, 1#, 0)
'Set Background'
    If BackgroundActive Then
        Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, False)
        Call D3DDevice.CopyRects(D3DSBackground, ByVal 0, 0, D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), ByVal 0)
        Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, True)
    End If
    Call UpdateMousePointer
    Call UpdateFrames
'Inizialise Meshes'
    RenderMesh = Mesh
End Function

Private Function UpdateMousePointer()
'Mouse Check'
    Call DIDMouse.GetDeviceStateMouse(MouseState)
    Mouse.x = Mouse.x + MouseState.lX
    Mouse.y = Mouse.y + MouseState.lY
    Mouse.z = Mouse.z + (MouseState.lZ / 12)
    If Mouse.x < 0 Then Mouse.x = 0
    If Mouse.y < 0 Then Mouse.y = 0
    If Mouse.z < -500 Then Mouse.z = -500
    If Mouse.x > (Screen.Width - 1) Then Mouse.x = Screen.Width - 1
    If Mouse.y > (Screen.Height - 1) Then Mouse.y = Screen.Height - 1
    If Mouse.z > 500 Then Mouse.z = 500
    Mouse.b(0) = MouseState.Buttons(0)
    Mouse.b(1) = MouseState.Buttons(1)
    Mouse.b(2) = MouseState.Buttons(2)
'Set Mouse Cursor'
    Call D3DDevice.SetCursorPosition(Mouse.x, Mouse.y, D3DUCURSOR_DEFAULT)
End Function

Private Function UpdateFrames()
    If GetTickCount - Frames.Last >= 1000 Then
        Frames.Current = Frames.Count
        If Frames.Current = 0 Then Frames.Current = 100 'Div by 0 safe
        Frames.Count = 0
        Frames.Last = GetTickCount
    End If
    Frames.Count = Frames.Count + 1
End Function

Private Function CreateVertex(x As Single, y As Single, z As Single, tu As Single, tv As Single) As CVERTEX
    With CreateVertex
        .x = x
        .y = y
        .z = z
        .nx = 0
        .ny = 0
        .nz = 0
        .tu = tu
        .tv = tv
    End With
End Function

Private Function CreateVector(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
    With CreateVector
        .x = x
        .y = y
        .z = z
    End With
End Function

Private Function CreateRect(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As RECT
    With CreateRect
        .bottom = y2
        .Left = x1
        .Right = x2
        .Top = y1
    End With
End Function

Private Function CreateNormalVector(p0 As CVERTEX, p1 As CVERTEX, p2 As CVERTEX) As D3DVECTOR
    Dim v01   As D3DVECTOR 'Vector from points 0 to 1
    Dim v02   As D3DVECTOR 'Vector from points 0 to 2
    Dim VNorm As D3DVECTOR 'The final vector

    D3DXVec3Subtract v01, CreateVector(p1.x, p1.y, p1.z), CreateVector(p0.x, p0.y, p0.z)
    D3DXVec3Subtract v02, CreateVector(p2.x, p2.y, p2.z), CreateVector(p0.x, p0.y, p0.z)
    D3DXVec3Cross VNorm, v01, v02
    D3DXVec3Normalize VNorm, VNorm
    CreateNormalVector.x = VNorm.x
    CreateNormalVector.y = VNorm.y
    CreateNormalVector.z = VNorm.z
End Function

Public Function RotateMeshAroundYAxis(MeshNr As Integer, RotationAngle As Single)
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    Dim RadAngle As Single
    Dim NewX     As Single
    Dim NewZ     As Single
    Dim x        As Integer
    
    RadAngle = RotationAngle * Radiant
    CosPhi = Cos(RadAngle)
    SinPhi = Sin(RadAngle)
    
    For x = 0 To (RenderMesh(MeshNr).VertexNum - 1) Step 1
        NewX = RenderMesh(MeshNr).Vertices(x).x * CosPhi - RenderMesh(MeshNr).Vertices(x).z * SinPhi
        NewZ = RenderMesh(MeshNr).Vertices(x).x * SinPhi + RenderMesh(MeshNr).Vertices(x).z * CosPhi
        RenderMesh(MeshNr).Vertices(x).x = NewX
        RenderMesh(MeshNr).Vertices(x).z = NewZ
        NewX = RenderMesh(MeshNr).Vertices(x).nx * CosPhi - RenderMesh(MeshNr).Vertices(x).nz * SinPhi
        NewZ = RenderMesh(MeshNr).Vertices(x).nx * SinPhi + RenderMesh(MeshNr).Vertices(x).nz * CosPhi
        RenderMesh(MeshNr).Vertices(x).nx = NewX
        RenderMesh(MeshNr).Vertices(x).nz = NewZ
    Next x
    Call D3DVertexBuffer8SetData(RenderMesh(MeshNr).Vbuffer, 0, Len(RenderMesh(MeshNr).Vertices(0)) * RenderMesh(MeshNr).VertexNum, 0, RenderMesh(MeshNr).Vertices(0))
End Function

Public Function RotateMeshAroundZAxis(MeshNr As Integer, RotationAngle As Single)
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    Dim RadAngle As Single
    Dim NewX     As Single
    Dim NewY     As Single
    Dim x        As Integer
    
    RadAngle = RotationAngle * Radiant * -1
    CosPhi = Cos(RadAngle)
    SinPhi = Sin(RadAngle)
    
    For x = 0 To (RenderMesh(MeshNr).VertexNum - 1) Step 1
        NewX = RenderMesh(MeshNr).Vertices(x).x * CosPhi - RenderMesh(MeshNr).Vertices(x).y * SinPhi
        NewY = RenderMesh(MeshNr).Vertices(x).x * SinPhi + RenderMesh(MeshNr).Vertices(x).y * CosPhi
        RenderMesh(MeshNr).Vertices(x).x = NewX
        RenderMesh(MeshNr).Vertices(x).y = NewY
        NewX = RenderMesh(MeshNr).Vertices(x).nx * CosPhi - RenderMesh(MeshNr).Vertices(x).ny * SinPhi
        NewY = RenderMesh(MeshNr).Vertices(x).nx * SinPhi + RenderMesh(MeshNr).Vertices(x).ny * CosPhi
        RenderMesh(MeshNr).Vertices(x).nx = NewX
        RenderMesh(MeshNr).Vertices(x).ny = NewY
    Next x
    Call D3DVertexBuffer8SetData(RenderMesh(MeshNr).Vbuffer, 0, Len(RenderMesh(MeshNr).Vertices(0)) * RenderMesh(MeshNr).VertexNum, 0, RenderMesh(MeshNr).Vertices(0))
End Function

Public Function RotateMeshAroundXAxis(MeshNr As Integer, RotationAngle As Single)
    Dim CosPhi   As Single
    Dim SinPhi   As Single
    Dim RadAngle As Single
    Dim NewZ     As Single
    Dim NewY     As Single
    Dim x        As Integer
    
    RadAngle = RotationAngle * Radiant
    CosPhi = Cos(RadAngle)
    SinPhi = Sin(RadAngle)
    
    For x = 0 To (RenderMesh(MeshNr).VertexNum - 1) Step 1
        NewZ = RenderMesh(MeshNr).Vertices(x).z * CosPhi - RenderMesh(MeshNr).Vertices(x).y * SinPhi
        NewY = RenderMesh(MeshNr).Vertices(x).z * SinPhi + RenderMesh(MeshNr).Vertices(x).y * CosPhi
        RenderMesh(MeshNr).Vertices(x).z = NewZ
        RenderMesh(MeshNr).Vertices(x).y = NewY
        NewZ = RenderMesh(MeshNr).Vertices(x).nz * CosPhi - RenderMesh(MeshNr).Vertices(x).ny * SinPhi
        NewY = RenderMesh(MeshNr).Vertices(x).nz * SinPhi + RenderMesh(MeshNr).Vertices(x).ny * CosPhi
        RenderMesh(MeshNr).Vertices(x).nz = NewZ
        RenderMesh(MeshNr).Vertices(x).ny = NewY
    Next x
    Call D3DVertexBuffer8SetData(RenderMesh(MeshNr).Vbuffer, 0, Len(RenderMesh(MeshNr).Vertices(0)) * RenderMesh(MeshNr).VertexNum, 0, RenderMesh(MeshNr).Vertices(0))
End Function

Public Function MoveMeshAlongXAxis(MeshNr As Integer, Range As Single)
    Dim NewX     As Single
    Dim x        As Integer
    
    For x = 0 To (RenderMesh(MeshNr).VertexNum - 1) Step 1
        NewX = RenderMesh(MeshNr).Vertices(x).x + Range
        RenderMesh(MeshNr).Vertices(x).x = NewX
    Next x
    Call D3DVertexBuffer8SetData(RenderMesh(MeshNr).Vbuffer, 0, Len(RenderMesh(MeshNr).Vertices(0)) * RenderMesh(MeshNr).VertexNum, 0, RenderMesh(MeshNr).Vertices(0))
End Function

Public Function MoveMeshAlongYAxis(MeshNr As Integer, Range As Single)
    Dim NewY     As Single
    Dim x        As Integer
    
    For x = 0 To (RenderMesh(MeshNr).VertexNum - 1) Step 1
        NewY = RenderMesh(MeshNr).Vertices(x).y + Range
        RenderMesh(MeshNr).Vertices(x).y = NewY
    Next x
    Call D3DVertexBuffer8SetData(RenderMesh(MeshNr).Vbuffer, 0, Len(RenderMesh(MeshNr).Vertices(0)) * RenderMesh(MeshNr).VertexNum, 0, RenderMesh(MeshNr).Vertices(0))
End Function

Public Function MoveMeshAlongZAxis(MeshNr As Integer, Range As Single)
    Dim NewZ     As Single
    Dim x        As Integer
    
    For x = 0 To (RenderMesh(MeshNr).VertexNum - 1) Step 1
        NewZ = RenderMesh(MeshNr).Vertices(x).z + Range
        RenderMesh(MeshNr).Vertices(x).z = NewZ
    Next x
    Call D3DVertexBuffer8SetData(RenderMesh(MeshNr).Vbuffer, 0, Len(RenderMesh(MeshNr).Vertices(0)) * RenderMesh(MeshNr).VertexNum, 0, RenderMesh(MeshNr).Vertices(0))
End Function

