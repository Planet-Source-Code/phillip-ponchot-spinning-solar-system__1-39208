VERSION 5.00
Object = "{08216199-47EA-11D3-9479-00AA006C473C}#2.1#0"; "RMCONTROL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8865
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   8040
      Top             =   1080
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   7200
      Top             =   1080
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   6240
      Top             =   1200
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   5280
      Top             =   1200
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   4440
      Top             =   1320
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   3480
      Top             =   1320
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   2520
      Top             =   1320
   End
   Begin VB.Timer Timer3 
      Interval        =   20000
      Left            =   1680
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   0
      Top             =   2280
   End
   Begin MSComDlg.CommonDialog D 
      Left            =   5760
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3000
      Top             =   3480
   End
   Begin RMControl7.RMCanvas RM 
      Height          =   6135
      Left            =   -120
      TabIndex        =   0
      Top             =   3000
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10821
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   7200
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   7200
      Width           =   4695
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare API to hide and reshow the mouse pointer
Dim switchcount As Integer
Dim switchcount4 As Integer
Dim switchcount5 As Integer
Dim switchcount6 As Integer
Dim switchcount7 As Integer
Dim switchcount8 As Integer
Dim switchcount9 As Integer
Dim switchcount10 As Integer

Private Declare Function ShowCursor _
Lib "user32" ( _
    ByVal bShow As Long _
) As Long

' Frame:    Holds data for a 3D Object
' Mesh:     What an object looks like 3D or "all around"
' Texture: The "skin" of an object, usually a picture.

'**Background Star
Dim FR_BS1(100) As Direct3DRMFrame3             ' Frame for holding data of the sphere
Dim MS_BS1(100) As Direct3DRMMeshBuilder3   ' Mesh object for the sphere (what it looks like)
Dim TX_BS1(100) As Direct3DRMTexture3          ' The texture of our sphere
'*****************

Dim FR_BallS As Direct3DRMFrame3             ' Frame for holding data of the sphere
Dim MS_BallS As Direct3DRMMeshBuilder3   ' Mesh object for the sphere (what it looks like)
Dim TX_BallS As Direct3DRMTexture3          ' The texture of our sphere

Dim FR_Ball2 As Direct3DRMFrame3             ' Frame for holding data of the sphere
Dim MS_Ball2 As Direct3DRMMeshBuilder3   ' Mesh object for the sphere (what it looks like)
Dim TX_Ball2 As Direct3DRMTexture3          ' The texture of our sphere

Dim FR_Ball3 As Direct3DRMFrame3             ' Frame for holding data of the sphere
Dim MS_Ball3 As Direct3DRMMeshBuilder3   ' Mesh object for the sphere (what it looks like)
Dim TX_Ball3 As Direct3DRMTexture3          ' The texture of our sphere

Dim FR_Ball4 As Direct3DRMFrame3             ' Frame for holding data of the sphere
Dim MS_Ball4 As Direct3DRMMeshBuilder3   ' Mesh object for the sphere (what it looks like)
Dim TX_Ball4 As Direct3DRMTexture3          ' The texture of our sphere

Dim FR_Ball5 As Direct3DRMFrame3             ' Frame for holding data of the sphere
Dim MS_Ball5 As Direct3DRMMeshBuilder3   ' Mesh object for the sphere (what it looks like)
Dim TX_Ball5 As Direct3DRMTexture3          ' The texture of our sphere

Dim FR_Ball6 As Direct3DRMFrame3             ' Frame for holding data of the sphere
Dim MS_Ball6 As Direct3DRMMeshBuilder3   ' Mesh object for the sphere (what it looks like)
Dim TX_Ball6 As Direct3DRMTexture3          ' The texture of our sphere

Dim FR_Ball7 As Direct3DRMFrame3             ' Frame for holding data of the sphere
Dim MS_Ball7 As Direct3DRMMeshBuilder3   ' Mesh object for the sphere (what it looks like)
Dim TX_Ball7 As Direct3DRMTexture3          ' The texture of our sphere

'**Background Star
Dim FR_BackStar(100) As Direct3DRMFrame3
Dim MS_BackStar(100) As Direct3DRMMeshBuilder3 ' Mesh for the Venus
Dim TX_BackStar(100) As Direct3DRMTexture3
'***************************************

Dim FR_Sun As Direct3DRMFrame3
Dim MS_Sun As Direct3DRMMeshBuilder3 ' Mesh for the Venus
Dim TX_Sun As Direct3DRMTexture3

Dim FR_Second As Direct3DRMFrame3
Dim MS_Second As Direct3DRMMeshBuilder3 ' Mesh for the Venus
Dim TX_Second As Direct3DRMTexture3

Dim FR_Third As Direct3DRMFrame3
Dim MS_Third As Direct3DRMMeshBuilder3 ' Mesh for the Earth
Dim TX_Third As Direct3DRMTexture3

Dim FR_Fourth As Direct3DRMFrame3
Dim MS_Fourth As Direct3DRMMeshBuilder3 ' Mesh for the Earth
Dim TX_Fourth As Direct3DRMTexture3

Dim FR_Fifth As Direct3DRMFrame3
Dim MS_Fifth As Direct3DRMMeshBuilder3 ' Mesh for the Earth
Dim TX_Fifth As Direct3DRMTexture3

Dim FR_Star As Direct3DRMFrame3
Dim MS_Star As Direct3DRMMeshBuilder3 ' Mesh for the Earth
Dim TX_Star As Direct3DRMTexture3

Dim FR_Star1 As Direct3DRMFrame3
Dim MS_Star1 As Direct3DRMMeshBuilder3 ' Mesh for the Earth
Dim TX_Star1 As Direct3DRMTexture3
Dim QuitFlag As Boolean
Const Sin5 = 8.715574E-02!                      ' 5 Degrees
Const Cos5 = 0.9961947!                         ' 5 Degrees
Public Sub DX_Init() ' Initialize our objects
Dim intcount As Integer
intcount = 0
With RM
    .StartWindowed ' Start our 3D Scene
    .Viewport.SetBack 500 ' How far we can see back without objects disappearing
    .SceneFrame.SetSceneBackgroundRGB 0#, 0#, 0#    ' Background color
    '**Background Star
    While intcount <= 100
     Set FR_BS1(intcount) = .D3DRM.CreateFrame(.SceneFrame)
     intcount = intcount + 1
    Wend
    intcount = 0
    '*****************************************
    Set FR_BallS = .D3DRM.CreateFrame(.SceneFrame)
    Set FR_Ball2 = .D3DRM.CreateFrame(.SceneFrame)
    Set FR_Ball3 = .D3DRM.CreateFrame(.SceneFrame)
    Set FR_Ball4 = .D3DRM.CreateFrame(.SceneFrame)
    Set FR_Ball5 = .D3DRM.CreateFrame(.SceneFrame)
    Set FR_Ball6 = .D3DRM.CreateFrame(.SceneFrame)
    Set FR_Ball7 = .D3DRM.CreateFrame(.SceneFrame)
    '**Background Star
    While intcount <= 100
     Set FR_BackStar(intcount) = .D3DRM.CreateFrame(FR_BS1(intcount))
     intcount = intcount + 1
    Wend
    '*************************
    Set FR_Sun = .D3DRM.CreateFrame(FR_BallS)
    Set FR_Second = .D3DRM.CreateFrame(FR_Ball2)
    Set FR_Third = .D3DRM.CreateFrame(FR_Ball3)
    Set FR_Fourth = .D3DRM.CreateFrame(FR_Ball4)
    Set FR_Fifth = .D3DRM.CreateFrame(FR_Ball5)
    Set FR_Star = .D3DRM.CreateFrame(FR_Ball6)
    Set FR_Star1 = .D3DRM.CreateFrame(FR_Ball7)
End With
End Sub
Public Sub DX_MakeObjects() ' Make objects and visualizes them for the user to see

Set MS_BallS = RM.D3DRM.CreateMeshBuilder() ' Create a mesh builder for our sphere
MS_BallS.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' This is the 3D object

Set MS_Ball2 = RM.D3DRM.CreateMeshBuilder() ' Create a mesh builder for our sphere
MS_Ball2.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' This is the 3D object

Set MS_Ball3 = RM.D3DRM.CreateMeshBuilder() ' Create a mesh builder for our sphere
MS_Ball3.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' This is the 3D object

Set MS_Ball4 = RM.D3DRM.CreateMeshBuilder() ' Create a mesh builder for our sphere
MS_Ball4.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' This is the 3D object

Set MS_Ball5 = RM.D3DRM.CreateMeshBuilder() ' Create a mesh builder for our sphere
MS_Ball5.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' This is the 3D object

Set MS_Ball6 = RM.D3DRM.CreateMeshBuilder() ' Create a mesh builder for our sphere
MS_Ball6.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' This is the 3D object

Set MS_Ball7 = RM.D3DRM.CreateMeshBuilder() ' Create a mesh builder for our sphere
MS_Ball7.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' This is the 3D object

'*****Starting Background Star Setup
Set MS_BS1(0) = RM.D3DRM.CreateMeshBuilder() ' Create a mesh builder for our sphere
MS_BS1(0).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' This is the 3D object

Set MS_BS1(1) = RM.D3DRM.CreateMeshBuilder() ' Create a mesh builder for our sphere
MS_BS1(1).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' This is the 3D object

Set MS_BS1(2) = RM.D3DRM.CreateMeshBuilder() ' Create a mesh builder for our sphere
MS_BS1(2).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing ' This is the 3D object

'*************************************
End Sub
Private Sub Form_Load()

switchcount = 0
switchcount4 = 0
switchcount5 = 0
switchcount6 = 0
switchcount7 = 0
switchcount8 = 0
switchcount9 = 0
switchcount10 = 0

QuitFlag = False
'X = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 0, ByVal 0&, 0)
x = ShowCursor(False)
FrmMain.Width = Screen.Width
FrmMain.Height = Screen.Height
FrmMain.Left = 0
FrmMain.Top = 0
RM.Height = FrmMain.Height
RM.Width = FrmMain.Width
RM.Top = 0
RM.Left = 0

DX_Init
DX_MakeObjects

'***BackStar
Set MS_BackStar(0) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(0) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(0).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(0).SetTexture TX_BackStar(0)
FR_BackStar(0).SetPosition FR_BS1(0), 4, 0, 5
'change size
MS_BackStar(0).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(0).AddVisual MS_BackStar(0)

Set MS_BackStar(1) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\pink.bmp"
Set TX_BackStar(1) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(1).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(1).SetTexture TX_BackStar(1)
FR_BackStar(1).SetPosition FR_BS1(1), 2, 0, 5
'change size
MS_BackStar(1).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(1).AddVisual MS_BackStar(1)

Set MS_BackStar(2) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(2) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(2).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(2).SetTexture TX_BackStar(2)
FR_BackStar(2).SetPosition FR_BS1(2), 2, 0, 3
'change size
MS_BackStar(2).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(2).AddVisual MS_BackStar(2)

Set MS_BackStar(3) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(3) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(3).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(3).SetTexture TX_BackStar(3)
FR_BackStar(3).SetPosition FR_BS1(3), 2, 3, 3
'change size
MS_BackStar(3).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(3).AddVisual MS_BackStar(3)

Set MS_BackStar(4) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\pink.bmp"
Set TX_BackStar(4) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(4).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(4).SetTexture TX_BackStar(4)
FR_BackStar(4).SetPosition FR_BS1(4), 2, -2, 3
'change size
MS_BackStar(4).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(4).AddVisual MS_BackStar(4)

Set MS_BackStar(5) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(5) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(5).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(5).SetTexture TX_BackStar(5)
FR_BackStar(5).SetPosition FR_BS1(5), 1, -3, 4
'change size
MS_BackStar(5).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(5).AddVisual MS_BackStar(5)

Set MS_BackStar(6) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(6) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(6).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(6).SetTexture TX_BackStar(6)
FR_BackStar(6).SetPosition FR_BS1(6), 2, -4, 4
'change size
MS_BackStar(6).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(6).AddVisual MS_BackStar(6)

Set MS_BackStar(7) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\green.bmp"
Set TX_BackStar(7) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(7).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(7).SetTexture TX_BackStar(7)
FR_BackStar(7).SetPosition FR_BS1(7), 3, -3, 3
'change size
MS_BackStar(7).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(7).AddVisual MS_BackStar(7)


Set MS_BackStar(8) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\white.bmp"
Set TX_BackStar(8) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(8).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(8).SetTexture TX_BackStar(8)
FR_BackStar(8).SetPosition FR_BS1(8), 4, -4, 4
'change size
MS_BackStar(8).ScaleMesh 0.03, 0.03, 0.03
FR_BackStar(8).AddVisual MS_BackStar(8)


Set MS_BackStar(9) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(9) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(9).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(9).SetTexture TX_BackStar(9)
FR_BackStar(9).SetPosition FR_BS1(9), 4, -2, -1
'change size
MS_BackStar(9).ScaleMesh 0.04, 0.04, 0.04
FR_BackStar(9).AddVisual MS_BackStar(9)

Set MS_BackStar(10) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(10) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(10).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(10).SetTexture TX_BackStar(10)
FR_BackStar(10).SetPosition FR_BS1(10), 4, -3, -2
'change size
MS_BackStar(10).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(10).AddVisual MS_BackStar(10)

Set MS_BackStar(11) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(11) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(11).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(11).SetTexture TX_BackStar(11)
FR_BackStar(11).SetPosition FR_BS1(11), -2, -2, -2
'change size
MS_BackStar(11).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(11).AddVisual MS_BackStar(11)


Set MS_BackStar(12) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\white.bmp"
Set TX_BackStar(12) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(12).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(12).SetTexture TX_BackStar(12)
FR_BackStar(12).SetPosition FR_BS1(12), -3, -1, -2
'change size
MS_BackStar(12).ScaleMesh 0.03, 0.03, 0.03
FR_BackStar(12).AddVisual MS_BackStar(12)


Set MS_BackStar(13) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(13) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(13).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(13).SetTexture TX_BackStar(13)
FR_BackStar(13).SetPosition FR_BS1(13), -2, 0, -2
'change size
MS_BackStar(13).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(13).AddVisual MS_BackStar(13)

Set MS_BackStar(14) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(14) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(14).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(14).SetTexture TX_BackStar(14)
FR_BackStar(14).SetPosition FR_BS1(14), -3, 1, -2
'change size
MS_BackStar(14).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(14).AddVisual MS_BackStar(14)

Set MS_BackStar(15) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(15) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(15).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(15).SetTexture TX_BackStar(15)
FR_BackStar(15).SetPosition FR_BS1(15), -4, 3, -4
'change size
MS_BackStar(15).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(15).AddVisual MS_BackStar(15)

Set MS_BackStar(16) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(16) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(16).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(16).SetTexture TX_BackStar(16)
FR_BackStar(16).SetPosition FR_BS1(16), -1, 1, 6
'change size
MS_BackStar(16).ScaleMesh 0.04, 0.04, 0.04
FR_BackStar(16).AddVisual MS_BackStar(16)

Set MS_BackStar(17) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(17) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(17).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(17).SetTexture TX_BackStar(17)
FR_BackStar(17).SetPosition FR_BS1(17), -1, 3, -3
'change size
MS_BackStar(17).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(17).AddVisual MS_BackStar(17)

Set MS_BackStar(18) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(18) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(18).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(18).SetTexture TX_BackStar(18)
FR_BackStar(18).SetPosition FR_BS1(18), -2, 3, 1
'change size
MS_BackStar(18).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(18).AddVisual MS_BackStar(18)


Set MS_BackStar(19) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(19) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(19).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(19).SetTexture TX_BackStar(19)
FR_BackStar(19).SetPosition FR_BS1(19), 4, 4, 3
'change size
MS_BackStar(19).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(19).AddVisual MS_BackStar(19)

Set MS_BackStar(20) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(20) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(20).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(20).SetTexture TX_BackStar(20)
FR_BackStar(20).SetPosition FR_BS1(20), -4, 3, -4
'change size
MS_BackStar(20).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(20).AddVisual MS_BackStar(20)

Set MS_BackStar(21) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(21) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(21).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(21).SetTexture TX_BackStar(21)
FR_BackStar(21).SetPosition FR_BS1(21), -4, 3, 1
'change size
MS_BackStar(21).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(21).AddVisual MS_BackStar(21)

Set MS_BackStar(22) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(22) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(22).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(22).SetTexture TX_BackStar(22)
FR_BackStar(22).SetPosition FR_BS1(22), 0, 2, 0
'change size
MS_BackStar(22).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(22).AddVisual MS_BackStar(22)

Set MS_BackStar(23) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\green.bmp"
Set TX_BackStar(23) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(23).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(23).SetTexture TX_BackStar(23)
FR_BackStar(23).SetPosition FR_BS1(23), 0, -2, 0
'change size
MS_BackStar(23).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(23).AddVisual MS_BackStar(23)


Set MS_BackStar(24) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(24) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(24).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(24).SetTexture TX_BackStar(24)
FR_BackStar(24).SetPosition FR_BS1(24), 3, 1, -1
'change size
MS_BackStar(24).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(24).AddVisual MS_BackStar(24)

'***************************************************

Set MS_Sun = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\sun.bmp"
Set TX_Sun = RM.D3DRM.LoadTexture(texture1)
MS_Sun.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_Sun.SetTexture TX_Sun
FR_Sun.SetPosition FR_BallS, 0, 0, 0
'change size
MS_Sun.ScaleMesh 0.8, 0.8, 0.8
FR_Sun.AddVisual MS_Sun

Set MS_Second = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\merc.bmp"
Set TX_Second = RM.D3DRM.LoadTexture(texture1)
MS_Second.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_Second.SetTexture TX_Second
FR_Second.SetPosition FR_Ball2, 2, 0, 0
MS_Second.ScaleMesh 0.4, 0.4, 0.4
FR_Second.AddVisual MS_Second

Set MS_Third = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\venus.bmp"
Set TX_Third = RM.D3DRM.LoadTexture(texture1)
MS_Third.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_Third.SetTexture TX_Third
FR_Third.SetPosition FR_Ball3, 2.8, 0.5, 0
MS_Third.ScaleMesh 0.4, 0.4, 0.4
FR_Third.AddVisual MS_Third

Set MS_Fourth = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\earth.bmp"
Set TX_Fourth = RM.D3DRM.LoadTexture(texture1)
MS_Fourth.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_Fourth.SetTexture TX_Fourth
FR_Fourth.SetPosition FR_Ball4, 2.5, -0.5, 3
MS_Fourth.ScaleMesh 0.4, 0.4, 0.4
FR_Fourth.AddVisual MS_Fourth

Set MS_Fifth = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\water.bmp"
Set TX_Fifth = RM.D3DRM.LoadTexture(texture1)
MS_Fifth.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_Fifth.SetTexture TX_Fifth
FR_Fifth.SetPosition FR_Ball5, 4, 0.5, 0
MS_Fifth.ScaleMesh 0.4, 0.4, 0.4
FR_Fifth.AddVisual MS_Fifth

Set MS_Star = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\earth2.bmp"
Set TX_Star = RM.D3DRM.LoadTexture(texture1)
MS_Star.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_Star.SetTexture TX_Star
' firstnumber = left,right
'secondnumber = up,down
'thirdnumber = front/back
FR_Star.SetPosition FR_Ball6, -3, -1, 0
MS_Star.ScaleMesh 0.4, 0.4, 0.4
FR_Star.AddVisual MS_Star

Set MS_Star1 = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\jupiter.Bmp"
Set TX_Star1 = RM.D3DRM.LoadTexture(texture1)
MS_Star1.LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_Star1.SetTexture TX_Star1
FR_Star1.SetPosition FR_Ball7, 1.5, 1, 0
MS_Star1.ScaleMesh 0.5, 0.5, 0.5
FR_Star1.AddVisual MS_Star1

MS_BallS.ScaleMesh 1, 1, 1
FR_BallS.SetRotation FR_BallS, 0, -Sin5, 0, 0.03

MS_Ball2.ScaleMesh 1, 1, 1
FR_Ball2.SetRotation FR_Ball2, 0, -Sin5, 0, 0.06
'FR_Ball2.SetOrientation FR_Ball2, Sin5, 0, Cos5, 0, 2, 0 ' Spin sphere left

MS_Ball3.ScaleMesh 1, 1, 1
FR_Ball3.SetRotation FR_Ball3, 0, -Sin5, 0, 0.03

MS_Ball4.ScaleMesh 1, 1, 1
FR_Ball4.SetRotation FR_Ball4, 0, -Sin5, 0, 0.02

MS_Ball5.ScaleMesh 1, 1, 1
FR_Ball5.SetRotation FR_Ball5, 0, -Sin5, 0, 0.02

MS_Ball6.ScaleMesh 1, 1, 1
FR_Ball6.SetRotation FR_Ball6, 0, -Sin5, 0, 0.03

MS_Ball7.ScaleMesh 1, 1, 1
FR_Ball7.SetRotation FR_Ball7, 0, -Sin5, 0, 0.04
End Sub
Private Sub RM_KeyDown(KeyCode As Integer, Shift As Integer)
QuitFlag = True
End Sub

Private Sub Timer1_Timer()
RM.Update ' Keeps our scene nice and updated
' Note:  This program doesn't work without this timer to keep everything updated
End Sub

Private Sub Timer10_Timer()

If switchcount10 = 0 Then
Set MS_BackStar(21) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(21) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(21).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(21).SetTexture TX_BackStar(21)
FR_BackStar(21).SetPosition FR_BS1(21), -4, 3, 1
'change size
MS_BackStar(21).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(21).AddVisual MS_BackStar(21)

Set MS_BackStar(22) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(22) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(22).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(22).SetTexture TX_BackStar(22)
FR_BackStar(22).SetPosition FR_BS1(22), 0, 2, 0
'change size
MS_BackStar(22).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(22).AddVisual MS_BackStar(22)

Set MS_BackStar(23) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\green.bmp"
Set TX_BackStar(23) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(23).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(23).SetTexture TX_BackStar(23)
FR_BackStar(23).SetPosition FR_BS1(23), 0, -2, 0
'change size
MS_BackStar(23).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(23).AddVisual MS_BackStar(23)


Set MS_BackStar(24) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(24) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(24).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(24).SetTexture TX_BackStar(24)
FR_BackStar(24).SetPosition FR_BS1(24), 3, 1, -1
'change size
MS_BackStar(24).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(24).AddVisual MS_BackStar(24)

Else

Set MS_BackStar(21) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(21) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(21).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(21).SetTexture TX_BackStar(21)
FR_BackStar(21).SetPosition FR_BS1(21), -4, 3, 1
'change size
MS_BackStar(21).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(21).AddVisual MS_BackStar(21)

Set MS_BackStar(22) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(22) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(22).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(22).SetTexture TX_BackStar(22)
FR_BackStar(22).SetPosition FR_BS1(22), 0, 2, 0
'change size
MS_BackStar(22).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(22).AddVisual MS_BackStar(22)

Set MS_BackStar(23) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\white.bmp"
Set TX_BackStar(23) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(23).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(23).SetTexture TX_BackStar(23)
FR_BackStar(23).SetPosition FR_BS1(23), 0, -2, 0
'change size
MS_BackStar(23).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(23).AddVisual MS_BackStar(23)


Set MS_BackStar(24) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(24) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(24).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(24).SetTexture TX_BackStar(24)
FR_BackStar(24).SetPosition FR_BS1(24), 3, 1, -1
'change size
MS_BackStar(24).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(24).AddVisual MS_BackStar(24)

End If

switchcount10 = switchcount10 + 1
If switchcount10 = 2 Then switchcount10 = 0

Timer10.Enabled = False
Timer3.Enabled = True

End Sub

Private Sub Timer2_Timer()
If QuitFlag = True Then
   x = ShowCursor(True)
  Unload Me
End If
End Sub

Private Sub Timer3_Timer()
If switchcount = 0 Then
Set MS_BackStar(6) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(6) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(6).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(6).SetTexture TX_BackStar(6)
FR_BackStar(6).SetPosition FR_BS1(6), 2, -4, 4
'change size
MS_BackStar(6).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(6).AddVisual MS_BackStar(6)

Set MS_BackStar(1) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\green.bmp"
Set TX_BackStar(1) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(1).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(1).SetTexture TX_BackStar(1)
FR_BackStar(1).SetPosition FR_BS1(1), 2, 0, 5
'change size
MS_BackStar(1).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(1).AddVisual MS_BackStar(1)

Set MS_BackStar(2) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\white.bmp"
Set TX_BackStar(2) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(2).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(2).SetTexture TX_BackStar(2)
FR_BackStar(2).SetPosition FR_BS1(2), 2, 0, 3
'change size
MS_BackStar(2).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(2).AddVisual MS_BackStar(2)

Else

Set MS_BackStar(6) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(6) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(6).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(6).SetTexture TX_BackStar(6)
FR_BackStar(6).SetPosition FR_BS1(6), 2, -4, 4
'change size
MS_BackStar(6).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(6).AddVisual MS_BackStar(6)

Set MS_BackStar(1) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\pink.bmp"
Set TX_BackStar(1) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(1).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(1).SetTexture TX_BackStar(1)
FR_BackStar(1).SetPosition FR_BS1(1), 2, 0, 5
'change size
MS_BackStar(1).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(1).AddVisual MS_BackStar(1)

Set MS_BackStar(2) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(2) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(2).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(2).SetTexture TX_BackStar(2)
FR_BackStar(2).SetPosition FR_BS1(2), 2, 0, 3
'change size
MS_BackStar(2).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(2).AddVisual MS_BackStar(2)
End If
switchcount = switchcount + 1
If switchcount = 2 Then switchcount = 0
Timer3.Enabled = False
Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
DoEvents
If switchcount4 = 0 Then
Set MS_BackStar(3) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\green.bmp"
Set TX_BackStar(3) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(3).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(3).SetTexture TX_BackStar(3)
FR_BackStar(3).SetPosition FR_BS1(3), 2, 3, 3
'change size
MS_BackStar(3).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(3).AddVisual MS_BackStar(3)

Set MS_BackStar(4) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\white.bmp"
Set TX_BackStar(4) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(4).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(4).SetTexture TX_BackStar(4)
FR_BackStar(4).SetPosition FR_BS1(4), 2, -2, 3
'change size
MS_BackStar(4).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(4).AddVisual MS_BackStar(4)

Set MS_BackStar(5) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\green.bmp"
Set TX_BackStar(5) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(5).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(5).SetTexture TX_BackStar(5)
FR_BackStar(5).SetPosition FR_BS1(5), 1, -3, 4
'change size
MS_BackStar(5).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(5).AddVisual MS_BackStar(5)
Else
Set MS_BackStar(3) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(3) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(3).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(3).SetTexture TX_BackStar(3)
FR_BackStar(3).SetPosition FR_BS1(3), 2, 3, 3
'change size
MS_BackStar(3).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(3).AddVisual MS_BackStar(3)

Set MS_BackStar(4) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\pink.bmp"
Set TX_BackStar(4) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(4).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(4).SetTexture TX_BackStar(4)
FR_BackStar(4).SetPosition FR_BS1(4), 2, -2, 3
'change size
MS_BackStar(4).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(4).AddVisual MS_BackStar(4)

Set MS_BackStar(5) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(5) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(5).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(5).SetTexture TX_BackStar(5)
FR_BackStar(5).SetPosition FR_BS1(5), 1, -3, 4
'change size
MS_BackStar(5).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(5).AddVisual MS_BackStar(5)
switchcount4 = switchcount4 + 1
If switchcount4 = 2 Then switchcount4 = 0
End If

Timer4.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Timer5_Timer()
DoEvents
If switchcount5 = 0 Then
Set MS_BackStar(7) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(7) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(7).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(7).SetTexture TX_BackStar(7)
FR_BackStar(7).SetPosition FR_BS1(7), 3, -3, 3
'change size
MS_BackStar(7).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(7).AddVisual MS_BackStar(7)


Set MS_BackStar(8) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(8) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(8).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(8).SetTexture TX_BackStar(8)
FR_BackStar(8).SetPosition FR_BS1(8), 4, -4, 4
'change size
MS_BackStar(8).ScaleMesh 0.03, 0.03, 0.03
FR_BackStar(8).AddVisual MS_BackStar(8)


Set MS_BackStar(9) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\white.bmp"
Set TX_BackStar(9) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(9).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(9).SetTexture TX_BackStar(9)
FR_BackStar(9).SetPosition FR_BS1(9), 4, -2, -1
'change size
MS_BackStar(9).ScaleMesh 0.04, 0.04, 0.04
FR_BackStar(9).AddVisual MS_BackStar(9)
Else
Set MS_BackStar(7) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\green.bmp"
Set TX_BackStar(7) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(7).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(7).SetTexture TX_BackStar(7)
FR_BackStar(7).SetPosition FR_BS1(7), 3, -3, 3
'change size
MS_BackStar(7).ScaleMesh 0.05, 0.05, 0.05
FR_BackStar(7).AddVisual MS_BackStar(7)


Set MS_BackStar(8) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\white.bmp"
Set TX_BackStar(8) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(8).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(8).SetTexture TX_BackStar(8)
FR_BackStar(8).SetPosition FR_BS1(8), 4, -4, 4
'change size
MS_BackStar(8).ScaleMesh 0.03, 0.03, 0.03
FR_BackStar(8).AddVisual MS_BackStar(8)


Set MS_BackStar(9) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(9) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(9).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(9).SetTexture TX_BackStar(9)
FR_BackStar(9).SetPosition FR_BS1(9), 4, -2, -1
'change size
MS_BackStar(9).ScaleMesh 0.04, 0.04, 0.04
FR_BackStar(9).AddVisual MS_BackStar(9)
End If

switchcount5 = switchcount5 + 1
If switchcount5 = 2 Then switchcount5 = 0
Timer5.Enabled = False
Timer6.Enabled = True
End Sub

Private Sub Timer6_Timer()
DoEvents
If switchcount6 = 0 Then
Set MS_BackStar(10) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(10) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(10).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(10).SetTexture TX_BackStar(10)
FR_BackStar(10).SetPosition FR_BS1(10), 4, -3, -2
'change size
MS_BackStar(10).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(10).AddVisual MS_BackStar(10)

Set MS_BackStar(11) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\green.bmp"
Set TX_BackStar(11) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(11).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(11).SetTexture TX_BackStar(11)
FR_BackStar(11).SetPosition FR_BS1(11), -2, -2, -2
'change size
MS_BackStar(11).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(11).AddVisual MS_BackStar(11)


Set MS_BackStar(12) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(12) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(12).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(12).SetTexture TX_BackStar(12)
FR_BackStar(12).SetPosition FR_BS1(12), -3, -1, -2
'change size
MS_BackStar(12).ScaleMesh 0.03, 0.03, 0.03
FR_BackStar(12).AddVisual MS_BackStar(12)

Else

Set MS_BackStar(10) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(10) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(10).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(10).SetTexture TX_BackStar(10)
FR_BackStar(10).SetPosition FR_BS1(10), 4, -3, -2
'change size
MS_BackStar(10).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(10).AddVisual MS_BackStar(10)

Set MS_BackStar(11) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(11) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(11).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(11).SetTexture TX_BackStar(11)
FR_BackStar(11).SetPosition FR_BS1(11), -2, -2, -2
'change size
MS_BackStar(11).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(11).AddVisual MS_BackStar(11)

Set MS_BackStar(12) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\white.bmp"
Set TX_BackStar(12) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(12).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(12).SetTexture TX_BackStar(12)
FR_BackStar(12).SetPosition FR_BS1(12), -3, -1, -2
'change size
MS_BackStar(12).ScaleMesh 0.03, 0.03, 0.03
FR_BackStar(12).AddVisual MS_BackStar(12)
End If

switchcount6 = switchcount6 + 1
If swichcount6 = 2 Then switchcount6 = 0
Timer6.Enabled = False
Timer7.Enabled = True


End Sub

Private Sub Timer7_Timer()
DoEvents
If switchcount7 = 0 Then
Set MS_BackStar(13) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(13) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(13).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(13).SetTexture TX_BackStar(13)
FR_BackStar(13).SetPosition FR_BS1(13), -2, 0, -2
'change size
MS_BackStar(13).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(13).AddVisual MS_BackStar(13)

Set MS_BackStar(14) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\white.bmp"
Set TX_BackStar(14) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(14).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(14).SetTexture TX_BackStar(14)
FR_BackStar(14).SetPosition FR_BS1(14), -3, 1, -2
'change size
MS_BackStar(14).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(14).AddVisual MS_BackStar(14)

Set MS_BackStar(15) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(15) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(15).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(15).SetTexture TX_BackStar(15)
FR_BackStar(15).SetPosition FR_BS1(15), -4, 3, -4
'change size
MS_BackStar(15).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(15).AddVisual MS_BackStar(15)

Else

Set MS_BackStar(13) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(13) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(13).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(13).SetTexture TX_BackStar(13)
FR_BackStar(13).SetPosition FR_BS1(13), -2, 0, -2
'change size
MS_BackStar(13).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(13).AddVisual MS_BackStar(13)

Set MS_BackStar(14) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(14) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(14).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(14).SetTexture TX_BackStar(14)
FR_BackStar(14).SetPosition FR_BS1(14), -3, 1, -2
'change size
MS_BackStar(14).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(14).AddVisual MS_BackStar(14)

Set MS_BackStar(15) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(15) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(15).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(15).SetTexture TX_BackStar(15)
FR_BackStar(15).SetPosition FR_BS1(15), -4, 3, -4
'change size
MS_BackStar(15).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(15).AddVisual MS_BackStar(15)
End If

switchcount7 = switchcount7 + 1
If switchcount7 = 2 Then switchcount7 = 0
Timer7.Enabled = False
Timer8.Enabled = True

End Sub

Private Sub Timer8_Timer()
DoEvents
If switchcount8 = 0 Then
 Set MS_BackStar(16) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(16) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(16).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(16).SetTexture TX_BackStar(16)
FR_BackStar(16).SetPosition FR_BS1(16), -1, 1, 6
'change size
MS_BackStar(16).ScaleMesh 0.04, 0.04, 0.04
FR_BackStar(16).AddVisual MS_BackStar(16)

Set MS_BackStar(17) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(17) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(17).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(17).SetTexture TX_BackStar(17)
FR_BackStar(17).SetPosition FR_BS1(17), -1, 3, -3
'change size
MS_BackStar(17).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(17).AddVisual MS_BackStar(17)

Set MS_BackStar(18) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(18) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(18).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(18).SetTexture TX_BackStar(18)
FR_BackStar(18).SetPosition FR_BS1(18), -2, 3, 1
'change size
MS_BackStar(18).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(18).AddVisual MS_BackStar(18)

Else

Set MS_BackStar(16) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\green.bmp"
Set TX_BackStar(16) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(16).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(16).SetTexture TX_BackStar(16)
FR_BackStar(16).SetPosition FR_BS1(16), -1, 1, 6
'change size
MS_BackStar(16).ScaleMesh 0.04, 0.04, 0.04
FR_BackStar(16).AddVisual MS_BackStar(16)

Set MS_BackStar(17) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(17) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(17).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(17).SetTexture TX_BackStar(17)
FR_BackStar(17).SetPosition FR_BS1(17), -1, 3, -3
'change size
MS_BackStar(17).ScaleMesh 0.01, 0.01, 0.01
FR_BackStar(17).AddVisual MS_BackStar(17)

Set MS_BackStar(18) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(18) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(18).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(18).SetTexture TX_BackStar(18)
FR_BackStar(18).SetPosition FR_BS1(18), -2, 3, 1
'change size
MS_BackStar(18).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(18).AddVisual MS_BackStar(18)
End If

switchcount8 = switchcount8 + 1
If switchcount8 = 2 Then switchcount8 = 0

Timer8.Enabled = False
Timer9.Enabled = True

End Sub

Private Sub Timer9_Timer()
DoEvents
If switchcount9 = 0 Then

Set MS_BackStar(19) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\purp.bmp"
Set TX_BackStar(19) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(19).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(19).SetTexture TX_BackStar(19)
FR_BackStar(19).SetPosition FR_BS1(19), 4, 4, 3
'change size
MS_BackStar(19).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(19).AddVisual MS_BackStar(19)

Set MS_BackStar(20) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\aqua.bmp"
Set TX_BackStar(20) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(20).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(20).SetTexture TX_BackStar(20)
FR_BackStar(20).SetPosition FR_BS1(20), -4, 3, -4
'change size
MS_BackStar(20).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(20).AddVisual MS_BackStar(20)

Set MS_BackStar(21) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(21) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(21).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(21).SetTexture TX_BackStar(21)
FR_BackStar(21).SetPosition FR_BS1(21), -4, 3, 1
'change size
MS_BackStar(21).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(21).AddVisual MS_BackStar(21)

Else

Set MS_BackStar(19) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\red.bmp"
Set TX_BackStar(19) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(19).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(19).SetTexture TX_BackStar(19)
FR_BackStar(19).SetPosition FR_BS1(19), 4, 4, 3
'change size
MS_BackStar(19).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(19).AddVisual MS_BackStar(19)

Set MS_BackStar(20) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\green.bmp"
Set TX_BackStar(20) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(20).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(20).SetTexture TX_BackStar(20)
FR_BackStar(20).SetPosition FR_BS1(20), -4, 3, -4
'change size
MS_BackStar(20).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(20).AddVisual MS_BackStar(20)

Set MS_BackStar(21) = RM.D3DRM.CreateMeshBuilder()
texture1 = App.Path & "\pink.bmp"
Set TX_BackStar(21) = RM.D3DRM.LoadTexture(texture1)
MS_BackStar(21).LoadFromFile App.Path & "\SPHERE.X", 0, 0, Nothing, Nothing
MS_BackStar(21).SetTexture TX_BackStar(21)
FR_BackStar(21).SetPosition FR_BS1(21), -4, 3, 1
'change size
MS_BackStar(21).ScaleMesh 0.02, 0.02, 0.02
FR_BackStar(21).AddVisual MS_BackStar(21)
End If

switchcount9 = switchcount9 + 1
If switchcount9 = 2 Then switchcount9 = 0

Timer9.Enabled = False
Timer10.Enabled = True
End Sub
