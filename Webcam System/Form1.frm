VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectShow Webcam Minimal Example"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   15795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   15795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   12720
      TabIndex        =   10
      Top             =   5160
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   5160
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox outpic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3960
      Left            =   10320
      ScaleHeight     =   262
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   8
      Top             =   360
      Width           =   4830
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   9960
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmd_connectport 
      Caption         =   "connect port"
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   4680
      Width           =   1815
   End
   Begin VB.PictureBox MSComm1 
      Height          =   480
      Left            =   11160
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   9
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   4680
   End
   Begin VB.TextBox txtTimerInterval 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Text            =   "500"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton amdAutoClick 
      Caption         =   "Auto Click"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.PictureBox picSnapshot 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3960
      Left            =   5040
      ScaleHeight     =   262
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   3
      Top             =   360
      Width           =   4830
   End
   Begin VB.CommandButton cmdSnap 
      Caption         =   "Snap"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   11040
      TabIndex        =   17
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   9120
      TabIndex        =   16
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Snapshot"
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   60
      Width           =   4830
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Preview"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4830
   End
   Begin VB.Image imgPlaceHolder 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3960
      Left            =   60
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4830
   End
   Begin VB.Menu mnuCameras 
      Caption         =   "Cameras"
      Begin VB.Menu mnuCamerasChoice 
         Caption         =   "none"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuCamerasDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCamerasAddNew 
         Caption         =   "Add new camera..."
      End
      Begin VB.Menu mnuCamerasDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCamerasRemove 
         Caption         =   "Remove camera list"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


'Requires a reference to:
'
'   ActiveMovie control type library (quartz.dll).
'

Dim rr As Long
Dim gg As Long
Dim bb As Long
Dim inp As String
Dim xx As Long
Dim yy As Long


Dim pixelval As Long


Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const MASKBORDERLESS = Not (WS_BORDER Or WS_DLGFRAME Or WS_SYSMENU Or WS_THICKFRAME)
Private Const MASKBORDERMIN = Not (WS_DLGFRAME Or WS_SYSMENU Or WS_THICKFRAME)

'FILTER_STATE values, should have been defined in Quartz.dll,
'but another item Microsoft left out.
Private Enum FILTER_STATE
    State_Stopped = 0
    State_Paused = 1
    State_Running = 2
End Enum

Private Const E_FAIL As Long = &H80004005

'These are "scripts" followed by BuildGraph() below to create a
'DirectShow FilterGraph for webcam viewing.
'
'FILTERLIST is incomplete, and must be prepended with the name
'of your webcam's Video Capture Source filter.  Since there may
'be multiples, FILTERLIST begins with "~Capture" which is used
'when BuildGraph() interprets this script to select one having
'a pin named "Capture".
Private Const FILTERLIST As String = _
        "~Capture|" _
      & "AVI Decompressor|" _
      & "Color Space Converter|" _
      & "Video Renderer"
Private Const CONNECTIONLIST As String = _
        "Capture~XForm In|" _
      & "XForm Out~Input|" _
      & "XForm Out~VMR Input0"

Private fgmVidCap As QuartzTypeLib.FilgraphManager 'Not "Is Nothing" means camera is previewing.
Private bv2VidCap As QuartzTypeLib.IBasicVideo2
Private vwVidCap As QuartzTypeLib.IVideoWindow
Private SelectedCamera As Integer '-1 means none selected.
Private InsideWidth As Double
Private AspectRatio As Double

Private Function BuildGraph( _
    ByVal FGM As QuartzTypeLib.FilgraphManager, _
    ByVal Filters As String, _
    ByVal Connections As String) As Integer
    'Returns -1 on success, or FilterIndex when not found, or
    'ConnIndex + 100 when a pin of the connection not found.
    '
    'Filters:
    '
    '   A string with Filter Name values separated by "|" delimiters
    '   and optionally each of these can be followed by one required
    '   Pin Name value separated by a "~" delimiter for use as a tie
    '   breaker when there might be multiple filters with the same
    '   Name value.
    '
    'Connections:
    '
    '   A string with a list of output pins to be connected to
    '   input pins.  Each pin-pair is separated by "|" delimiters
    '   and each pair has out and in pins separated by a "~"
    '   delimiter.  The pin-pairs should be one less than the number
    '   of filters.
    Dim FilterNames() As String
    Dim FilterIndex As Integer
    Dim FilterParts() As String
    Dim FoundFilter As Boolean
    Dim rfiEach As QuartzTypeLib.IRegFilterInfo
    Dim fiFilters() As QuartzTypeLib.IFilterInfo
    Dim Conns() As String
    Dim ConnIndex As Integer
    Dim ConnParts() As String
    Dim piEach As QuartzTypeLib.IPinInfo
    Dim piOut As QuartzTypeLib.IPinInfo
    Dim piIn As QuartzTypeLib.IPinInfo
    
    'Setup for filter script processing.
    FilterNames = Split(UCase$(Filters), "|")
    ReDim fiFilters(UBound(FilterNames))
    
    'Find and add filters.
    For FilterIndex = 0 To UBound(FilterNames)
        FilterParts = Split(FilterNames(FilterIndex), "~")
        For Each rfiEach In FGM.RegFilterCollection
            If UCase$(rfiEach.Name) = FilterParts(0) Then
                rfiEach.Filter fiFilters(FilterIndex)
                If UBound(FilterParts) > 0 Then
                    For Each piEach In fiFilters(FilterIndex).Pins
                        If UCase$(piEach.Name) = FilterParts(1) Then
                            FoundFilter = True
                            Exit For
                        End If
                    Next
                Else
                    FoundFilter = True
                    Exit For
                End If
            End If
        Next
        If FoundFilter Then
            FoundFilter = False
        Else
            BuildGraph = FilterIndex
            Exit Function 'Error result will be 0, 1, etc.
        End If
    Next
    BuildGraph = -1
    
    'Setup for connection script processing.
    Conns = Split(UCase$(Connections), "|")
    FilterIndex = 0
    
    'Find and connect pins.
    For ConnIndex = 0 To UBound(Conns)
        ConnParts = Split(Conns(ConnIndex), "~")
        For Each piEach In fiFilters(FilterIndex).Pins
            If UCase$(piEach.Name) = ConnParts(0) Then
                Set piOut = piEach
                Exit For
            End If
        Next
        For Each piEach In fiFilters(FilterIndex + 1).Pins
            If UCase$(piEach.Name) = ConnParts(1) Then
                Set piIn = piEach
                Exit For
            End If
        Next
        If piOut Is Nothing Or piIn Is Nothing Then
            'Error, missing a pin.
            BuildGraph = ConnIndex + 100 'Error result will be 100, 101, etc.
            Exit Function
        End If
        piOut.ConnectDirect piIn
        FilterIndex = FilterIndex + 1
    Next
End Function

Private Sub DeselectFailedCamera(ByVal Error As Long)
    Dim CameraName As String
    
    With mnuCamerasChoice(SelectedCamera)
        .Checked = False
        CameraName = .Caption
    End With
    SelectedCamera = -1
    SaveSettings
    MsgBox "Selected camera failed, may not be connected:" & vbNewLine _
         & vbNewLine _
         & CameraName & vbNewLine _
         & vbNewLine _
         & "BuildGraph error " & CStr(Error), _
           vbOKOnly Or vbInformation
End Sub

Private Function IsCameraInMenu(ByVal CameraName As String) As Boolean
    Dim C As Integer
    
    If SelectedCamera >= 0 Then
        For C = 0 To mnuCamerasChoice.UBound
            If CameraName = mnuCamerasChoice(C).Caption Then
                IsCameraInMenu = True
                Exit For
            End If
        Next
    End If
End Function

Private Sub LoadSettings()
    Dim F As Integer
    Dim C As Integer
    Dim CameraName As String
    
    SelectedCamera = -1 'None.
    On Error Resume Next
    GetAttr "Settings.txt"
    If Err.Number = 0 Then
        On Error GoTo 0
        F = FreeFile(0)
        Open "Settings.txt" For Input As #F
        Input #F, SelectedCamera
        Do Until EOF(F)
            Input #F, CameraName
            If C > 0 Then Load mnuCamerasChoice(C)
            With mnuCamerasChoice(C)
                .Enabled = True
                .Caption = CameraName
                .Checked = C = SelectedCamera
            End With
            C = C + 1
        Loop
        Close #F
        mnuCamerasRemove.Enabled = True
    End If
End Sub

Private Sub SaveSettings()
    Dim F As Integer
    Dim C As Integer
    
    F = FreeFile(0)
    Open "Settings.txt" For Output As #F
    Write #F, SelectedCamera
    For C = 0 To mnuCamerasChoice.UBound
        Write #F, mnuCamerasChoice(C).Caption
    Next
    Close #F
End Sub

Private Function StartCamera(ByVal CamName As String) As Integer
    'Returns -1 on success, or BuildGraph() error on failures.
    
    Set fgmVidCap = New QuartzTypeLib.FilgraphManager
    'Tack camera name onto FILTERLIST and try to start it.
    StartCamera = BuildGraph(fgmVidCap, CamName & FILTERLIST, CONNECTIONLIST)
    If StartCamera >= 0 Then Exit Function
    
    Set bv2VidCap = fgmVidCap
    With bv2VidCap
        AspectRatio = CDbl(.VideoHeight) / CDbl(.VideoWidth)
    End With
    
    Set vwVidCap = fgmVidCap
    With vwVidCap
        .FullScreenMode = False
        .Left = ScaleX(imgPlaceHolder.Left, ScaleMode, vbPixels)
        .Top = ScaleY(imgPlaceHolder.Top, ScaleMode, vbPixels)
        .Width = ScaleX(InsideWidth, ScaleMode, vbPixels) + 2
        .Height = ScaleY(InsideWidth * AspectRatio, ScaleMode, vbPixels) + 2
        picSnapshot.Height = InsideWidth * AspectRatio + ScaleY(2, vbPixels, ScaleMode)
        imgPlaceHolder.Visible = False
        .WindowStyle = .WindowStyle And MASKBORDERMIN
        .Owner = hWnd
        .Visible = True
    End With
    
    StartCamera = -1
    cmdSnap.Enabled = True
    fgmVidCap.Run
End Function

Private Sub StopCamera()
    Const StopWaitMs As Long = 40
    Dim State As FILTER_STATE
    
    If Not fgmVidCap Is Nothing Then
        With fgmVidCap
            .Stop
            Do
                .GetState StopWaitMs, State
            Loop Until State = State_Stopped Or Err.Number = E_FAIL
        End With
        With vwVidCap
            .Visible = False
            .Owner = 0
        End With
        Set vwVidCap = Nothing
        Set bv2VidCap = Nothing
        Set fgmVidCap = Nothing
    End If
    imgPlaceHolder.Visible = True
    cmdSnap.Enabled = False
End Sub

Private Sub convertToRGB(ByVal lngclr As Long)
Dim tmpval As Long
tmpval = lngclr

'rr = (tmpval And 16711680) \ 65535
'gg = (tmpval And 65535) \ 256
'bb = tmpval And 255


rr = (tmpval And 255)
gg = (tmpval And 65535) \ 256
bb = (tmpval And 16777215) \ 65535

End Sub

Private Sub amdAutoClick_Click()
Timer1.Interval = Val(txtTimerInterval.Text)
Timer1.Enabled = Not Timer1.Enabled
If (Timer1.Enabled = Not Timer1.Enabled) Then
For xx = 1 To picSnapshot.ScaleWidth
    For yy = 1 To picSnapshot.ScaleHeight
        pixelval = picSnapshot.Point(xx, yy)
        convertToRGB (pixelval)
       If (rr > 170 And gg < 110 And bb < 110) Then '  red colour finding
       'If (rr < 250) Then
           'outpic.PSet (xx, yy), RGB(0, 0, 0)
      '  Else
          '  outpic.PSet (xx, yy), RGB(255, 255, 255)
           Debug.Print "RED color"
            If MSComm1.PortOpen = True Then
                
                 Debug.Print "STOP"
                  MSComm1.Output = Chr(115)
            End If
            
            If (rr < 100 And gg > 200 And bb < 110) Then '  red colour finding
       'If (rr < 250) Then
           'outpic.PSet (xx, yy), RGB(0, 0, 0)
      '  Else
          '  outpic.PSet (xx, yy), RGB(255, 255, 255)
           Debug.Print "GREEN color"
            If MSComm1.PortOpen = True Then
                
                 Debug.Print "GO"
                  MSComm1.Output = Chr(102)
            End If
            
'        If minx > xx Then
'        minx = xx
'        End If
'
'        If maxx < xx Then
'        maxx = xx
'        End If
'
'        If miny > yy Then
'        miny = yy
'        End If
'
'        If maxy < yy Then
'        maxy = yy
'        End If
        End If
        
        If (rr < 100 And gg > 120 And bb < 100) Then '  green colour finding
       'If (rr < 250) Then
           'outpic.PSet (xx, yy), RGB(0, 0, 0)
      '  Else
          '  outpic.PSet (xx, yy), RGB(255, 255, 255)
           Debug.Print "GREEN COLOR"
        If MSComm1.PortOpen = True Then
                
                  MSComm1.Output = Chr(102)
            End If
'        If minx > xx Then
'        minx = xx
'        End If
'
'        If maxx < xx Then
'        maxx = xx
'        End If
'
'        If miny > yy Then
'        miny = yy
'        End If
'
'        If maxy < yy Then
'        maxy = yy
'        End If
        End If
        
'        If (rr < 130 And gg < 130 And bb > 140) Then '  blue colour finding
'       'If (rr < 250) Then
'           'outpic.PSet (xx, yy), RGB(0, 0, 0)
'      '  Else
'          '  outpic.PSet (xx, yy), RGB(255, 255, 255)
'
'        If minx > xx Then
'        minx = xx
'        End If
'
'        If maxx < xx Then
'        maxx = xx
'        End If
'
'        If miny > yy Then
'        miny = yy
'        End If
'
'        If maxy < yy Then
'        maxy = yy
'        End If
'        End If
    Next
Next
End If

End Sub

Private Sub cmd_connectport_Click()
MSComm2.Settings = ("115200,n,8,1")
 MSComm2.CommPort = 2
 MSComm2.PortOpen = True
 MSComm2.RThreshold = 1
Debug.Print MSComm2.PortOpen
End Sub

Private Sub cmdSnap_Click()
    Const PauseWaitMs As Long = 16
    Const biSize = 40 'BITMAPINFOHEADER and not BITMAPV4HEADER, etc. but we don't get those.
    Dim State As FILTER_STATE
    Dim Size As Long
    Dim DIB() As Long
    Dim hBitmap As Long
    Dim Pic As StdPicture
    
    With fgmVidCap
        .Pause
        Do
            .GetState PauseWaitMs, State
        Loop Until State = State_Paused Or Err.Number = E_FAIL
        If Err.Number = E_FAIL Then
            MsgBox "Failed to pause webcam preview for snapshot!", _
                   vbOKOnly Or vbExclamation
            Exit Sub
        End If
    
        With bv2VidCap
            'Estimate size.  Correct for 32-bit RGB and generous
            'for anything with fewer bits per pixel, compressed,
            'or palette-ized (we hope).
            Size = biSize + .VideoWidth * .VideoHeight
            ReDim DIB(Size - 1)
            Size = Size * 4 'To bytes.
            .GetCurrentImage Size, DIB(0)
        End With
        
        .Run
    End With
    
    hBitmap = LongDIB2HBitmap(DIB)
    If hBitmap <> 0 Then
        Set Pic = HBitmap2Picture(hBitmap, 0)
        If Not Pic Is Nothing Then
            With picSnapshot
                .AutoRedraw = True
                .PaintPicture Pic, 0, 0, .ScaleWidth, .ScaleHeight
                .AutoRedraw = False
            End With
        End If
        DeleteObject hBitmap
    End If
End Sub

Private Sub Command1_Click()
 Call cmdSnap_Click
 
For xx = 1 To picSnapshot.ScaleWidth
    For yy = 1 To picSnapshot.ScaleHeight
        'pixelval = mypic.Point(xx, yy)
        pixelval = GetPixel(picSnapshot.hDC, xx, yy)
        convertToRGB (pixelval)
        'Label3.Caption = rr
       ' Label4.Caption = gg
       ' Label5.Caption = bb
        
        If (rr > 170 And rr < 220 And gg < 90 And bb > 50 And bb < 90) Then '  red colour finding
       'If (rr < 250) Then
           'outpic.PSet (xx, yy), RGB(0, 0, 0)
      '  Else
          '  outpic.PSet (xx, yy), RGB(255, 255, 255)
          Debug.Print "RED color"
          Label3.Caption = rr
        Label4.Caption = gg
        Label5.Caption = bb
        
            If MSComm2.PortOpen = True Then
                
                ' Debug.Print "STOP"
                  MSComm2.Output = Chr(115)
            End If
        End If
        
         If (rr < 100 And gg > 200 And bb > 110) Then '  green colour finding
       'If (rr < 250) Then
           'outpic.PSet (xx, yy), RGB(0, 0, 0)
      '  Else
          '  outpic.PSet (xx, yy), RGB(255, 255, 255)
          Debug.Print "GREEN color"
          Label3.Caption = rr
        Label4.Caption = gg
        Label5.Caption = bb
            If MSComm2.PortOpen = True Then
                
                 Debug.Print "go"
                 MSComm2.Output = Chr(98)
            End If
        End If
        
        If (rr > 170 And gg < 80 And bb < 50) Then '  blue colour finding
'       'If (rr < 250) Then
'           'outpic.PSet (xx, yy), RGB(0, 0, 0)
'      '  Else
'          '  outpic.PSet (xx, yy), RGB(255, 255, 255)
            Label3.Caption = rr
                 Label4.Caption = gg
                Label5.Caption = bb
           Debug.Print "Ball detected"
            If MSComm2.PortOpen = True Then
'
                Debug.Print "ST"
                 MSComm2.Output = Chr(102)
           End If
        End If
       ' If (rr > 40 And rr < 70 And gg < 60 And bb < 50) Then ' blue colour finding
'       'If (rr < 250) Then
'           'outpic.PSet (xx, yy), RGB(0, 0, 0)
'      '  Else
'          '  outpic.PSet (xx, yy), RGB(255, 255, 255)
            ' Label3.Caption = rr
                ' Label4.Caption = gg
                'Label5.Caption = bb
          ' Debug.Print "yellow"
          ' If MSComm2.PortOpen = True Then
'
               ' Debug.Print "STOP"
                '  MSComm2.Output = Chr(115)
         '  End If
       ' End If
       ' outpic.PSet (xx, yy), RGB((rr + gg + bb) / 3, (rr + gg + bb) / 3, (rr + gg + bb) / 3)
       SetPixelV outpic.hDC, xx, yy, RGB((rr + gg + bb) / 3, (rr + gg + bb) / 3, (rr + gg + bb) / 3)
    
    Next
    DoEvents
Next
 
End Sub

Private Sub Command2_Click()
Timer1.Interval = Val(txtTimerInterval.Text)
Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Form_Load()
    Dim StartResult As Integer
    
    InsideWidth = picSnapshot.Width - ScaleX(2, vbPixels, ScaleMode)
    LoadSettings
    If SelectedCamera >= 0 Then
        StartResult = StartCamera(mnuCamerasChoice(SelectedCamera).Caption)
        If StartResult >= 0 Then DeselectFailedCamera StartResult
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Make sure we break the link, don't want to have a memory
    'leak or hang the camera or something:
    StopCamera
    
    Unload Form2
End Sub

Private Sub mnuCamerasAddNew_Click()
    Dim StartResult As Integer
    
    Form2.Show vbModal, Me
    If Form2.Oked Then
        If IsCameraInMenu(Form2.CameraName) Then
            MsgBox "Camera:" & vbNewLine _
                 & vbNewLine _
                 & Form2.CameraName & vbNewLine _
                 & vbNewLine _
                 & "Is already in the menu."
        Else
            StopCamera
            StartResult = StartCamera(Form2.CameraName)
            If StartResult < 0 Then
                If SelectedCamera >= 0 Then
                    mnuCamerasChoice(SelectedCamera).Checked = False
                End If
                If SelectedCamera >= 0 Then
                    SelectedCamera = mnuCamerasChoice.UBound + 1
                    Load mnuCamerasChoice(SelectedCamera)
                Else
                    SelectedCamera = 0
                End If
                With mnuCamerasChoice(SelectedCamera)
                    .Caption = Form2.CameraName
                    .Checked = True
                    .Enabled = True
                End With
                SaveSettings
                mnuCamerasRemove.Enabled = True
            Else
                MsgBox "This doesn't seems to be a valid webcam:" & vbNewLine _
                     & vbNewLine _
                     & Form2.CameraName & vbNewLine _
                     & vbNewLine _
                     & "BuildGraph error " & CStr(Error), _
                       vbOKOnly Or vbInformation
                'Try to go back to previous camera.
                If SelectedCamera > -1 Then
                    StartResult = StartCamera(mnuCamerasChoice(SelectedCamera).Caption)
                    If StartResult >= 0 Then DeselectFailedCamera StartResult
                End If
            End If
        End If
    End If
End Sub

Private Sub mnuCamerasChoice_Click(Index As Integer)
    Dim StartResult As Integer
    
    If Index <> SelectedCamera Then
        If SelectedCamera >= 0 Then
            StopCamera
            mnuCamerasChoice(SelectedCamera).Checked = False
        End If
        SelectedCamera = Index
        mnuCamerasChoice(SelectedCamera).Checked = True
        StartResult = StartCamera(mnuCamerasChoice(SelectedCamera).Caption)
        If StartResult < 0 Then
            SaveSettings
        Else
            DeselectFailedCamera StartResult
        End If
    End If
End Sub

Private Sub mnuCamerasRemove_Click()
    Dim C As Integer
    
    StopCamera
    SelectedCamera = -1
    With mnuCamerasChoice(0)
        .Caption = "none"
        .Enabled = False
    End With
    For C = mnuCamerasChoice.UBound To 1 Step -1
        Unload mnuCamerasChoice(C)
    Next
    mnuCamerasRemove.Enabled = False
    On Error Resume Next
    Kill "Settings.txt"
End Sub

Private Sub MSComm2_OnComm()
inp = MSComm2.Input
If inp <> " " Then
    Caption = inp
    Debug.Print inp
End If
'End Sub
End Sub

Private Sub Timer1_Timer()
Call cmdSnap_Click
Call Command1_Click
End Sub
