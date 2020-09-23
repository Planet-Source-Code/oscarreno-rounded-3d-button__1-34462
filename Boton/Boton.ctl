VERSION 5.00
Begin VB.UserControl Boton 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   855
   ScaleWidth      =   2955
   ToolboxBitmap   =   "Boton.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   360
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   1800
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   120
      Top             =   120
      Width           =   250
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Botón"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "Boton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim sh As Long ' ScaleHeight
Dim sw As Long ' ScaleWidth
Type Colores
    Red As Long
    Green As Long
    Blue As Long
End Type
Dim HasFocus As Boolean
Dim IsDown As Boolean
Dim MouseIn As Boolean
Dim OriginalForeColor As Long
Dim OriginalBackColor As Long

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
'Default Property Values:
Const m_def_Roundness3D = 20
Const m_def_CreateRegion = True
Const m_def_Light = 45
'Const m_def_Default = False
'Const m_def_Cancel = False
Const m_def_RoundSize = 20
'Property Variables:
Dim m_Roundness3D As Long
Dim m_CreateRegion As Boolean
Dim m_Light As Long
'Dim m_Default As Boolean
'Dim m_Cancel As Boolean
Dim m_RoundSize As Long


Private Sub PaintGradient(Optional inFocus As Boolean, Optional inverse As Boolean)
    ' Three user-type variables to store separated RGB values of a color
    Dim ColorRGB1 As Colores
    Dim ColorRGB2 As Colores
    Dim ColorRGB3 As Colores
    ' The number of steps of the cycle
    Dim Pasos As Long
    
    ' working variables
    Dim Red As Double
    Dim Green As Double
    Dim Blue As Double
    Dim i As Long
    Dim start As Long
    Dim finish As Long
    Dim forSteps As Long
    Dim back As Long
    Dim porc As Double
    Dim porc2 As Double
    Dim Y As Double
    
    ' increment variables
    Dim incR1 As Double
    Dim incG1 As Double
    Dim incB1 As Double
    Dim incR2 As Double
    Dim incG2 As Double
    Dim incB2 As Double
    
    Dim osh As Long ' Original ScaleHeight
    Dim osw As Long ' Original ScaleWidth
    
    ' Coordinates of focus square
    Dim x1 As Long
    Dim x2 As Long
    Dim y1 As Long
    Dim y2 As Long
    
    'This started with a Joshua Foster's idea on PSC, but I changed it a lot!!
    
    'First assign the steps the cycle will do.
    ' I put 1/10th of the SH
    Pasos = UserControl.ScaleHeight / 10
    
    
    ' Store ocx's original scales
    osh = UserControl.ScaleHeight
    osw = UserControl.ScaleWidth
    
    ' If the control's so small, we better exit
    If osh < 2 Or osw < 2 Then Exit Sub
    
    ' assign new scales
    UserControl.ScaleHeight = Pasos
    UserControl.ScaleWidth = 1
    UserControl.DrawStyle = vbSolid
    
    ' Then we assign the separated RGB values to the middle color
    back = OriginalBackColor
    ColorRGB2 = Long2RGB(CorrectColor(back))
    
    If inFocus Or HasFocus Then
        ' If the inFocus parameter is on, or the control has the focus, then
        ' let's light the control a bit more
        ColorRGB2.Red = ColorRGB2.Red + (Light \ 2)
        ColorRGB2.Green = ColorRGB2.Green + (Light \ 2)
        ColorRGB2.Blue = ColorRGB2.Blue + (Light \ 2)
    End If
    
    ' In base on that colors, let's make the top color a bit more light
    ColorRGB1.Blue = NoNegative(ColorRGB2.Blue + Light)
    ColorRGB1.Green = NoNegative(ColorRGB2.Green + Light)
    ColorRGB1.Red = NoNegative(ColorRGB2.Red + Light)
    ' And we do the same with the bottom color, only darker
    ColorRGB3.Blue = NoNegative(ColorRGB2.Blue - Light)
    ColorRGB3.Green = NoNegative(ColorRGB2.Green - Light)
    ColorRGB3.Red = NoNegative(ColorRGB2.Red - Light)
    
    ' If button's not enabled then make darker
    If Not UserControl.Enabled Then
        ColorRGB1.Blue = ColorRGB2.Blue
        ColorRGB1.Green = ColorRGB2.Green
        ColorRGB1.Red = ColorRGB2.Red
        ColorRGB2.Blue = ColorRGB3.Blue
        ColorRGB2.Green = ColorRGB3.Green
        ColorRGB2.Red = ColorRGB3.Red
        i = 50 ' The color of label is lightly brighter
        Red = NoNegative(ColorRGB2.Red + i)
        Green = NoNegative(ColorRGB2.Green + i)
        Blue = NoNegative(ColorRGB2.Blue + i)
        Label.ForeColor = RGB(Red, Green, Blue)
    End If
    
    ' If the inverse parameter is on (when a button is pressed)
    If inverse Then
        start = Pasos
        finish = 0
        forSteps = -1
    Else
        start = 0
        finish = Pasos
        forSteps = 1
    End If
    
    'Calculate the increment factor of each RGB color
    porc = Roundness3D / 100
    
    If porc <= 0 Then porc = 0.01
    incR1 = (ColorRGB2.Red - ColorRGB1.Red) / (Pasos * porc)
    incG1 = (ColorRGB2.Green - ColorRGB1.Green) / (Pasos * porc)
    incB1 = (ColorRGB2.Blue - ColorRGB1.Blue) / (Pasos * porc)
    incR2 = (ColorRGB3.Red - ColorRGB2.Red) / (Pasos * porc)
    incG2 = (ColorRGB3.Green - ColorRGB2.Green) / (Pasos * porc)
    incB2 = (ColorRGB3.Blue - ColorRGB2.Blue) / (Pasos * porc)
    
    ' Assign the first color to work variables
    Red = ColorRGB1.Red
    Green = ColorRGB1.Green
    Blue = ColorRGB1.Blue
    
    ' Let's PAINT!
    For i = start To finish Step forSteps
        ' draw colored lines
        UserControl.Line (0, i)-(1, i), RGB(Red, Green, Blue)
        ' calculate the increment

        If i < (Pasos * porc) Then
          Red = NoNegative(Red + incR1)
          Green = NoNegative(Green + incG1)
          Blue = NoNegative(Blue + incB1)
        ElseIf i > (Pasos * (1 - porc)) Then ' now we are at the bottom
          Red = NoNegative(Red + incR2)
          Green = NoNegative(Green + incG2)
          Blue = NoNegative(Blue + incB2)
        End If
        
    Next i
    
    
    
    ''' VERTICAL BORDERS
    
    
    
    ' Shall we draw some vertical lines at the borders?... YES PLEASE
    porc2 = 5
    Pasos = osw / 10
    'UserControl.ScaleHeight = osh
    UserControl.ScaleWidth = Pasos
    incR1 = (ColorRGB2.Red - ColorRGB1.Red) / porc2
    incG1 = (ColorRGB2.Green - ColorRGB1.Green) / porc2
    incB1 = (ColorRGB2.Blue - ColorRGB1.Blue) / porc2
    incR2 = (ColorRGB3.Red - ColorRGB2.Red) / porc2
    incG2 = (ColorRGB3.Green - ColorRGB2.Green) / porc2
    incB2 = (ColorRGB3.Blue - ColorRGB2.Blue) / porc2

    ' Assign the first color to work variables again...
    Red = ColorRGB2.Red
    Green = ColorRGB2.Green
    Blue = ColorRGB2.Blue
    
    If inverse Then
        start = Pasos
        finish = 0
        forSteps = -1
    Else
        start = 0
        finish = Pasos
        forSteps = 1
    End If
    
    Y = 0
    ' Let's PAINT!
    For i = start To finish Step forSteps
        If i <= porc Then
            Red = NoNegative(Red + incR1)
            Green = NoNegative(Green + incG1)
            Blue = NoNegative(Blue + incB1)
            Y = Y + ((ScaleHeight * porc) / Pasos)
            UserControl.Line (i, Y)-(i, ScaleHeight - Y), RGB(Red, Green, Blue)
        ElseIf i > (ScaleWidth - porc2) Then ' now we are at the bottom
            Red = NoNegative(Red + incR2)
            Green = NoNegative(Green + incG2)
            Blue = NoNegative(Blue + incB2)
            Y = Y + (ScaleHeight * porc) / Pasos
            UserControl.Line (i, Y)-(i, ScaleHeight - Y), RGB(Red, Green, Blue)
        Else
            Red = ColorRGB2.Red
            Green = ColorRGB2.Green
            Blue = ColorRGB2.Blue
        End If
          
    Next i

    
    ' return to original scale
    UserControl.ScaleHeight = osh
    UserControl.ScaleWidth = osw

End Sub

Private Function NoNegative(i As Integer) As Integer
    NoNegative = IIf(i < 0, 0, i)
End Function


Private Sub Label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Pic_KeyDown(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub Pic_KeyPress(KeyAscii As Integer)
    Call UserControl_KeyPress(KeyAscii)
End Sub

Private Sub Pic_KeyUp(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyUp(KeyCode, Shift)
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub


Private Sub Timer1_Timer()
    If Not IsHot(hWnd) Then
        PaintGradient True, True
        PaintGradient HasFocus, IsDown
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If UserControl.Enabled Then
        Call UserControl_KeyDown(KeyAscii, 0)
        Call UserControl_KeyUp(KeyAscii, 0)
    End If
End Sub

Private Sub UserControl_EnterFocus()
    HasFocus = True
    PaintGradient True, False
End Sub

Private Sub UserControl_ExitFocus()
    HasFocus = False
    PaintGradient
End Sub

Private Sub UserControl_GotFocus()
    HasFocus = True
    PaintGradient
End Sub

Private Sub UserControl_LostFocus()
    HasFocus = False
    PaintGradient
End Sub

Private Sub UserControl_Resize()
    
    RaiseEvent Resize

    CreaNuevaRegion
    
End Sub

Private Sub Acomodate()
    Dim plus As Long
    ' If the button is down, then put the controls a bit more down
    plus = IIf(IsDown, 15, 0)
    'If plus = 0 Then Stop
    ' Acommodate the controls
    Label.Top = ((UserControl.Height - Label.Height) \ 2) + plus
    
    If Pic.Picture <> 0 Then
        Pic.Top = ((UserControl.Height - Pic.Height) \ 2) + plus
        If Label.Caption <> "" Then
            Pic.Left = ((UserControl.Width - Pic.Width - Label.Width - 50) \ 2) + plus
        Else
            Pic.Left = ((UserControl.Width - Pic.Width) \ 2) + plus
        End If
        Label.Left = Pic.Left + Pic.Width + 50 + plus
    Else
        Label.Left = ((UserControl.Width - Label.Width) \ 2) + plus
    End If

End Sub

Private Sub UserControl_Show()
    CreaNuevaRegion
End Sub


' Converts a Color(Long) to RGB values
Private Function Long2RGB(ByVal LongRGB As Long) As Colores
    Long2RGB.Red = LongRGB And 255
    Long2RGB.Green = (LongRGB \ 256) And 255
    Long2RGB.Blue = (LongRGB \ 65536) And 255
End Function


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label,Label,-1,Caption
Public Property Get Caption() As String
    Caption = Label.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label.Caption() = New_Caption
    CreaNuevaRegion
    PropertyChanged "Caption"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    If New_Enabled Then Label.ForeColor = OriginalForeColor
    Pic.Enabled = New_Enabled
    CreaNuevaRegion
    PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label,Label,-1,Font
Public Property Get Font() As Font
    Set Font = Label.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label.Font = New_Font
    CreaNuevaRegion
    PaintGradient
    PropertyChanged "Font"
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    HasFocus = True
    If (KeyCode = vbKeySpace) Or _
       (KeyCode = vbKeyEscape And Extender.Cancel) Or _
       (KeyCode = vbKeyReturn And Extender.Default) Then
        
            IsDown = True
            HasFocus = True
            PaintGradient False, True
            Acomodate
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeySpace) Or _
       (KeyCode = vbKeyEscape) Or _
       (KeyCode = vbKeyReturn) Then
            IsDown = False
            HasFocus = True
            PaintGradient
            Acomodate
            RaiseEvent Click
    End If
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If IsDown Then
        IsDown = False
    End If
    If Button = vbLeftButton Then
        HasFocus = True
        Acomodate
        PaintGradient False, True
        IsDown = True
    End If
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'DoEvents
    If Button = vbLeftButton Then
            If IsHot(UserControl.hWnd) Then
                HasFocus = True
                IsDown = True
            Else
                IsDown = False
                HasFocus = False
            End If
            PaintGradient HasFocus, IsDown
            Acomodate
    ElseIf Button = 0 And IsHot(hWnd) Then
        'Beep
        If Not Timer1.Enabled Then
            PaintGradient True, False
            Timer1.Enabled = True
        End If
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        PaintGradient
        IsDown = False
        Acomodate
        'If IsHot(UserControl.hWnd) Or IsHot(Pic.hWnd) Then RaiseEvent Click
        If IsHot(UserControl.hWnd) Then RaiseEvent Click
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Pic,Pic,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = Pic.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Pic.Picture = New_Picture
    If New_Picture Is Nothing Then
        Pic.Visible = False
    Else
        Pic.Visible = True
    End If
    CreaNuevaRegion
    PropertyChanged "Picture"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    UserControl.Refresh
End Sub
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=8,0,0,11842740
'Public Property Get Color() As Long
'    Color = m_Color
'End Property
'
'Public Property Let Color(ByVal New_Color As Long)
'    m_Color = New_Color
'    PropertyChanged "Color"
'End Property

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
'    m_Color = m_def_Color
    m_RoundSize = m_def_RoundSize
'    m_Default = m_def_Default
'    m_Cancel = m_def_Cancel
    OriginalBackColor = UserControl.BackColor
    UserControl.BackColor = CorrectColor(UserControl.BackColor)
    OriginalForeColor = UserControl.ForeColor
    Label.ForeColor = CorrectColor(UserControl.ForeColor)
    m_Light = m_def_Light
    m_CreateRegion = m_def_CreateRegion
    m_Roundness3D = m_def_Roundness3D
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label.Caption = PropBag.ReadProperty("Caption", "Botón")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Label.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    OriginalBackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    Label.ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
    m_RoundSize = PropBag.ReadProperty("RoundSize", m_def_RoundSize)
    m_Light = PropBag.ReadProperty("Light", m_def_Light)
    Label.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    OriginalForeColor = Label.ForeColor
    m_CreateRegion = PropBag.ReadProperty("CreateRegion", m_def_CreateRegion)
    m_Roundness3D = PropBag.ReadProperty("Roundness3D", m_def_Roundness3D)
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", Label.Caption, "Label")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Label.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", OriginalBackColor, vbButtonFace)
    Call PropBag.WriteProperty("ForeColor", Label.ForeColor, vbButtonText)
    Call PropBag.WriteProperty("RoundSize", m_RoundSize, m_def_RoundSize)
    Call PropBag.WriteProperty("Light", m_Light, m_def_Light)
    Call PropBag.WriteProperty("ToolTipText", Label.ToolTipText, "")
    Call PropBag.WriteProperty("CreateRegion", m_CreateRegion, m_def_CreateRegion)
    Call PropBag.WriteProperty("Roundness3D", m_Roundness3D, m_def_Roundness3D)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    OriginalBackColor = New_BackColor
    UserControl.BackColor() = CorrectColor(New_BackColor)
    CreaNuevaRegion
    PropertyChanged "BackColor"
End Property

Private Function CorrectColor(color As Long) As Long
    If color < -1 Then  ' If Is a System color
         color = GetSysColor(color And &HFF&)
    End If
    CorrectColor = color

End Function

Private Function ConvertToSysColor(ByVal lColor As Long) As Long
' THIS ROUTINE WAS SUBMITED in PSC BY Rocky Clark's Color Coder 3.0
'Find a system color that matches lColor

Dim lIdx As Long
Dim sHex As String

    If lColor < 0 Then
        'Already a system color
        ConvertToSysColor = lColor
    Else
        For lIdx = 0 To 24
            If GetSysColor(lIdx) = lColor Then
                'Found a match
                sHex = Hex$(lIdx)
                If Len(sHex) < 2 Then
                    sHex = "0" & sHex
                End If
                ConvertToSysColor = Val("&H800000" & sHex)
                Exit For
            End If
        Next
        If lIdx > 24 Then
            'Didn't find a match
            ConvertToSysColor = -1
        End If
    End If
    
End Function



'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label,Label,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Label.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    OriginalForeColor = New_ForeColor
    Label.ForeColor() = CorrectColor(New_ForeColor)
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,20
Public Property Get RoundSize() As Long
    RoundSize = m_RoundSize
End Property

Public Property Let RoundSize(ByVal New_RoundSize As Long)
    m_RoundSize = New_RoundSize
    CreaNuevaRegion
    PropertyChanged "RoundSize"
End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=0,0,0,false
'Public Property Get Default() As Boolean
'    Default = m_Default
'End Property
'
'Public Property Let Default(ByVal New_Default As Boolean)
'    m_Default = New_Default
'    PropertyChanged "Default"
'End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=0,0,0,false
'Public Property Get Cancel() As Boolean
'    Cancel = m_Cancel
'End Property
'
'Public Property Let Cancel(ByVal New_Cancel As Boolean)
'    m_Cancel = New_Cancel
'    PropertyChanged "Cancel"
'End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,25
Public Property Get Light() As Long
    Light = m_Light
End Property

Public Property Let Light(ByVal New_Light As Long)
    m_Light = New_Light
    CreaNuevaRegion
    PropertyChanged "Light"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label,Label,-1,ToolTipText
Public Property Get ToolTipText() As String
    ToolTipText = Label.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Label.ToolTipText() = New_ToolTipText
    Pic.ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Function BackColorOriginal() As Long
    BackColorOriginal = OriginalBackColor
End Function

Public Sub CreaNuevaRegion()
    Dim hRgn&
    ' Thanks to Mick Doherty for this precious hint
    picMask.Width = Width
    picMask.Height = Height
    picMask.Cls
    
    If CreateRegion Then
        hRgn& = CreateRoundRectRgn(0, 0, Width \ 15, Height \ 15, RoundSize, RoundSize)
    Else
        hRgn& = CreateRectRgn(0, 0, Width \ 15, Height \ 15)
    End If
        
    FillRgn picMask.hdc, hRgn&, GetStockObject(BLACKBRUSH)
    MaskPicture = picMask.Image
    ' Paint the Gradient
    PaintGradient

    ' Colocate controls
    Acomodate


End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,true
Public Property Get CreateRegion() As Boolean
    CreateRegion = m_CreateRegion
End Property

Public Property Let CreateRegion(ByVal New_CreateRegion As Boolean)
    m_CreateRegion = New_CreateRegion
    CreaNuevaRegion
    PropertyChanged "CreateRegion"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,20
Public Property Get Roundness3D() As Long
    Roundness3D = m_Roundness3D
End Property

Public Property Let Roundness3D(ByVal New_Roundness3D As Long)
    m_Roundness3D = New_Roundness3D
    CreaNuevaRegion
    PropertyChanged "Roundness3D"
End Property



