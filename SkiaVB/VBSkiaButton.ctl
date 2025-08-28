VERSION 5.00
Begin VB.UserControl VBSkiaButton 
   BackColor       =   &H00F0F0F0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "VBSkiaButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Click()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseDown()
Public Event MouseUp()

Private m_SkiaButton As SkiaButton
Private WithEvents m_PictureBox As PictureBox
Attribute m_PictureBox.VB_VarHelpID = -1
Private m_IsHovered As Boolean
Private m_IsPressed As Boolean
Private m_Enabled As Boolean
Private m_Initialized As Boolean

Private m_Text As String
Private m_BackColor As Long
Private m_ForeColor As Long
Private m_BorderColor As Long
Private m_BorderWidth As Single
Private m_CornerRadius As Single
Private m_FontName As String
Private m_FontSize As Single
Private m_FontBold As Boolean
Private m_UseGradient As Boolean
Private m_GradientStartColor As Long
Private m_GradientEndColor As Long

Private Sub UserControl_Initialize()
1         InitializeDefaults
2         CreateComponents
End Sub

Private Sub InitializeDefaults()
1         On Error GoTo InitializeDefaults_Error
2
3         m_Text = "SkiaButton"
4         m_BackColor = &H4285F4
5         m_ForeColor = &HFFFFFF
6         m_BorderColor = &H1976D2
7         m_BorderWidth = 2
8         m_CornerRadius = 8
9         m_FontName = "Segoe UI"
10        m_FontSize = 12
11        m_FontBold = False
12        m_Enabled = True
13        m_IsHovered = False
14        m_IsPressed = False
15        m_UseGradient = False
16        m_GradientStartColor = &H6A5ACD
17        m_GradientEndColor = &H4169E1
18        m_Initialized = False
            
19        On Error GoTo 0
20        Exit Sub

InitializeDefaults_Error:

21        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitializeDefaults, line " & Erl & "."

End Sub

Private Sub CreateComponents()
1         On Error GoTo ErrorHandler
          
2         Set m_SkiaButton = New SkiaButton
         
          
3         UserControl.BackColor = &HF0F0F0
4         UserControl.ScaleMode = vbTwips
          
          
5         Set m_PictureBox = UserControl.Controls.Add("VB.PictureBox", "InternalPictureBox")
           
6         If Not m_PictureBox Is Nothing Then
7             With m_PictureBox
8                 .Left = 0
9                 .Top = 0
10                .Width = 2500
11                .Height = 800
12                .BackColor = vbBlack
13                .BorderStyle = 0
14                .AutoRedraw = True
15                .Visible = True
16                .ScaleMode = vbTwips
17            End With
18            m_Initialized = True
19            RefreshButton
20        End If
          
21        Exit Sub
          
ErrorHandler:
22      MsgBox "Erro ao criar componentes: " & Err.Number & " - " & Err.Description & " - Linha" & Erl

23        m_Initialized = False
End Sub

Private Sub UserControl_Show()
1         If Not m_Initialized Then
2             CreateComponents
3         End If
4         RefreshButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
1         m_Text = PropBag.ReadProperty("Text", "SkiaButton")
2         m_BackColor = PropBag.ReadProperty("BackColor", &H4285F4)
3         m_ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
4         m_BorderColor = PropBag.ReadProperty("BorderColor", &H1976D2)
5         m_BorderWidth = PropBag.ReadProperty("BorderWidth", 2)
6         m_CornerRadius = PropBag.ReadProperty("CornerRadius", 8)
7         m_FontName = PropBag.ReadProperty("FontName", "Segoe UI")
8         m_FontSize = PropBag.ReadProperty("FontSize", 12)
9         m_FontBold = PropBag.ReadProperty("FontBold", False)
10        m_Enabled = PropBag.ReadProperty("Enabled", True)
11        m_UseGradient = PropBag.ReadProperty("UseGradient", False)
12        m_GradientStartColor = PropBag.ReadProperty("GradientStartColor", &H6A5ACD)
13        m_GradientEndColor = PropBag.ReadProperty("GradientEndColor", &H4169E1)
          
14        If m_Initialized Then
15            RefreshButton
16        End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
1         PropBag.WriteProperty "Text", m_Text, "SkiaButton"
2         PropBag.WriteProperty "BackColor", m_BackColor, &H4285F4
3         PropBag.WriteProperty "ForeColor", m_ForeColor, &HFFFFFF
4         PropBag.WriteProperty "BorderColor", m_BorderColor, &H1976D2
5         PropBag.WriteProperty "BorderWidth", m_BorderWidth, 2
6         PropBag.WriteProperty "CornerRadius", m_CornerRadius, 8
7         PropBag.WriteProperty "FontName", m_FontName, "Segoe UI"
8         PropBag.WriteProperty "FontSize", m_FontSize, 12
9         PropBag.WriteProperty "FontBold", m_FontBold, False
10        PropBag.WriteProperty "Enabled", m_Enabled, True
11        PropBag.WriteProperty "UseGradient", m_UseGradient, False
12        PropBag.WriteProperty "GradientStartColor", m_GradientStartColor, &H6A5ACD
13        PropBag.WriteProperty "GradientEndColor", m_GradientEndColor, &H4169E1
End Sub

Private Sub UserControl_Resize()
1         If m_Initialized And Not m_PictureBox Is Nothing Then
2             With m_PictureBox
3                 .Left = 0
4                 .Top = 0
5                 .Width = UserControl.Width
6                 .Height = UserControl.Height
7             End With
8             RefreshButton
9         End If
End Sub

Public Property Get Text() As String
1         Text = m_Text
End Property

Public Property Let Text(ByVal NewText As String)
1         m_Text = NewText
2         PropertyChanged "Text"
3         If m_Initialized Then RefreshButton
End Property

Public Property Get BackColor() As Long
1         BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewColor As Long)
1         m_BackColor = NewColor
2         PropertyChanged "BackColor"
3         If m_Initialized Then RefreshButton
End Property

Public Property Get ForeColor() As Long
1         ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal NewColor As Long)
1         m_ForeColor = NewColor
2         PropertyChanged "ForeColor"
3         If m_Initialized Then RefreshButton
End Property

Public Property Get BorderColor() As Long
1         BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewColor As Long)
1         m_BorderColor = NewColor
2         PropertyChanged "BorderColor"
3         If m_Initialized Then RefreshButton
End Property

Public Property Get BorderWidth() As Single
1         BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewWidth As Single)
1         m_BorderWidth = NewWidth
2         PropertyChanged "BorderWidth"
3         If m_Initialized Then RefreshButton
End Property

Public Property Get CornerRadius() As Single
1         CornerRadius = m_CornerRadius
End Property

Public Property Let CornerRadius(ByVal NewRadius As Single)
1         m_CornerRadius = NewRadius
2         PropertyChanged "CornerRadius"
3         If m_Initialized Then RefreshButton
End Property

Public Property Get FontName() As String
1         FontName = m_FontName
End Property

Public Property Let FontName(ByVal NewFont As String)
1         m_FontName = NewFont
2         PropertyChanged "FontName"
3         If m_Initialized Then RefreshButton
End Property

Public Property Get FontSize() As Single
1         FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal NewSize As Single)
1         m_FontSize = NewSize
2         PropertyChanged "FontSize"
3         If m_Initialized Then RefreshButton
End Property

Public Property Get FontBold() As Boolean
1         FontBold = m_FontBold
End Property

Public Property Let FontBold(ByVal NewBold As Boolean)
1         m_FontBold = NewBold
2         PropertyChanged "FontBold"
3         If m_Initialized Then RefreshButton
End Property

Public Property Get Enabled() As Boolean
1         Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
1         m_Enabled = NewEnabled
2         PropertyChanged "Enabled"
3         UserControl.Enabled = NewEnabled
4         If m_Initialized Then RefreshButton
End Property

Public Property Get UseGradient() As Boolean
1         UseGradient = m_UseGradient
End Property

Public Property Let UseGradient(ByVal NewUseGradient As Boolean)
1         m_UseGradient = NewUseGradient
2         PropertyChanged "UseGradient"
3         If m_Initialized Then RefreshButton
End Property

Public Property Get GradientStartColor() As Long
1         GradientStartColor = m_GradientStartColor
End Property

Public Property Let GradientStartColor(ByVal NewColor As Long)
1         m_GradientStartColor = NewColor
2         PropertyChanged "GradientStartColor"
3         If m_Initialized Then RefreshButton
End Property

Public Property Get GradientEndColor() As Long
1         GradientEndColor = m_GradientEndColor
End Property

Public Property Let GradientEndColor(ByVal NewColor As Long)
1         m_GradientEndColor = NewColor
2         PropertyChanged "GradientEndColor"
3         If m_Initialized Then RefreshButton
End Property


Public Sub SetGradient(StartColor As Long, EndColor As Long)
1         m_GradientStartColor = StartColor
2         m_GradientEndColor = EndColor
3         m_UseGradient = True
4         If m_Initialized Then RefreshButton
End Sub

Public Sub SetTextShadow(OffsetX As Single, OffsetY As Single, BlurRadius As Single, ShadowColor As Long)
1         If Not m_SkiaButton Is Nothing Then
2             m_SkiaButton.SetTextShadow OffsetX, OffsetY, BlurRadius, ConvertColorToARGB(ShadowColor, 128)
3             If m_Initialized Then RefreshButton
4         End If
End Sub

Public Sub Refresh()
1         If m_Initialized Then RefreshButton
End Sub

Private Sub RefreshButton()
    If Not m_Initialized Or m_SkiaButton Is Nothing Or m_PictureBox Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler

    With m_SkiaButton
        .Text = m_Text
         
        .Width = m_PictureBox.ScaleWidth \ Screen.TwipsPerPixelX
        .Height = m_PictureBox.ScaleHeight \ Screen.TwipsPerPixelY
  
        .TextColor = ConvertColorToARGB(m_ForeColor, 255)

        .BorderColor = ConvertColorToARGB(m_BorderColor, 255)
        .BorderWidth = m_BorderWidth
        .CornerRadius = m_CornerRadius
       
        .FontFamily = m_FontName
        .FontSize = m_FontSize
        .Bold = m_FontBold
        .Enabled = m_Enabled
        .IsHovered = m_IsHovered
        .IsPressed = m_IsPressed

        If m_UseGradient Then
            .SetGradientBackground ConvertColorToARGB(m_GradientStartColor, 255), ConvertColorToARGB(m_GradientEndColor, 255)
        Else
            .BackgroundColor = ConvertColorToARGB(m_BackColor, 255)
        End If
    End With
    
    Dim picture As IPictureDisp
    Set picture = m_SkiaButton.RenderButton()
    
    If Not picture Is Nothing Then
        Set m_PictureBox.picture = picture
        m_PictureBox.Refresh
    Else
        DrawFallbackButton "Aqui nao foi"
    End If
    
    Exit Sub
    
ErrorHandler:
    DrawFallbackButton "error rendering button: " & Err.Number & "-" & Err.Description
End Sub

Private Sub DrawFallbackButton(Erro As String)
1         If Not m_Initialized Or m_PictureBox Is Nothing Then Exit Sub
          
2         On Error Resume Next
          
3         With m_PictureBox
4             .Cls
5             .BackColor = m_BackColor
6             .ForeColor = m_ForeColor
7             .FontName = m_FontName
8             .FontSize = m_FontSize
9             .FontBold = m_FontBold
              
       
10            .ToolTipText = Erro
   
              Dim textWidth As Single, textHeight As Single
11            textWidth = .textWidth(m_Text)
12            textHeight = .textHeight(m_Text)
              
13            .CurrentX = (.ScaleWidth - textWidth) / 2
14            .CurrentY = (.ScaleHeight - textHeight) / 2
       
15            .Refresh
16        End With
End Sub

Private Function ConvertColorToARGB(rgbColor As Long, alpha As Byte) As Long
          Dim r As Byte, g As Byte, b As Byte
          
1        ConvertColorToARGB = rgbColor
2         Exit Function
3         r = rgbColor And &HFF
4         g = (rgbColor And &HFF00&) \ &H100
5         b = (rgbColor And &HFF0000) \ &H10000
          
6         ConvertColorToARGB = (CLng(alpha) * &H1000000) Or (CLng(b) * &H10000) Or (CLng(g) * &H100) Or CLng(r)
End Function

Private Sub m_PictureBox_Click()
1         If m_Enabled Then
2             RaiseEvent Click
3         End If
End Sub

Private Sub m_PictureBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1         If m_Enabled And Button = vbLeftButton Then
2             m_IsPressed = True
3             RefreshButton
4             RaiseEvent MouseDown
5         End If
End Sub

Private Sub m_PictureBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
1         If m_IsPressed Then
2             m_IsPressed = False
3             RefreshButton
4             RaiseEvent MouseUp
5         End If
End Sub

Private Sub m_PictureBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
1         If Not m_IsHovered And m_Enabled Then
2             m_IsHovered = True
3             RefreshButton
4             RaiseEvent MouseEnter
5         End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
1         If m_IsHovered Then
2             If X < 0 Or Y < 0 Or X > UserControl.Width Or Y > UserControl.Height Then
3                 m_IsHovered = False
4                 RefreshButton
5                 RaiseEvent MouseLeave
6             End If
7         End If
End Sub

Private Sub UserControl_Terminate()
1         Set m_SkiaButton = Nothing
2         Set m_PictureBox = Nothing
End Sub
