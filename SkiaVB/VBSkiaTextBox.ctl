VERSION 5.00
Begin VB.UserControl VBSkiaTextBox 
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ScaleHeight     =   2640
   ScaleWidth      =   4680
End
Attribute VB_Name = "VBSkiaTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Eventos públicos
Public Event Change()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private Declare Function GetTickCount Lib "kernel32" () As Long

' Componentes internos
Private m_SkiaTextBox As SkiaTextBox
Private WithEvents m_PictureBox As PictureBox
Attribute m_PictureBox.VB_VarHelpID = -1
Private WithEvents m_Timer As Timer
Attribute m_Timer.VB_VarHelpID = -1

' Estado interno
Private m_Initialized As Boolean
Private m_HasFocus As Boolean
Private m_CursorBlinkState As Boolean
Private m_LastClickTime As Long
Private m_LastClickX As Single
Private m_LastClickY As Single

' Propriedades
Private m_Text As String
Private m_PlaceholderText As String
Private m_MaxLength As Integer
Private m_ReadOnly As Boolean
Private m_Multiline As Boolean
Private m_TextAlign As Integer
Private m_BackColor As Long
Private m_ForeColor As Long
Private m_PlaceholderColor As Long
Private m_BorderColor As Long
Private m_FocusBorderColor As Long
Private m_SelectionColor As Long
Private m_BorderWidth As Single
Private m_CornerRadius As Single
Private m_FontName As String
Private m_FontSize As Single
Private m_FontBold As Boolean
Private m_Enabled As Boolean
Private m_PaddingLeft As Single
Private m_PaddingTop As Single
Private m_PaddingRight As Single
Private m_PaddingBottom As Single

Private Sub UserControl_Initialize()
    InitializeDefaults
    CreateComponents
End Sub

Private Sub InitializeDefaults()
    On Error GoTo ErrorHandler
    
    m_Text = ""
    m_PlaceholderText = "Digite aqui..."
    m_MaxLength = 0
    m_ReadOnly = False
    m_Multiline = False
    m_TextAlign = 0 ' Left
    m_BackColor = &HFFFFFFFF ' Branco
    m_ForeColor = &H33333333 ' Cinza escuro
    m_PlaceholderColor = &H999999 ' Cinza médio
    m_BorderColor = &HCCCCCC ' Cinza claro
    m_FocusBorderColor = &H4285F4 ' Azul Google
    m_SelectionColor = &H4285F4
    m_BorderWidth = 1
    m_CornerRadius = 4
    m_FontName = "Segoe UI"
    m_FontSize = 14
    m_FontBold = False
    m_Enabled = True
    m_PaddingLeft = 8
    m_PaddingTop = 6
    m_PaddingRight = 8
    m_PaddingBottom = 6
    m_HasFocus = False
    m_CursorBlinkState = True
    m_Initialized = False
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in InitializeDefaults"
End Sub

Private Sub CreateComponents()
    On Error GoTo ErrorHandler
    
    ' Criar instância do SkiaTextBox
    Set m_SkiaTextBox = New SkiaTextBox
    
    ' Configurar UserControl
    UserControl.BackColor = &HF0F0F0
    UserControl.ScaleMode = vbTwips
    
    ' Criar PictureBox interno
    Set m_PictureBox = UserControl.Controls.Add("VB.PictureBox", "InternalPictureBox")
    
    If Not m_PictureBox Is Nothing Then
        With m_PictureBox
            .Left = 0
            .Top = 0
            .Width = UserControl.Width
            .Height = UserControl.Height
            .BackColor = vbWhite
            .BorderStyle = 0
            .AutoRedraw = False
            .Visible = True
            .ScaleMode = vbTwips
        End With
    End If
    
    ' Criar Timer para cursor piscante
    Set m_Timer = UserControl.Controls.Add("VB.Timer", "CursorTimer")
    If Not m_Timer Is Nothing Then
        m_Timer.Interval = 500 ' 500ms
        m_Timer.Enabled = False
    End If
    
    m_Initialized = True
    RefreshTextBox
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao criar componentes: " & Err.Number & " - " & Err.Description
    m_Initialized = False
End Sub

Private Sub UserControl_Show()
    If Not m_Initialized Then
        CreateComponents
    End If
    RefreshTextBox
End Sub

Private Sub UserControl_Resize()
    If m_Initialized And Not m_PictureBox Is Nothing Then
        With m_PictureBox
            .Left = 0
            .Top = 0
            .Width = UserControl.Width
            .Height = UserControl.Height
        End With
        RefreshTextBox
    End If
End Sub

' Propriedades públicas
Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal newText As String)
    m_Text = newText
    PropertyChanged "Text"
    If m_Initialized Then RefreshTextBox
    RaiseEvent Change
End Property

Public Property Get PlaceholderText() As String
    PlaceholderText = m_PlaceholderText
End Property

Public Property Let PlaceholderText(ByVal newText As String)
    m_PlaceholderText = newText
    PropertyChanged "PlaceholderText"
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get MaxLength() As Integer
    MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(ByVal NewLength As Integer)
    m_MaxLength = NewLength
    PropertyChanged "MaxLength"
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = m_ReadOnly
End Property

Public Property Let ReadOnly(ByVal NewReadOnly As Boolean)
    m_ReadOnly = NewReadOnly
    PropertyChanged "ReadOnly"
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get Multiline() As Boolean
    Multiline = m_Multiline
End Property

Public Property Let Multiline(ByVal NewMultiline As Boolean)
    m_Multiline = NewMultiline
    PropertyChanged "Multiline"
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get TextAlign() As Integer
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal NewAlign As Integer)
    m_TextAlign = NewAlign
    PropertyChanged "TextAlign"
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get BackColor() As Long
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewColor As Long)
    m_BackColor = NewColor
    PropertyChanged "BackColor"
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get ForeColor() As Long
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal NewColor As Long)
    m_ForeColor = NewColor
    PropertyChanged "ForeColor"
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get BorderColor() As Long
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewColor As Long)
    m_BorderColor = NewColor
    PropertyChanged "BorderColor"
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get FontName() As String
    FontName = m_FontName
End Property

Public Property Let FontName(ByVal NewFont As String)
    m_FontName = NewFont
    PropertyChanged "FontName"
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal NewSize As Single)
    m_FontSize = NewSize
    PropertyChanged "FontSize"
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
    m_Enabled = NewEnabled
    PropertyChanged "Enabled"
    UserControl.Enabled = NewEnabled
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get CornerRadius() As Single
    CornerRadius = m_CornerRadius
End Property

Public Property Let CornerRadius(ByVal NewRadius As Single)
    m_CornerRadius = NewRadius
    PropertyChanged "CornerRadius"
    If m_Initialized Then RefreshTextBox
End Property

Public Property Get BorderWidth() As Single
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewWidth As Single)
    m_BorderWidth = NewWidth
    PropertyChanged "BorderWidth"
    If m_Initialized Then RefreshTextBox
End Property

' Métodos públicos
Public Sub SelectAll()
    If Not m_SkiaTextBox Is Nothing Then
        m_SkiaTextBox.SelectAll
        RefreshTextBox
    End If
End Sub

Public Sub ClearSelection()
    If Not m_SkiaTextBox Is Nothing Then
        m_SkiaTextBox.ClearSelection
        RefreshTextBox
    End If
End Sub

Public Function GetSelectedText() As String
    If Not m_SkiaTextBox Is Nothing Then
        GetSelectedText = m_SkiaTextBox.GetSelectedText()
    Else
        GetSelectedText = ""
    End If
End Function

Public Sub SetSelection(Start As Integer, Length As Integer)
    If Not m_SkiaTextBox Is Nothing Then
        m_SkiaTextBox.SetSelection Start, Length
        RefreshTextBox
    End If
End Sub

Public Sub SetFocus()
    If m_Enabled And Not m_ReadOnly Then
        m_HasFocus = True
        If Not m_Timer Is Nothing Then
            m_Timer.Enabled = True
        End If
        RefreshTextBox
    End If
End Sub

Public Sub KillFocus()
    m_HasFocus = False
    If Not m_Timer Is Nothing Then
        m_Timer.Enabled = False
    End If
    RefreshTextBox
End Sub

' Eventos do PictureBox
Private Sub m_PictureBox_Click()
    If m_Enabled Then
        SetFocus
        RaiseEvent Click
    End If
End Sub

Private Sub m_PictureBox_DblClick()
    If m_Enabled Then
        SelectAll
        RaiseEvent DblClick
    End If
End Sub

Private Sub m_PictureBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled And Button = vbLeftButton Then
        SetFocus
        
        ' Posicionar cursor baseado no clique
        If Not m_SkiaTextBox Is Nothing Then
            Dim pixelX As Single, pixelY As Single
            pixelX = X / Screen.TwipsPerPixelX
            pixelY = Y / Screen.TwipsPerPixelY
            
            Dim cursorPos As Integer
            cursorPos = m_SkiaTextBox.GetCursorPositionFromPoint(pixelX, pixelY)
            m_SkiaTextBox.CursorPosition = cursorPos
            m_SkiaTextBox.ClearSelection
            RefreshTextBox
        End If
        
        ' Detectar duplo clique
        Dim currentTime As Long
        currentTime = GetTickCount()
        
        If (currentTime - m_LastClickTime) < 300 And _
           Abs(X - m_LastClickX) < 50 And Abs(Y - m_LastClickY) < 50 Then
            SelectAll
        End If
        
        m_LastClickTime = currentTime
        m_LastClickX = X
        m_LastClickY = Y
        
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub m_PictureBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub m_PictureBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub m_PictureBox_KeyPress(KeyAscii As Integer)
    HandleKeyPress KeyAscii
End Sub

Private Sub m_PictureBox_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeyDown KeyCode, Shift
End Sub

Private Sub m_PictureBox_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub m_PictureBox_GotFocus()
    SetFocus
End Sub

Private Sub m_PictureBox_LostFocus()
    KillFocus
End Sub

' Eventos do UserControl
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    HandleKeyPress KeyAscii
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeyDown KeyCode, Shift
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_GotFocus()
    SetFocus
End Sub

Private Sub UserControl_LostFocus()
    KillFocus
End Sub

' Timer para cursor piscante
Private Sub m_Timer_Timer()
    If m_HasFocus Then
        m_CursorBlinkState = Not m_CursorBlinkState
        If Not m_SkiaTextBox Is Nothing Then
            m_SkiaTextBox.ShowCursor = m_CursorBlinkState
            RefreshTextBox
        End If
    End If
End Sub

' Manipulação de teclas
Private Sub HandleKeyPress(KeyAscii As Integer)
    If Not m_Enabled Or m_ReadOnly Or m_SkiaTextBox Is Nothing Then
        Exit Sub
    End If
    
    Select Case KeyAscii
        Case vbKeyBack ' Backspace
            HandleBackspace
            KeyAscii = 0
            
        Case vbKeyReturn ' Enter
            If Not m_Multiline Then
                KeyAscii = 0
            Else
                InsertCharacter (Chr$(13))
                KeyAscii = 0
            End If
            
        Case 32 To 126 ' Caracteres imprimíveis
            InsertCharacter (Chr$(KeyAscii))
            KeyAscii = 0
            
        Case Else
            ' Outros caracteres especiais podem ser tratados aqui
    End Select
    
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub HandleKeyDown(KeyCode As Integer, Shift As Integer)
    If Not m_Enabled Or m_SkiaTextBox Is Nothing Then
        Exit Sub
    End If
    
    Select Case KeyCode
        Case vbKeyLeft
            HandleLeftArrow Shift
            
        Case vbKeyRight
            HandleRightArrow Shift
            
        Case vbKeyHome
            HandleHome Shift
            
        Case vbKeyEnd
            HandleEnd Shift
            
        Case vbKeyDelete
            If Not m_ReadOnly Then
                HandleDelete
            End If
            
        Case 65 ' Ctrl+A
            If Shift And vbCtrlMask Then
                SelectAll
            End If
            
        Case 67 ' Ctrl+C
            If Shift And vbCtrlMask Then
                HandleCopy
            End If
            
        Case 86 ' Ctrl+V
            If (Shift And vbCtrlMask) And Not m_ReadOnly Then
                HandlePaste
            End If
            
        Case 88 ' Ctrl+X
            If (Shift And vbCtrlMask) And Not m_ReadOnly Then
                HandleCut
            End If
    End Select
    
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

' Manipuladores de teclas específicas
Private Sub HandleBackspace()
    If m_SkiaTextBox.SelectionLength > 0 Then
        DeleteSelectedText
    ElseIf m_SkiaTextBox.CursorPosition > 0 Then
        m_Text = Left$(m_Text, m_SkiaTextBox.CursorPosition - 1) & Mid$(m_Text, m_SkiaTextBox.CursorPosition + 1)
        m_SkiaTextBox.Text = m_Text
        m_SkiaTextBox.CursorPosition = m_SkiaTextBox.CursorPosition - 1
        RefreshTextBox
        RaiseEvent Change
    End If
End Sub

Private Sub HandleDelete()
    If m_SkiaTextBox.SelectionLength > 0 Then
        DeleteSelectedText
    ElseIf m_SkiaTextBox.CursorPosition < Len(m_Text) Then
        m_Text = Left$(m_Text, m_SkiaTextBox.CursorPosition) & Mid$(m_Text, m_SkiaTextBox.CursorPosition + 2)
        m_SkiaTextBox.Text = m_Text
        RefreshTextBox
        RaiseEvent Change
    End If
End Sub

Private Sub HandleLeftArrow(Shift As Integer)
    If Shift And vbShiftMask Then
        ' Shift+Left: Estender seleção para a esquerda
        ExtendSelectionLeft
    Else
        ' Left: Mover cursor para a esquerda
        If m_SkiaTextBox.SelectionLength > 0 Then
            m_SkiaTextBox.CursorPosition = m_SkiaTextBox.SelectionStart
            m_SkiaTextBox.ClearSelection
        ElseIf m_SkiaTextBox.CursorPosition > 0 Then
            m_SkiaTextBox.CursorPosition = m_SkiaTextBox.CursorPosition - 1
        End If
        RefreshTextBox
    End If
End Sub

Private Sub HandleRightArrow(Shift As Integer)
    If Shift And vbShiftMask Then
        ' Shift+Right: Estender seleção para a direita
        ExtendSelectionRight
    Else
        ' Right: Mover cursor para a direita
        If m_SkiaTextBox.SelectionLength > 0 Then
            m_SkiaTextBox.CursorPosition = m_SkiaTextBox.SelectionStart + m_SkiaTextBox.SelectionLength
            m_SkiaTextBox.ClearSelection
        ElseIf m_SkiaTextBox.CursorPosition < Len(m_Text) Then
            m_SkiaTextBox.CursorPosition = m_SkiaTextBox.CursorPosition + 1
        End If
        RefreshTextBox
    End If
End Sub

Private Sub HandleHome(Shift As Integer)
    If Shift And vbShiftMask Then
        ' Shift+Home: Selecionar até o início
        m_SkiaTextBox.SetSelection 0, m_SkiaTextBox.CursorPosition
    Else
        ' Home: Ir para o início
        m_SkiaTextBox.MoveCursorToStart
    End If
    RefreshTextBox
End Sub

Private Sub HandleEnd(Shift As Integer)
    If Shift And vbShiftMask Then
        ' Shift+End: Selecionar até o fim
        m_SkiaTextBox.SetSelection m_SkiaTextBox.CursorPosition, Len(m_Text) - m_SkiaTextBox.CursorPosition
    Else
        ' End: Ir para o fim
        m_SkiaTextBox.MoveCursorToEnd
    End If
    RefreshTextBox
End Sub

Private Sub ExtendSelectionLeft()
    Dim currentPos As Integer
    currentPos = m_SkiaTextBox.CursorPosition
    
    If m_SkiaTextBox.SelectionLength = 0 Then
        ' Iniciar seleção
        If currentPos > 0 Then
            m_SkiaTextBox.SetSelection currentPos - 1, 1
            m_SkiaTextBox.CursorPosition = currentPos - 1
        End If
    Else
        ' Estender seleção existente
        If m_SkiaTextBox.SelectionStart = currentPos Then
            ' Estender para a esquerda
            If m_SkiaTextBox.SelectionStart > 0 Then
                m_SkiaTextBox.SetSelection m_SkiaTextBox.SelectionStart - 1, m_SkiaTextBox.SelectionLength + 1
                m_SkiaTextBox.CursorPosition = m_SkiaTextBox.SelectionStart
            End If
        Else
            ' Reduzir seleção
            m_SkiaTextBox.SetSelection m_SkiaTextBox.SelectionStart, m_SkiaTextBox.SelectionLength - 1
            m_SkiaTextBox.CursorPosition = m_SkiaTextBox.SelectionStart + m_SkiaTextBox.SelectionLength
        End If
    End If
    RefreshTextBox
End Sub

Private Sub ExtendSelectionRight()
    Dim currentPos As Integer
    currentPos = m_SkiaTextBox.CursorPosition
    
    If m_SkiaTextBox.SelectionLength = 0 Then
        ' Iniciar seleção
        If currentPos < Len(m_Text) Then
            m_SkiaTextBox.SetSelection currentPos, 1
            m_SkiaTextBox.CursorPosition = currentPos + 1
        End If
    Else
        ' Estender seleção existente
        If m_SkiaTextBox.SelectionStart + m_SkiaTextBox.SelectionLength = currentPos Then
            ' Estender para a direita
            If currentPos < Len(m_Text) Then
                m_SkiaTextBox.SetSelection m_SkiaTextBox.SelectionStart, m_SkiaTextBox.SelectionLength + 1
                m_SkiaTextBox.CursorPosition = m_SkiaTextBox.SelectionStart + m_SkiaTextBox.SelectionLength
            End If
        Else
            ' Reduzir seleção
            m_SkiaTextBox.SetSelection m_SkiaTextBox.SelectionStart + 1, m_SkiaTextBox.SelectionLength - 1
            m_SkiaTextBox.CursorPosition = m_SkiaTextBox.SelectionStart
        End If
    End If
    RefreshTextBox
End Sub

Private Sub HandleCopy()
    Dim selectedText As String
    selectedText = GetSelectedText()
    If Len(selectedText) > 0 Then
        Clipboard.SetText selectedText
    End If
End Sub

Private Sub HandleCut()
    HandleCopy
    DeleteSelectedText
End Sub

Private Sub HandlePaste()
    Dim clipText As String
    clipText = Clipboard.GetText()
    If Len(clipText) > 0 Then
        InsertText clipText
    End If
End Sub

Private Sub InsertCharacter(char As String)
    InsertText char
End Sub

Private Sub InsertText(newText As String)
    If m_MaxLength > 0 And Len(m_Text) + Len(newText) - m_SkiaTextBox.SelectionLength > m_MaxLength Then
        Exit Sub ' Não inserir se exceder o limite
    End If
    
    ' Deletar texto selecionado primeiro
    If m_SkiaTextBox.SelectionLength > 0 Then
        DeleteSelectedText
    End If
    
    ' Inserir novo texto na posição do cursor
    Dim cursorPos As Integer
    cursorPos = m_SkiaTextBox.CursorPosition
    
    m_Text = Left$(m_Text, cursorPos) & newText & Mid$(m_Text, cursorPos + 1)
    m_SkiaTextBox.Text = m_Text
    m_SkiaTextBox.CursorPosition = cursorPos + Len(newText)
    
    RefreshTextBox
    RaiseEvent Change
End Sub

Private Sub DeleteSelectedText()
    If m_SkiaTextBox.SelectionLength > 0 Then
        Dim selStart As Integer, selEnd As Integer
        selStart = m_SkiaTextBox.SelectionStart
        selEnd = selStart + m_SkiaTextBox.SelectionLength
        
        m_Text = Left$(m_Text, selStart) & Mid$(m_Text, selEnd + 1)
        m_SkiaTextBox.Text = m_Text
        m_SkiaTextBox.CursorPosition = selStart
        m_SkiaTextBox.ClearSelection
        
        RefreshTextBox
        RaiseEvent Change
    End If
End Sub

Private Sub RefreshTextBox()
    If Not m_Initialized Or m_SkiaTextBox Is Nothing Or m_PictureBox Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    With m_SkiaTextBox
        .Text = m_Text
        .PlaceholderText = m_PlaceholderText
        .MaxLength = m_MaxLength
        .ReadOnly = m_ReadOnly
        .Multiline = m_Multiline
        .TextAlign = m_TextAlign
        
        .Width = m_PictureBox.ScaleWidth \ Screen.TwipsPerPixelX
        .Height = m_PictureBox.ScaleHeight \ Screen.TwipsPerPixelY
        
        .BackgroundColor = ConvertColorToARGB(m_BackColor, 255)
        .TextColor = ConvertColorToARGB(m_ForeColor, 255)
        .PlaceholderColor = ConvertColorToARGB(m_PlaceholderColor, 255)
        .BorderColor = ConvertColorToARGB(m_BorderColor, 255)
        .FocusBorderColor = ConvertColorToARGB(m_FocusBorderColor, 255)
        .SelectionColor = ConvertColorToARGB(m_SelectionColor, 255)
        .BorderWidth = m_BorderWidth
        .CornerRadius = m_CornerRadius
        
        .FontFamily = m_FontName
        .FontSize = m_FontSize
        .Bold = m_FontBold
        .Enabled = m_Enabled
        .HasFocus = m_HasFocus
        .ShowCursor = m_CursorBlinkState
        
        .PaddingLeft = m_PaddingLeft
        .PaddingTop = m_PaddingTop
        .PaddingRight = m_PaddingRight
        .PaddingBottom = m_PaddingBottom
    End With
    
    Dim picture As IPictureDisp
    Set picture = m_SkiaTextBox.RenderTextBox()
    
    If Not picture Is Nothing Then
        Set m_PictureBox.picture = picture
        m_PictureBox.Refresh
    Else
        DrawFallbackTextBox "Erro no rendering"
    End If
    
    Exit Sub
    
ErrorHandler:
    DrawFallbackTextBox "Erro: " & Err.Number & " - " & Err.Description
End Sub

Private Sub DrawFallbackTextBox(erro As String)
    If Not m_Initialized Or m_PictureBox Is Nothing Then Exit Sub
    
    On Error Resume Next
    
    With m_PictureBox
        .Cls
        .BackColor = m_BackColor
        .ForeColor = m_ForeColor
        .FontName = m_FontName
        .FontSize = 8
        
        .ToolTipText = erro
        
        Dim displayText As String
        If Len(m_Text) > 0 Then
            displayText = m_Text
        Else
            displayText = m_PlaceholderText
            .ForeColor = m_PlaceholderColor
        End If
        
        .CurrentX = m_PaddingLeft * 15 ' Converter para twips aproximadamente
        .CurrentY = (.ScaleHeight - .textHeight(displayText)) / 2
        
        .Refresh
    End With
End Sub

Private Function ConvertColorToARGB(rgbColor As Long, alpha As Byte) As Long
    ConvertColorToARGB = rgbColor
End Function

' Eventos de propriedades para persistência
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Text = PropBag.ReadProperty("Text", "")
    m_PlaceholderText = PropBag.ReadProperty("PlaceholderText", "Digite aqui...")
    m_MaxLength = PropBag.ReadProperty("MaxLength", 0)
    m_ReadOnly = PropBag.ReadProperty("ReadOnly", False)
    m_Multiline = PropBag.ReadProperty("Multiline", False)
    m_TextAlign = PropBag.ReadProperty("TextAlign", 0)
    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H333333)
    m_PlaceholderColor = PropBag.ReadProperty("PlaceholderColor", &H999999)
    m_BorderColor = PropBag.ReadProperty("BorderColor", &HCCCCCC)
    m_FocusBorderColor = PropBag.ReadProperty("FocusBorderColor", &H4285F4)
    m_SelectionColor = PropBag.ReadProperty("SelectionColor", &H4285F4)
    m_BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    m_CornerRadius = PropBag.ReadProperty("CornerRadius", 4)
    m_FontName = PropBag.ReadProperty("FontName", "Segoe UI")
    m_FontSize = PropBag.ReadProperty("FontSize", 14)
    m_FontBold = PropBag.ReadProperty("FontBold", False)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    
    If m_Initialized Then RefreshTextBox
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Text", m_Text, ""
    PropBag.WriteProperty "PlaceholderText", m_PlaceholderText, "Digite aqui..."
    PropBag.WriteProperty "MaxLength", m_MaxLength, 0
    PropBag.WriteProperty "ReadOnly", m_ReadOnly, False
    PropBag.WriteProperty "Multiline", m_Multiline, False
    PropBag.WriteProperty "TextAlign", m_TextAlign, 0
    PropBag.WriteProperty "BackColor", m_BackColor, &HFFFFFF
    PropBag.WriteProperty "ForeColor", m_ForeColor, &H333333
    PropBag.WriteProperty "PlaceholderColor", m_PlaceholderColor, &H999999
    PropBag.WriteProperty "BorderColor", m_BorderColor, &HCCCCCC
    PropBag.WriteProperty "FocusBorderColor", m_FocusBorderColor, &H4285F4
    PropBag.WriteProperty "SelectionColor", m_SelectionColor, &H4285F4
    PropBag.WriteProperty "BorderWidth", m_BorderWidth, 1
    PropBag.WriteProperty "CornerRadius", m_CornerRadius, 4
    PropBag.WriteProperty "FontName", m_FontName, "Segoe UI"
    PropBag.WriteProperty "FontSize", m_FontSize, 14
    PropBag.WriteProperty "FontBold", m_FontBold, False
    PropBag.WriteProperty "Enabled", m_Enabled, True
End Sub

Private Sub UserControl_Terminate()
    Set m_SkiaTextBox = Nothing
    Set m_PictureBox = Nothing
    Set m_Timer = Nothing
End Sub


