VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "UserControl"
      Height          =   435
      Left            =   900
      TabIndex        =   0
      Top             =   5580
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private btnPrimary As SkiaButton
Attribute btnPrimary.VB_VarHelpID = -1
Private btnSecondary As SkiaButton
Attribute btnSecondary.VB_VarHelpID = -1
Private btnSuccess As SkiaButton
Attribute btnSuccess.VB_VarHelpID = -1
Private btnWarning As SkiaButton
Attribute btnWarning.VB_VarHelpID = -1
Private btnDanger As SkiaButton
Attribute btnDanger.VB_VarHelpID = -1
Private btnCustom As SkiaButton
Attribute btnCustom.VB_VarHelpID = -1

Private WithEvents picPrimary As PictureBox
Attribute picPrimary.VB_VarHelpID = -1
Private WithEvents picSecondary As PictureBox
Attribute picSecondary.VB_VarHelpID = -1
Private WithEvents picSuccess As PictureBox
Attribute picSuccess.VB_VarHelpID = -1
Private WithEvents picWarning As PictureBox
Attribute picWarning.VB_VarHelpID = -1
Private WithEvents picDanger As PictureBox
Attribute picDanger.VB_VarHelpID = -1
Private WithEvents picCustom As PictureBox
Attribute picCustom.VB_VarHelpID = -1

Private animTimer As Timer
Private animStep As Integer
Private animatingButton As Integer


Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Form_Load()
    InitializeForm
    CreateButtons
    SetupButtonStyles
    RenderAllButtons
    
End Sub

Private Sub InitializeForm()
    Me.Caption = "Skia Advanced Button Demo - VB6"
    Me.Width = 10000
    Me.Height = 8000
    Me.BackColor = &HF8F9FA
    
    Set animTimer = Me.Controls.Add("VB.Timer", "animTimer")
    animTimer.Interval = 50
    animTimer.Enabled = False
       
    Dim lblTitle As Label
    Set lblTitle = Me.Controls.Add("VB.Label", "lblTitle")
    With lblTitle
        .Caption = "Skia Button Showcase"
        .Font.Size = 16
        .Font.Bold = True
        .Left = 100
        .Top = 200
        .Width = 4000
        .Height = 400
        .BackStyle = 0
        .ForeColor = &H2C3E50
        .Visible = True
    End With
End Sub

Private Sub CreateButtons()
    Set btnPrimary = New SkiaButton
    Set btnSecondary = New SkiaButton
    Set btnSuccess = New SkiaButton
    Set btnWarning = New SkiaButton
    Set btnDanger = New SkiaButton
    Set btnCustom = New SkiaButton
    
    Dim i As Integer
    Dim pics(5) As PictureBox
    Dim picNames As Variant
    picNames = Array("picPrimary", "picSecondary", "picSuccess", "picWarning", "picDanger", "picCustom")
    
    For i = 0 To 5
        Set pics(i) = Me.Controls.Add("VB.PictureBox", picNames(i))
        With pics(i)
            .Left = 500 + (i Mod 3) * 3000
            .Top = 1000 + (i \ 3) * 1200
            .Width = 2500
            .Height = 800
            .BorderStyle = 0
            .AutoRedraw = True
            .Visible = True
        End With
    Next i
    
    Set picPrimary = pics(0)
    Set picSecondary = pics(1)
    Set picSuccess = pics(2)
    Set picWarning = pics(3)
    Set picDanger = pics(4)
    Set picCustom = pics(5)
End Sub

Private Sub SetupButtonStyles()
 
    
    With btnPrimary
        .Text = "Primary"
        .Width = 180
        .Height = 50
        .SetGradientBackground &HFF007BFF, &HFF0056B3
        .TextColor = &HFFFFFFFF
        .BorderWidth = 0
        .CornerRadius = 6
        .FontFamily = "Segoe UI"
        .FontSize = 14
        .Bold = False
        .SetTextShadow 0, 1, 2, &H40000000
    End With
    

    With btnSecondary
        .Text = "Secondary"
        .Width = 180
        .Height = 50
        .BackgroundColor = &HFF6C757D
        .TextColor = &HFFFFFFFF
        .BorderColor = &HFF545B62
        .BorderWidth = 1
        .CornerRadius = 6
        .FontFamily = "Segoe UI"
        .FontSize = 14
        .Bold = False
    End With
    

    With btnSuccess
        .Text = "Success"
        .Width = 180
        .Height = 50
        .SetGradientBackground &HFF28A745, &HFF1E7E34
        .TextColor = &HFFFFFFFF
        .BorderWidth = 0
        .CornerRadius = 6
        .FontFamily = "Segoe UI"
        .FontSize = 14
        .Bold = False
        .SetTextShadow 0, 1, 2, &H40000000
    End With
    

    With btnWarning
        .Text = "Warning"
        .Width = 180
        .Height = 50
        .SetGradientBackground &HFFFFC107, &HFFE0A800
        .TextColor = &HFF212529
        .BorderWidth = 0
        .CornerRadius = 6
        .FontFamily = "Segoe UI"
        .FontSize = 14
        .Bold = False
        .SetTextShadow 0, 1, 2, &H40FFFFFF
    End With
    

    With btnDanger
        .Text = "Danger"
        .Width = 180
        .Height = 50
        .SetGradientBackground &HFFDC3545, &HFFC82333
        .TextColor = &HFFFFFFFF
        .BorderWidth = 0
        .CornerRadius = 6
        .FontFamily = "Segoe UI"
        .FontSize = 14
        .Bold = False
        .SetTextShadow 0, 1, 2, &H40000000
    End With
    
 
    With btnCustom
        .Text = "Custom Style"
        .Width = 180
        .Height = 50
        .SetGradientBackground &HFF6F42C1, &HFF563D7C
        .TextColor = &HFFFFFFFF
        .BorderColor = &HFFAB7FE4
        .BorderWidth = 2
        .CornerRadius = 25
        .FontFamily = "Arial"
        .FontSize = 13
        .Bold = True
        .SetTextShadow 0, 2, 4, &H60000000
    End With
End Sub

Private Sub RenderAllButtons()
    RenderButton btnPrimary, picPrimary
    RenderButton btnSecondary, picSecondary
    RenderButton btnSuccess, picSuccess
    RenderButton btnWarning, picWarning
    RenderButton btnDanger, picDanger
    RenderButton btnCustom, picCustom
End Sub

Private Sub RenderButton(btn As SkiaButton, pic As PictureBox)
    On Error GoTo ErrorHandler
    
    Dim picture As IPictureDisp
   Set picture = btn.RenderButton()
    
    If Not picture Is Nothing Then
        Set pic.picture = picture
    End If
    
    Exit Sub
    
ErrorHandler:

    pic.Cls
    pic.Print "render error"
End Sub

' Eventos de clique dos botões
Private Sub picPrimary_Click()
    AnimateButtonClick 1
    MsgBox "Primary button clicked!", vbInformation, "Button Event"
End Sub

Private Sub picSecondary_Click()
    AnimateButtonClick 2
    MsgBox "Secondary button clicked!", vbInformation, "Button Event"
End Sub

Private Sub picSuccess_Click()
    AnimateButtonClick 3
    MsgBox "Success button clicked!", vbInformation, "Button Event"
End Sub

Private Sub picWarning_Click()
    AnimateButtonClick 4
    MsgBox "Warning button clicked!", vbExclamation, "Button Event"
End Sub

Private Sub picDanger_Click()
    AnimateButtonClick 5
    MsgBox "Danger button clicked!", vbCritical, "Button Event"
End Sub

Private Sub picCustom_Click()
    AnimateButtonClick 6
    ShowCustomDialog
End Sub

Private Sub AnimateButtonClick(buttonIndex As Integer)
    animatingButton = buttonIndex
    animStep = 0
    animTimer.Enabled = True
    
    Select Case buttonIndex
        Case 1:
            btnPrimary.IsPressed = True
            RenderButton btnPrimary, picPrimary
        Case 2:
            btnSecondary.IsPressed = True
            RenderButton btnSecondary, picSecondary
        Case 3:
            btnSuccess.IsPressed = True
            RenderButton btnSuccess, picSuccess
        Case 4:
            btnWarning.IsPressed = True
            RenderButton btnWarning, picWarning
        Case 5:
            btnDanger.IsPressed = True
            RenderButton btnDanger, picDanger
        Case 6:
            btnCustom.IsPressed = True
            RenderButton btnCustom, picCustom
    End Select
End Sub

Private Sub animTimer_Timer()
    animStep = animStep + 1
    
    If animStep >= 4 Then
        animTimer.Enabled = False
        
        Select Case animatingButton
            Case 1:
                btnPrimary.IsPressed = False
                RenderButton btnPrimary, picPrimary
            Case 2:
                btnSecondary.IsPressed = False
                RenderButton btnSecondary, picSecondary
            Case 3:
                btnSuccess.IsPressed = False
                RenderButton btnSuccess, picSuccess
            Case 4:
                btnWarning.IsPressed = False
                RenderButton btnWarning, picWarning
            Case 5:
                btnDanger.IsPressed = False
                RenderButton btnDanger, picDanger
            Case 6:
                btnCustom.IsPressed = False
                RenderButton btnCustom, picCustom
        End Select
    End If
End Sub

Private Sub picPrimary_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not btnPrimary.IsHovered Then
        btnPrimary.IsHovered = True
        RenderButton btnPrimary, picPrimary
    End If
End Sub

Private Sub picSecondary_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not btnSecondary.IsHovered Then
        btnSecondary.IsHovered = True
        RenderButton btnSecondary, picSecondary
    End If
End Sub

Private Sub picSuccess_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not btnSuccess.IsHovered Then
        btnSuccess.IsHovered = True
        RenderButton btnSuccess, picSuccess
    End If
End Sub

Private Sub picWarning_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not btnWarning.IsHovered Then
        btnWarning.IsHovered = True
        RenderButton btnWarning, picWarning
    End If
End Sub

Private Sub picDanger_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not btnDanger.IsHovered Then
        btnDanger.IsHovered = True
        RenderButton btnDanger, picDanger
    End If
End Sub

Private Sub picCustom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not btnCustom.IsHovered Then
        btnCustom.IsHovered = True
        RenderButton btnCustom, picCustom
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If btnPrimary.IsHovered Then
        btnPrimary.IsHovered = False
        RenderButton btnPrimary, picPrimary
    End If
    If btnSecondary.IsHovered Then
        btnSecondary.IsHovered = False
        RenderButton btnSecondary, picSecondary
    End If
    If btnSuccess.IsHovered Then
        btnSuccess.IsHovered = False
        RenderButton btnSuccess, picSuccess
    End If
    If btnWarning.IsHovered Then
        btnWarning.IsHovered = False
        RenderButton btnWarning, picWarning
    End If
    If btnDanger.IsHovered Then
        btnDanger.IsHovered = False
        RenderButton btnDanger, picDanger
    End If
    If btnCustom.IsHovered Then
        btnCustom.IsHovered = False
        RenderButton btnCustom, picCustom
    End If
End Sub

Private Sub ShowCustomDialog()
    Dim response As VbMsgBoxResult
    response = MsgBox("create another custom button?", _
                     vbQuestion + vbYesNo, "Custom Button")
    
    If response = vbYes Then
        CreateNewCustomButton
    End If
End Sub

Private Sub CreateNewCustomButton()
    
    Dim newBtn As New SkiaButton
    Dim newPic As PictureBox
    
    Set newPic = Me.Controls.Add("VB.PictureBox", "newCustomPic")
    With newPic
        .Left = 4000
        .Top = 4000
        .Width = 3000
        .Height = 1000
        .BorderStyle = 0
        .AutoRedraw = True
        .Visible = True
    End With
    
    
    With newBtn
        .Text = "Novo Botão!"
        .Width = 220
        .Height = 60
        .SetGradientBackground &HFFFF6B6B, &HFFFF8E53
        .TextColor = &HFFFFFFFF
        .BorderColor = &HFFFF9F43
        .BorderWidth = 3
        .CornerRadius = 30
        .FontFamily = "Impact"
        .FontSize = 16
        .Bold = True
        .SetTextShadow 2, 2, 4, &H80000000
    End With
    
 
    Dim picture As Object
    Set picture = newBtn.RenderButton()
    If Not picture Is Nothing Then
        Set newPic.picture = picture
    End If
    
 
    MsgBox "Created!"
End Sub
