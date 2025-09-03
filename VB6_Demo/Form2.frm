VERSION 5.00
Begin VB.Form Form2 
   Caption         =   " "
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6915
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private WithEvents skiaBtn As VBSkiaButton
Attribute skiaBtn.VB_VarHelpID = -1
Private WithEvents skiaBtn2 As VBSkiaButton
Attribute skiaBtn2.VB_VarHelpID = -1
Private WithEvents skiaBtn3 As VBSkiaButton
Attribute skiaBtn3.VB_VarHelpID = -1
Private WithEvents skiaText As VBSkiaTextBox
Attribute skiaText.VB_VarHelpID = -1

Private Sub Form_Load()
    Me.Caption = "Teste SkiaButtonControl"
    Me.Width = 8000
    Me.Height = 6000
    Me.BackColor = &HF0F0F0
    
    CreateUserControlButtons
End Sub

Private Sub CreateUserControlButtons()
    On Error GoTo ErrorHandler
    
    Set skiaBtn = Me.Controls.Add("SkiaVB.VBSkiaButton", "MySkiaButton1")
    
    If Not skiaBtn Is Nothing Then
        With skiaBtn
            .Left = 500
            .Top = 500
            .Width = 2500
            .Height = 800
            .Text = "Primary Button"
            .UseGradient = True
            .GradientStartColor = &HFF007BFF
            .GradientEndColor = &HFF0056B3
            .ForeColor = &HFFFFFFFF
            .CornerRadius = 8
            .FontSize = 12
            .FontBold = False
            .Visible = True
        End With
    End If

    Set skiaBtn2 = Me.Controls.Add("SkiaVB.VBSkiaButton", "MySkiaButton2")
    
    If Not skiaBtn2 Is Nothing Then
        With skiaBtn2
            .Left = 500
            .Top = 1500
            .Width = 2500
            .Height = 800
            .Text = "Secondary Button"
            
            .BackColor = &HFF6C757D
            .ForeColor = &HFFFFFFFF
            .BorderColor = &HFF545B62
            .UseGradient = False
            .CornerRadius = 15
            .FontSize = 14
            .FontBold = True
            .Visible = True
        End With
    End If
    
    Set skiaBtn3 = Me.Controls.Add("SkiaVB.VBSkiaButton", "MySkiaButton3")
    
    If Not skiaBtn3 Is Nothing Then
        With skiaBtn3
            .Left = 500
            .Top = 2500
            .Width = 2500
            .Height = 800
            .Text = "Danger Style"
            .GradientStartColor = &HFFDC3545
            .GradientEndColor = &HFFC82333
            .UseGradient = True
            .ForeColor = &HFFFFFFFF
            .BorderColor = &HE91E63
            .BorderWidth = 6
            .CornerRadius = 20
            .FontSize = 13
            .FontBold = True
            .Visible = True
        End With
    End If

    CreateLabels
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating controls: " & Err.Number & "  " & Err.Description & "Line:" & Erl & vbCrLf, vbExclamation, "Error"
End Sub

Private Sub CreateLabels()
    Dim lbl1 As Label, lbl2 As Label, lbl3 As Label
    
    ' Label 1
    Set lbl1 = Me.Controls.Add("VB.Label", "lbl1")
    With lbl1
        .Caption = "Basic Button with solid color"
        .Left = 3200
        .Top = 700
        .Width = 3000
        .Height = 300
        .BackStyle = 0
        .ForeColor = &H333333
        .Visible = True
    End With
    
    ' Label 2
    Set lbl2 = Me.Controls.Add("VB.Label", "lbl2")
    With lbl2
        .Caption = "Gradient Button"
        .Left = 3200
        .Top = 1700
        .Width = 3000
        .Height = 300
        .BackStyle = 0
        .ForeColor = &H333333
        .Visible = True
    End With
    
    Set lbl3 = Me.Controls.Add("VB.Label", "lbl3")
    With lbl3
        .Caption = "Custom Button with border and shadow"
        .Left = 3200
        .Top = 2700
        .Width = 3000
        .Height = 300
        .BackStyle = 0
        .ForeColor = &H333333
        .Visible = True
    End With
End Sub

Private Sub skiaBtn_Click()
    MsgBox "Primary Button clicked!", vbInformation, "Evento Click"
End Sub

Private Sub skiaBtn_MouseEnter()
    Me.Caption = "Mouse hover Primary Button"
End Sub

Private Sub skiaBtn_MouseLeave()
    Me.Caption = "Test SkiaButtonControl"
End Sub

Private Sub skiaBtn2_Click()
    MsgBox "Gradient Button clicked!", vbInformation, "Click Event"
    
    With skiaBtn2
        If .Text = "Gradient Button" Then
            .Text = "Clicked!"
            .SetGradient &H4CAF50, &H2E7D32
        Else
            .Text = "Gradient Button"
            .SetGradient &HFF6B6B, &HFF8E53
        End If
    End With
End Sub

Private Sub skiaBtn3_Click()

    Dim i As Integer
    For i = 1 To 3
        skiaBtn3.BackColor = &HFFEB3B
        skiaBtn3.Refresh
        Sleep 100
        skiaBtn3.BackColor = &H9C27B0
        skiaBtn3.Refresh
        Sleep 100
    Next i
    
    MsgBox "Custom Button!", vbInformation, "Click Event"
End Sub


Private Sub CreateStaticButton()

End Sub

