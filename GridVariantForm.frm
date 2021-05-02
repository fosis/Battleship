VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GridVariantForm 
   Caption         =   "Select an enemy grid"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "GridVariantForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GridVariantForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EGCheck As Boolean

Private Sub CancelButton_Click()

    Unload Me

End Sub

Private Sub GridNumberBox_Change()

    EGNumber = GridNumberBox.Value
    
    If EGNumber = "" Then
    
        EGCheck = False
        Exit Sub
    
    ElseIf EGNumber Like "-*" Then
    
        EGCheck = False
        MsgBox "Typed value must be greater than 0 ;)"
        Exit Sub
    
    ElseIf IsNumeric(EGNumber) Then
        
        EGNumber = Int(EGNumber)
        If EGNumber <> Int(EGNumber) Then
        
            EGCheck = False
            MsgBox "Selected number is not whole number"
            Exit Sub
        
        ElseIf EGNumber > 6 Or EGNumber < 1 Then
        
            EGCheck = False
            MsgBox "Selected number is not between 1 and 6"
            Exit Sub
        
        Else
        
            EGCheck = True
        
        End If
    
    Else
    
        EGCheck = False
        MsgBox "Typed value is not numeric"
        Exit Sub
        
    End If

End Sub

Private Sub OKButton_Click()

    EGNumber = GridNumberBox.Value
    
    'check number
    If EGCheck = False Then
    
        MsgBox "Input is not correct"
        Exit Sub
        
    Else
    
        GridNumber = EGNumber
        
    End If
    
    Unload Me
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Unload Me

End Sub

