VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5DC35748-D70A-417E-93B7-A488F085B02F}#89.0#0"; "SMARTNETBUTTON.OCX"
Begin VB.Form frmCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSub2 
      Height          =   195
      Left            =   5880
      TabIndex        =   14
      Top             =   240
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkSub1 
      Height          =   195
      Left            =   5040
      TabIndex        =   13
      Top             =   240
      Value           =   1  'Checked
      Width           =   255
   End
   Begin SmartNetButtonProject.SmartNetButton cmdCheck 
      Height          =   375
      Left            =   45
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   661
      Caption         =   "Check Your Numbers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   1
      CaptionBackColor=   -2147483624
      CaptionAreaPercent=   100
      ShowCaption     =   -1  'True
   End
   Begin RichTextLib.RichTextBox txtWinAnswer 
      Height          =   4095
      Left            =   45
      TabIndex        =   9
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7223
      _Version        =   393217
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmCheck.frx":0000
   End
   Begin VB.TextBox txtNo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   7
      Left            =   5300
      MaxLength       =   3
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtNo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   6
      Left            =   4440
      MaxLength       =   3
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame fraNos 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   50
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   6
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox txtNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   5
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox txtNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   4
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox txtNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   3
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox txtNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   600
         MaxLength       =   3
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox txtNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         MaxLength       =   3
         TabIndex        =   1
         Top             =   0
         Width           =   495
      End
   End
   Begin SmartNetButtonProject.SmartNetButton cmdCancel 
      Height          =   375
      Left            =   3195
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   661
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   1
      CaptionBackColor=   -2147483624
      CaptionAreaPercent=   100
      ShowCaption     =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subs:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************
'*==================================*
'*=Created by Matthew Simon 2001(c)=*
'*==================================*
'************************************

Private Sub cmdCancel_Click()
Dim A As Integer
'hides showing form and enabled main one
    frmCheck.Hide
    frmMain.Enabled = True
    frmMain.Show
'clears all text boxes
    For A = 0 To 7
        txtNo(A).Text = ""
    Next
    txtWinAnswer.Text = ""
End Sub
Private Sub cmdCheck_Click()
Dim A As Integer, B As Integer, C As Integer, D As Integer, E As Integer, F As Integer
Dim templine As String, NextNo As Boolean, CharLen As Integer
Dim Char As String, Char2 As String, Char3 As String, TempSub As String
Dim NoCorrect(1 To 1000000) As WinNonLineNo
Dim tempNo As Integer, TempText As String, Winner As String

'First checks if numbers inputed is correct
For E = 0 To (ViewingNo - 1)
    If IsNumeric(txtNo(E).Text) = False Or Val(txtNo(E).Text) < 0 Then
    'tells user if wrong input
        MsgBox "Sorry, one of the numbers you have entered in is incorrect", vbOKOnly, "Error"
    'exits sub
        Exit Sub
    End If
    For F = 0 To (ViewingNo - 1)
        If txtNo(E).Text = txtNo(F).Text Then
            If E <> F Then
            'tells user if wrong input
                MsgBox "Sorry, one or more of the numbers you have entered in is the same", vbOKOnly, "Error"
            'exits sub
                Exit Sub
            End If
        End If
    Next
Next

tempNo = 0
NextNo = False

'opens the file that numbers saved in
Open frmMain.cdlOpen.FileName For Input As 1
        'loops until end of file
        Do Until EOF(1)
        'places line that currently reading in templine
        Line Input #1, templine
        'variable for use later to work out the line no.
            tempNo = tempNo + 1
        'as the first line is text the below If makes sure below code only
        'starts if the second line is being read
            If tempNo > 1 Then
                'puts down to line no.
                    NoCorrect(tempNo - 1).LineNo = Val(tempNo - 1)
                'loops for the amount of characters in each line
                For B = 1 To Len(templine)
                    'places the current character into char variable
                    Char = Mid(templine, B, 1)
                    'if a char variable was a space the loop before
                    If NextNo = True Then
                        NextNo = False
                        'loop below finds out if number is 1,2 or 3 characters (e.g 0 - 999)
                        Do Until D = 3
                            D = D + 1
                            'places the next 3 characters  after the character found in the char variable one by one after each loop
                            Char2 = Mid(templine, (B + D), 1)
                            'if char2 is a space, using the do loop it finds if the number is 1,2 or 3 digits
                            If Char2 = " " Then
                                'if 1 digit
                                If D = 1 Then
                                    CharLen = 1
                                    Exit Do
                                'if 2 digit
                                ElseIf D = 2 Then
                                    CharLen = 2
                                    Exit Do
                                'if 3 digit
                                Else
                                    CharLen = 3
                                    Exit Do
                                End If
                            End If
                        Loop
                        'resets d for next time
                        D = 0
                        'char3 is the actual final number
                        Char3 = Mid(templine, (B), CharLen)
                        'checks the number to each number in text boxes
                            For C = 1 To ViewingNo
                                If Char3 = txtNo(C - 1).Text Then
                                    'if the numbers are the same adds 1 to variable
                                    NoCorrect(tempNo - 1).NoCorrect = Val(NoCorrect(tempNo - 1).NoCorrect) + 1
                                End If
                            Next
                        'checks the subs
                            If chkSub1.Value = 1 And Char3 = txtNo(6).Text Then
                                If IsNumeric(txtNo(6).Text) = True Or Val(txtNo(6).Text) < 0 Then
                                    NoCorrect(tempNo - 1).NoSubCorrect = Val(NoCorrect(tempNo - 1).NoSubCorrect) + 1
                                Else
                                    'tells user if wrong input
                                        MsgBox "Sorry, the first sub has an incorrect value", vbOKOnly, "Error"
                                    'exits sub
                                        Exit Sub
                                End If
                            End If
                            If chkSub2.Value = 1 And Char3 = txtNo(7).Text Then
                                If IsNumeric(txtNo(7).Text) = True Or Val(txtNo(7).Text) < 0 Then
                                    NoCorrect(tempNo - 1).NoSubCorrect = Val(NoCorrect(tempNo - 1).NoSubCorrect) + 1
                                Else
                                    'tells user if wrong input
                                        MsgBox "Sorry, the second sub has an incorrect value", vbOKOnly, "Error"
                                    'exits sub
                                        Exit Sub
                                End If
                            End If
                        NextNo = False
                    End If
                'if the char is a space then nextno is true
                'this is done as the actual numbers each appear after a space
                'so if you find out the placing of the space the number will be the following 1,2 or 3 characters
                If Char = " " Then
                    NextNo = True
                End If
                Next
                If NoCorrect(tempNo - 1).NoCorrect = ViewingNo Then
                Winner = " --WINNER!--"
                Else
                Winner = ""
                End If
                'places answer into main text box
                If chkSub1.Value = 1 Or chkSub2.Value = 1 Then
                    TempSub = " -- " & NoCorrect(tempNo - 1).NoSubCorrect & " Sub(s)"
                Else
                    TempSub = ""
                End If
                    txtWinAnswer.Text = TempWin & NoCorrect(tempNo - 1).LineNo & ":" & Chr(32) & NoCorrect(tempNo - 1).NoCorrect & " Number(s)" & TempSub & Winner & vbCrLf
                'places all text into temp to be used next time
                    TempWin = txtWinAnswer.Text
            End If
        Loop
    Close #1

End Sub
