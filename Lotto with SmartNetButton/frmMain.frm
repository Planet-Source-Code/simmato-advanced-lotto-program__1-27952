VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5DC35748-D70A-417E-93B7-A488F085B02F}#89.0#0"; "SMARTNETBUTTON.OCX"
Begin VB.Form frmMain 
   Caption         =   "Your Lotto Numbers"
   ClientHeight    =   5220
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5205
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAll 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      Begin SmartNetButtonProject.SmartNetButton cmdNo 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Click to produce random numbers."
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         Caption         =   "Produce six random numbers"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionBackColor=   -2147483624
         CaptionAreaPercent=   100
         ShowCaption     =   -1  'True
         BackColorPush   =   -2147483646
      End
      Begin VB.Frame fraMain 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000E&
         Height          =   1095
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4410
         Begin VB.Frame fraNumbers 
            BackColor       =   &H80000002&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   4215
            Begin VB.Label lblNo 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   3600
               TabIndex        =   10
               ToolTipText     =   "Random Number 6"
               Top             =   0
               Width           =   615
            End
            Begin VB.Label lblNo 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   2880
               TabIndex        =   9
               ToolTipText     =   "Random Number 5"
               Top             =   0
               Width           =   615
            End
            Begin VB.Label lblNo 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   2160
               TabIndex        =   8
               ToolTipText     =   "Random Number 4"
               Top             =   0
               Width           =   615
            End
            Begin VB.Label lblNo 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   1440
               TabIndex        =   7
               ToolTipText     =   "Random Number 3"
               Top             =   0
               Width           =   615
            End
            Begin VB.Label lblNo 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   720
               TabIndex        =   6
               ToolTipText     =   "Random Number 2"
               Top             =   0
               Width           =   615
            End
            Begin VB.Label lblNo 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   0
               TabIndex        =   5
               ToolTipText     =   "Random Number 1"
               Top             =   0
               Width           =   615
            End
         End
      End
      Begin MSComDlg.CommonDialog cdlSave 
         Left            =   1080
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdlPrint 
         Left            =   1560
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         FromPage        =   1
         ToPage          =   1
      End
      Begin MSComDlg.CommonDialog cdlFont 
         Left            =   2040
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   3
         FontName        =   "Arial"
      End
      Begin MSComDlg.CommonDialog cdlOpen 
         Left            =   2520
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin RichTextLib.RichTextBox txtLottoNo 
         Height          =   3375
         Left            =   0
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Your Lotto Numbers"
         Top             =   1080
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   5953
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         DisableNoScroll =   -1  'True
         TextRTF         =   $"frmMain.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnudash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheck 
         Caption         =   "&Check Your Numbers"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnudash6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnudash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptionsMain 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPowerball 
         Caption         =   "Powerball Option"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy Numbers"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear &Numbers"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Font"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyNo 
         Caption         =   "Copy Numbers"
      End
      Begin VB.Menu mnuClearNo 
         Caption         =   "Clear Numbers"
      End
      Begin VB.Menu hjkhjik 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFonts 
         Caption         =   "Font"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************
'*==================================*
'*=Created by Matthew Simon 2001(c)=*
'*==================================*
'************************************

Option Explicit
Private Sub getNumbers()
Dim A As Integer, B As Integer, C As Integer, D As Integer, E As Integer, Z As Integer
Dim AllNo As String, no(1 To 6) As String
Dim finish As Boolean, noTemp As Integer, LeftNoValue As Integer, NoOfNumbers As Integer

'if powerball mnu checked goes to powerball sub
If mnuPowerball.Checked = True Then
'goes to powerball sub
    PowerBall
'exits current sub
    Exit Sub
End If
'gets the random no.s
    Randomize
For Z = 1 To ViewingNo
    no(Z) = Int(Rnd * (HighNo - LowNo)) + LowNo
Next

'makes sure the random no.s are not the same
For A = 1 To ViewingNo
    For B = 1 To ViewingNo
        If no(A) = no(B) Then
            If A <> B Then
                'if the number is the same the sub is done again
                getNumbers
                Exit Sub
            End If
        End If
    Next
Next

'places the random no.s into the text boxes
For C = 1 To ViewingNo
    lblNo(C - 1).Caption = no(C)
Next

If frmChangeValue.opt(0).Value = True Then
'code below sorts the no in ascending order
NoOfNumbers = ViewingNo
    Do
        finish = True
        
        For LeftNoValue = 1 To NoOfNumbers - 1
        'if a no to the right is larger thna the no to the left than they swap
            If Val(no(LeftNoValue)) > Val(no(LeftNoValue + 1)) Then
                'places left no. in temp string
                    noTemp = Val(no(LeftNoValue))
                'places right no. in lower no place
                    no(LeftNoValue) = Val(no(LeftNoValue + 1))
                'takes the no. that was in left side in right side
                    no(LeftNoValue + 1) = noTemp
                'finish = to false so does this again to check if in right order
                    finish = False
            End If
        Next
    
        NoOfNumbers = NoOfNumbers - 1
    
    Loop Until finish = True
ElseIf frmChangeValue.opt(1).Value = True Then
'code below sorts the no in descending order
NoOfNumbers = 6
    Do
        finish = True
        
        For LeftNoValue = 1 To NoOfNumbers - 1
        'if a no to the right is larger thna the no to the left than they swap
            If Val(no(LeftNoValue)) < Val(no(LeftNoValue + 1)) Then
                'places left no. in temp string
                    noTemp = Val(no(LeftNoValue))
                'places right no. in lower no place
                    no(LeftNoValue) = Val(no(LeftNoValue + 1))
                'takes the no. that was in left side in right side
                    no(LeftNoValue + 1) = noTemp
                'finish = to false so does this again to check if in right order
                    finish = False
            End If
        Next
    
        NoOfNumbers = NoOfNumbers - 1
    
    Loop Until finish = True
End If


'places all the no.s into one string
    AllNo = " " & no(1) & " " & no(2) & " " & no(3) & " " & no(4) & " " & no(5) & " " & no(6)
'tells user later on by this how many sets of numbers they have produced
    Clicktime = Clicktime + 1
'places everything together and put sit into the main text box
    txtLottoNo.Text = TempText & vbCrLf & Clicktime & ":" & AllNo
'places all the text in the text box into a tempory string
'to be used when this sub is done again to keep on adding text to text box
    TempText = txtLottoNo.Text
'puts focus at end of text box
    txtLottoNo.SelStart = Len(txtLottoNo.Text) - 1
'code below redoes this sub the no. of does the user siad in the options
NoPerClickTemp = NoPerClickTemp - 1
'refreshes frm to give appearance that no. are spinning
    frmMain.Refresh
'when sub has be redone the no. of times, it exit
If NoPerClickTemp = 0 Then
NoPerClickTemp = NoPerClick
    Exit Sub
End If
getNumbers
End Sub
Private Sub PowerBall()
Dim A As Integer, B As Integer, C As Integer, Z As Integer
Dim AllNo As String, no(1 To 6) As String
Dim finish As Boolean, noTemp As Integer, LeftNoValue As Integer, NoOfNumbers As Integer

'gets the random no.s
    Randomize
For Z = 2 To 6
    no(Z) = Int(Rnd * (HighNo - LowNo)) + LowNo
Next

'makes sure the random no.s are not the same
For A = 1 To 5
    For B = 1 To 5
        If no(A) = no(B) Then
            If A <> B Then
                'if the number is the same the sub is done again
                PowerBall
                Exit Sub
            End If
        End If
    Next
Next
'gets powerball
    no(1) = Int(Rnd * HighNo) + LowNo
'places the random no.s into the small text boxes
For C = 1 To 6
    lblNo(C - 1).Caption = no(C)
Next

If frmChangeValue.opt(0).Value = True Then
'code below sorts the no in ascending order
NoOfNumbers = 6
    Do
        finish = True
        
        For LeftNoValue = 2 To NoOfNumbers - 1
        'if a no to the right is larger thna the no to the left than they swap
            If Val(no(LeftNoValue)) > Val(no(LeftNoValue + 1)) Then
                'places left no. in temp string
                    noTemp = Val(no(LeftNoValue))
                'places right no. in lower no place
                    no(LeftNoValue) = Val(no(LeftNoValue + 1))
                'takes the no. that was in left side in right side
                    no(LeftNoValue + 1) = noTemp
                'finish = to false so does this again to check if in right order
                    finish = False
            End If
        Next
    
        NoOfNumbers = NoOfNumbers - 1
        If NoOfNumbers = 3 Then
        NoOfNumbers = 6
        End If
    Loop Until finish = True
ElseIf frmChangeValue.opt(1).Value = True Then
'code below sorts the no in descending order
NoOfNumbers = 6
    Do
        finish = True
        
        For LeftNoValue = 2 To NoOfNumbers - 1
        'if a no to the right is larger thna the no to the left than they swap
            If Val(no(LeftNoValue)) < Val(no(LeftNoValue + 1)) Then
                'places left no. in temp string
                    noTemp = Val(no(LeftNoValue))
                'places right no. in lower no place
                    no(LeftNoValue) = Val(no(LeftNoValue + 1))
                'takes the no. that was in left side in right side
                    no(LeftNoValue + 1) = noTemp
                'finish = to false so does this again to check if in right order
                    finish = False
            End If
        Next
    
        NoOfNumbers = NoOfNumbers - 1
    
    Loop Until finish = True
End If

'places all the no.s into one string except the first no (powerball)
    AllNo = " " & lblNo(1) & " " & lblNo(2) & " " & lblNo(3) & " " & lblNo(4) & " " & lblNo(5)
'tells user later on by this how many sets of numbers they have produced
    Clicktime = Clicktime + 1
'places everything together and put sit into the main text box
    txtLottoNo.Text = TempText & vbCrLf & Clicktime & ": " & lblNo(0) & AllNo
'places all the text in the text box into a tempory string
'to be used when this sub is done again to keep on adding text to text box
    TempText = txtLottoNo.Text
'puts focus at end of text box
    txtLottoNo.SelStart = Len(txtLottoNo.Text) - 1
'refreshes frm to give appearance that no. are spinning
    frmMain.Refresh
'code below redoes this sub the no. of does the user siad in the options
NoPerClickTemp = NoPerClickTemp - 1
'when sub has be redone the no. of times, it exit
If NoPerClickTemp = 0 Then
NoPerClickTemp = NoPerClick
    Exit Sub
End If
getNumbers
End Sub
Private Sub PrintNumbers()
Dim A As Integer, printoption As Integer, tempprint As String
'on error go to bottom of sub
    On Error GoTo cancelPrint
'shows printer dialog box
    cdlPrint.ShowPrinter
'sets printer options
    With Printer
        .FontName = cdlFont.FontName
        .FontSize = cdlFont.FontSize
        .FontBold = cdlFont.FontBold
        .FontItalic = cdlFont.FontItalic
        .FontStrikethru = cdlFont.FontStrikethru
        .FontUnderline = cdlFont.FontUnderline
        .Copies = cdlPrint.Copies
    End With
'gets tempprint string from main txt
    tempprint = txtLottoNo.Text
'prints stuff
    Printer.Print tempprint
'ends doc
    Printer.EndDoc
'if error exits sub
cancelPrint:
Exit Sub
End Sub
Private Sub cmdNo_Click()
'goes to sub to get no.s
    getNumbers
End Sub
Private Sub Form_Load()
'sets default random values
HighNo = 45
LowNo = 1
TempText = "Lotto Random Numbers between 0 and 45"
NoPerClick = 1
NoPerClickTemp = NoPerClick
ViewingNo = 6
frmChangeValue.comRNo.ListIndex = 5
End Sub
Private Sub Form_Resize()
'if resized less than supposed than put back to minimum size
If frmMain.Width < 4545 Then
    frmMain.Width = 4545
ElseIf frmMain.Height < 5190 Then
    frmMain.Height = 5190
Else
'resizes all frames, buttons etc to fit
    fraAll.Width = Me.Width
    fraAll.Height = Me.Height
    fraMain.Width = Me.Width - 135
    txtLottoNo.Height = Me.Height - 1815
    txtLottoNo.Width = Me.Width - 135
    cmdNo.Width = fraMain.Width - 195
    fraNumbers.Left = (fraMain.Width / 2) - (fraNumbers.Width / 2)
End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
frmMain.Enabled = False
End Sub
Private Sub mnuCheck_Click()
frmCheck.Show
frmMain.Enabled = False
End Sub
Private Sub mnuOptions_Click()
'shows options frm and disables main form
    frmChangeValue.Show
    frmMain.Enabled = False
End Sub
Private Sub mnuClear_Click()
Dim A As Integer
'clears main text box
    txtLottoNo.Text = ""
'clears the 6 small text boxes
    For A = 0 To 5
        lblNo(A).Caption = ""
    Next
'sets temptext back to just title so on next click old numbers dont appear again
    TempText = "Lotto Random Numbers between " & LowNo & " and " & HighNo
'sets clicktime back to 0
    Clicktime = 0
End Sub
Private Sub mnuClearNo_Click()
'goes to sub to clear
    mnuClear_Click
End Sub
Private Sub mnuCopy_Click()
'clears clipboard
    Clipboard.Clear
'sets text to clipboard
    Clipboard.SetText txtLottoNo.Text
End Sub
Private Sub mnuCopyNo_Click()
'goes to sub to copy
    mnuCopy_Click
End Sub
Private Sub mnuExit_Click()
Dim i As Integer
'End program and unload all forms
For i = Forms.Count - 1 To 0 Step -1
    Unload Forms(i)
Next i
End
End Sub
Private Sub mnuFont_Click()
'on error goes to bottom of sub
    On Error GoTo CancelFont
'shows font dialog
    cdlFont.ShowFont
'sets focus to text box.
txtLottoNo.SetFocus
' Start highlight before first character.
txtLottoNo.SelStart = 0
' Highlight to end of text.
txtLottoNo.SelLength = Len(txtLottoNo.Text)
'sets settings
    txtLottoNo.SelFontName = cdlFont.FontName
    txtLottoNo.SelFontSize = cdlFont.FontSize
    txtLottoNo.SelBold = cdlFont.FontBold
    txtLottoNo.SelItalic = cdlFont.FontItalic
    txtLottoNo.SelStrikeThru = cdlFont.FontStrikethru
    txtLottoNo.SelUnderline = cdlFont.FontUnderline
'takes away highlight
    txtLottoNo.SelStart = 0
    txtLottoNo.SelLength = ""
'exits sub if error
CancelFont:
Exit Sub
End Sub
Private Sub mnuFonts_Click()
'goes to font sub
    mnuFont_Click
End Sub
Private Sub mnuPowerball_Click()
Dim warning As Integer, A As Integer
'warns if carry on
    warning = MsgBox("Warning: By pressing OK your current random numbers will be cleared", vbExclamation + vbOKCancel, "Warning")
If warning = 1 Then
    If mnuPowerball.Checked = False Then
    'checks powerball option to show it is on
        mnuPowerball.Checked = True
    'clears main text box
        frmMain.txtLottoNo.Text = ""
    'clears the 6 small text boxes
        For A = 0 To 5
            frmMain.lblNo(A).Caption = ""
        Next
    'resets temptext
        TempText = "Powerball Lotto Random Numbers between " & LowNo & " and " & HighNo
    'sets clicktime back to 0
        Clicktime = 0
    'makes first small text box text bold and red to show powerball
        lblNo(0).FontBold = True
        lblNo(0).ForeColor = vbRed
    Else
    'checks powerball option to show it is off
        mnuPowerball.Checked = False
    'clears main text box
        frmMain.txtLottoNo.Text = ""
    'clears the 6 small text boxes
        For A = 0 To 5
            frmMain.lblNo(A).Caption = ""
        Next
    'resets temptext
        TempText = "Lotto Random Numbers between " & LowNo & " and " & HighNo
    'sets clicktime back to 0
        Clicktime = 0
    'makes first small text box text normal like the others
        lblNo(0).FontBold = False
        lblNo(0).ForeColor = vbBlack
    End If
End If
End Sub
Private Sub mnuPrint_Click()
'goes to print sub
    PrintNumbers
End Sub
Private Sub mnuOpen_Click()
Dim templine As String, TempText As String
'if error goes to bottom of sub
    On Error GoTo cancelOpen
'sets some properties to dialog box
    cdlOpen.Filter = "Your Lotto Numbers Document|*.yln|Text Documents|*.txt|All Files (*.*)|*.*"
'shows save dialog box
    cdlOpen.ShowOpen
'saves text to filename
    Open cdlOpen.FileName For Input As 1
        Do Until EOF(1)
        Line Input #1, templine
        TempText = TempText + templine & Chr(13) & Chr(10)
        Loop
    Close #1
txtLottoNo.Text = TempText
'exits sub if error
cancelOpen:
    Exit Sub
End Sub
Private Sub mnuSave_Click()
'if error goes to bottom of sub
    On Error GoTo cancelsave
'sets some properties to dialog box
    cdlSave.Filter = "Your Lotto Numbers Document|*.yln|Text Documents|*.txt|All Files (*.*)|*.*"
'shows save dialog box
    cdlSave.ShowSave
'saves text to filename
    Open cdlSave.FileName For Output As 1
        Print #1, txtLottoNo.Text
    Close #1
'exits sub if error
cancelsave:
    Exit Sub
End Sub
Private Sub txtLottoNo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'if rightclick on textbox
    If Button = vbRightButton Then
        PopupMenu frmMain.mnuPopUp
    End If
End Sub
