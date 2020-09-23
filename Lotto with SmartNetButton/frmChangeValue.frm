VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChangeValue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Random Value"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   Icon            =   "frmChangeValue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstMain 
      Height          =   3375
      Left            =   50
      TabIndex        =   2
      Top             =   0
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmChangeValue.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtNoPerClick"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "opt(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "opt(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "opt(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "comAlign"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "comRNo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Random No. Value"
      TabPicture(1)   =   "frmChangeValue.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape1"
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(2)=   "lbl2"
      Tab(1).Control(3)=   "lbl1"
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(5)=   "Label3"
      Tab(1).Control(6)=   "txtHighNo"
      Tab(1).Control(7)=   "txtLowNo"
      Tab(1).ControlCount=   8
      Begin VB.ComboBox comRNo 
         Height          =   315
         ItemData        =   "frmChangeValue.frx":0044
         Left            =   3360
         List            =   "frmChangeValue.frx":005A
         TabIndex        =   22
         Top             =   1360
         Width           =   1095
      End
      Begin VB.ComboBox comAlign 
         Height          =   315
         ItemData        =   "frmChangeValue.frx":0070
         Left            =   1320
         List            =   "frmChangeValue.frx":007D
         TabIndex        =   20
         Top             =   900
         Width           =   1455
      End
      Begin VB.OptionButton opt 
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   18
         Top             =   2880
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton opt 
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   16
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton opt 
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   15
         Top             =   2145
         Width           =   255
      End
      Begin VB.TextBox txtNoPerClick 
         Height          =   285
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "1"
         Top             =   440
         Width           =   300
      End
      Begin VB.TextBox txtLowNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74640
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "1"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtHighNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -71640
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "45"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "How many random numbers would you like?"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1425
         Width           =   3255
      End
      Begin VB.Label Label9 
         Caption         =   "Text box Justify."
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Appear in no numbered order."
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "C"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Appear in a numbered order from highest to lowest."
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "C"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         X1              =   1560
         X2              =   5640
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label6 
         Caption         =   "Appear in a numbered order from lowest to highest."
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "C"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Random Numbers "
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "C"
         Top             =   1800
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "How many sets of random numbers would you like per click?"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Note: The number must be between 0 -999. As well the difference between both numbers must be more than 10."
         Height          =   375
         Left            =   -74760
         TabIndex        =   9
         Top             =   2160
         Width           =   5295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "AND"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72495
         TabIndex        =   8
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lower Number"
         Height          =   255
         Left            =   -74280
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lbl2 
         BackStyle       =   0  'Transparent
         Caption         =   "Higher Number"
         Height          =   255
         Left            =   -71280
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Please enter the values you would like to get random numbers between."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   4215
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Left            =   -74760
         Top             =   1200
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4888
      TabIndex        =   0
      Top             =   3480
      Width           =   975
   End
End
Attribute VB_Name = "frmChangeValue"
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
Dim OptCheck As Integer
Dim tempLowNo As Integer
Dim tempHighNo As Integer
Dim tempNoPerClick As Integer
Dim tempOption As Integer
Dim tempJustify As String

Private Sub cmdCancel_Click()
'hides and shows forms
    frmMain.Enabled = True
    frmChangeValue.Hide
'resets texts boxes
    txtLowNo.Text = ""
    txtHighNo.Text = ""
'places old values
    txtLowNo.Text = tempLowNo
    txtHighNo.Text = tempHighNo
    txtNoPerClick.Text = tempNoPerClick
    opt(OptCheck).Value = True
    comAlign.ListIndex = Val(frmMain.txtLottoNo.SelAlignment)
    comRNo.ListIndex = (ViewingNo - 1)
End Sub

Private Sub cmdOK_Click()
Dim note As Integer, tempNo As Integer, B As Integer
    'makes sure user entered the right format
If Val(txtLowNo.Text) = Val(txtHighNo.Text) Or Val(txtLowNo.Text) > Val(txtHighNo.Text) Or IsNumeric(txtHighNo.Text) = False Or IsNumeric(txtLowNo.Text) = False Or (Val(txtHighNo.Text) - Val(txtLowNo.Text)) < 2 Or IsNumeric(txtNoPerClick.Text) = False Or txtNoPerClick.Text < 0 Then
'if the user entered something wrong, the user is notified
    MsgBox "Sorry you have entered an incorrect value", vbOKOnly, "Error"
    txtLowNo.Text = ""
    txtHighNo.Text = ""
Else
    'if no changes to the 2 low/high text boxes and arrangement option then no need to reset all text boxes in frmMain
    If tempLowNo = txtLowNo.Text And tempHighNo = txtHighNo.Text And opt(OptCheck).Value = True And comRNo.ListIndex = ViewingNo - 1 Then
       'aligns main text box
            frmMain.txtLottoNo.SelStart = 1
            frmMain.txtLottoNo.SelLength = Len(frmMain.txtLottoNo.Text)
            If comAlign.ListIndex = 0 Then
                frmMain.txtLottoNo.SelAlignment = 0
            ElseIf comAlign.ListIndex = 1 Then
                frmMain.txtLottoNo.SelAlignment = 1
            Else
                frmMain.txtLottoNo.SelAlignment = 2
            End If
            frmMain.txtLottoNo.SelStart = Len(frmMain.txtLottoNo.Text)
            frmMain.txtLottoNo.SelLength = 1
        'sets number of sets to show per click
            NoPerClick = txtNoPerClick.Text
            NoPerClickTemp = NoPerClick
        'hides and shows forms
            frmMain.Enabled = True
            frmChangeValue.Hide
            Exit Sub
    Else
    'asks if want to continue
        note = MsgBox("Warning: By pressing OK your current random numbers will be cleared", vbExclamation + vbOKCancel, "Warning")
    'if press OK
        If note = 1 Then
        'if everything is correct
        'sets temptext to tell user what random no is between
            TempText = "Lotto Random Numbers between " & txtLowNo.Text & " and " & txtHighNo.Text
        'clears all text boxes
            Dim A As Integer
        'clears main text box
            frmMain.txtLottoNo.Text = ""
        'clears the 6 small text boxes
            For A = 0 To 5
                frmMain.lblNo(A).Caption = ""
            Next
        'sets clicktime back to 0
            Clicktime = 0
        'aligns main text box
            If comAlign.ListIndex = 0 Then
                frmMain.txtLottoNo.SelAlignment = 0
            ElseIf comAlign.ListIndex = 1 Then
                frmMain.txtLottoNo.SelAlignment = 1
            Else
                frmMain.txtLottoNo.SelAlignment = 2
            End If
        'sets values to variables
            HighNo = txtHighNo.Text
            LowNo = txtLowNo.Text
        'sets number of sets to show per click
            NoPerClick = txtNoPerClick.Text
            NoPerClickTemp = NoPerClick
        'sets no of random no
            ViewingNo = comRNo.ListIndex + 1
        'disables boxes not in use
            tempNo = 6 - ViewingNo
            For B = 0 To 5
                frmMain.lblNo(B).Visible = True
                frmCheck.txtNo(B).Visible = True
            Next
            
            Do Until tempNo = 0
                frmMain.lblNo(6 - tempNo).Visible = False
                frmCheck.txtNo(6 - tempNo).Visible = False
                tempNo = tempNo - 1
            Loop
            
            frmMain.fraNumbers.Width = ((comRNo.ListIndex + 1) * frmMain.lblNo(0).Width) + (105 * ViewingNo - 1)
            If frmMain.WindowState = 0 Then
            frmMain.Width = frmMain.Width
            frmMain.Height = frmMain.Height
            End If
        'hides and shows forms
            frmMain.Enabled = True
            frmChangeValue.Hide
        End If
    End If
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
'puts values into temp so if cancel is pressed puts in theses values
tempLowNo = txtLowNo.Text
tempHighNo = txtHighNo.Text
tempNoPerClick = txtNoPerClick.Text
comAlign.ListIndex = Val(frmMain.txtLottoNo.SelAlignment)

For i = 0 To 2
    If opt(i).Value = True Then
        OptCheck = i
    End If
Next
End Sub
