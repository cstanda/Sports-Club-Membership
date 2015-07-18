VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Club Membership"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAddress3 
      Height          =   495
      Left            =   2160
      TabIndex        =   23
      Top             =   5280
      Width           =   4575
   End
   Begin VB.TextBox txtMtype 
      Height          =   615
      Left            =   9120
      TabIndex        =   22
      ToolTipText     =   "Choose membership type F/S/T/ B"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtGender 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9120
      MaxLength       =   1
      TabIndex        =   20
      ToolTipText     =   "Sex must be F (Female) or M (Male)"
      Top             =   1320
      Width           =   495
   End
   Begin VB.ComboBox txtSub 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "mmm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "SportsClubMembership.frx":0000
      Left            =   9360
      List            =   "SportsClubMembership.frx":0028
      TabIndex        =   18
      ToolTipText     =   "Pick subcription month (MMM) e.g Jan"
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Record "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox txtJoinDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   15
      ToolTipText     =   "Enter date of joining club (dd/mm/yyyy) e.g. 20/03/1990"
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txtDateOfBirth 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   13
      ToolTipText     =   "Enter Date of Birth (dd/mm/yyyy) e.g. 01/12/1990"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtPost 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      ToolTipText     =   "Enter your postal code"
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox txtAddress2 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   4560
      Width           =   4575
   End
   Begin VB.TextBox txtAddress1 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   3960
      Width           =   4575
   End
   Begin VB.TextBox txtLname 
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txtFname 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtMembership 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   2
      ToolTipText     =   "e.g 123456"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FF0000&
      Caption         =   "Chris Standa Solutions"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   24
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label Label 
      BackColor       =   &H00FF0000&
      Caption         =   "Type of Membership"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   21
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF0000&
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   19
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF0000&
      Caption         =   "SUBSCRIPTION DUE MONTH"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   16
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF0000&
      Caption         =   "Join Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      Caption         =   "Date of birth"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      Caption         =   "Postcode"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "Membership number"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "MEMBER FORM"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub DateOfBirth_Change()

End Sub

Private Sub cmdSave_Click()


If Not IsNumeric(txtMembership.Text) Then

        MsgBox "Membership Number is not numeric", vbExclamation, ""

Cancel = True

txtMembership.SetFocus
        Exit Sub
        
End If

If txtMembership.Text < 6 Then

        MsgBox "Membership Number is not 6 digits", vbExclamation, ""

Cancel = True

txtMembership.SetFocus
        Exit Sub
        
End If

If txtFname.Text = "" Then
        MsgBox "Please enter your first name!", vbExclamation, ""
        Cancel = True
        
        txtFname.SetFocus
        Exit Sub
End If

If txtLname.Text = "" Then
        MsgBox "Please enter your last name!", vbExclamation, ""
        Cancel = True
        
        txtLname.SetFocus
        Exit Sub
End If

If txtAddress1.Text = "" Then
        MsgBox "Please enter an address", vbExclamation, ""
        Cancel = True
        
        txtAddress1.SetFocus
        Exit Sub
End If

If txtPost.Text = "" Then
        MsgBox "Please provide a post code", vbExclamation, ""
        Cancel = True
        
        txtPost.SetFocus
        Exit Sub
End If

If txtGender.Text = "" Then
        MsgBox "Sex must be F or M", vbExclamation, ""
        Cancel = True
        
        txtGender.SetFocus
        Exit Sub
        
End If

If txtDateOfBirth.Text = "" Then
        MsgBox "Invalid Date of Birth", vbExclamation, ""
        Cancel = True
        
        txtDateOfBirth.SetFocus
        Exit Sub
End If

If txtJoinDate.Text = "" Then
        MsgBox "Invalid Join Date", vbExclamation, ""
        Cancel = True
        
        txtJoinDate.SetFocus
        Exit Sub
End If

If txtMtype.Text = "" Then
        MsgBox "Membership type must be F, S, T or B", vbExclamation, ""
        Cancel = True
End If

If txtMtype.Text = "" Then
        MsgBox "Subscription month invalid", vbExclamation, ""
        Cancel = True
        
        txtMtype.SetFocus
        Exit Sub
End If

MsgBox "First Name : " + txtFname.Text + Chr(0) + "Last Name : " + txtLname.Text + Chr(0) + "Address : " + txtAddress1.Text + Chr(0) + "Post Code: " + txtPostText + Chr(0) + "Gender : " + txtGender.Text + Chr(1) + "Date of Birth : " + txtDateOfBirth.Text + Chr(0) + "Join Date : " + txtJoinDate.Text + Chr(0) + "Type of Membership : " + txtMtype.Text + Chr(1) + "Subscription Due Month : " + txtSub.Text, vbInformation, -"Your registration to the club is successful !!!"

End
    
End Sub


Private Sub Form_Load()
On Error Resume Next
    ValidateControls
    If Err = 380 Then
        Cancel = True
End If
End Sub

