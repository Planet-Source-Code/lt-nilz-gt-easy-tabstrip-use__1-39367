VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   4
      Top             =   480
      Width           =   4215
      Begin VB.Label Label1 
         Caption         =   "Number 4 !!!!!!"
         Height          =   735
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   3
      Top             =   480
      Width           =   4215
      Begin VB.Label Label2 
         Caption         =   "Number 3"
         Height          =   735
         Left            =   1200
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   480
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "Tab 2"
         Height          =   855
         Left            =   480
         TabIndex        =   8
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Number 2"
         Height          =   735
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   480
      Width           =   4215
      Begin VB.CommandButton Command2 
         Caption         =   "Number 1"
         Height          =   855
         Left            =   840
         TabIndex        =   10
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Number 1"
         Height          =   855
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "First Tab !!!"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Second Tab !!!"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Third Tab !!!"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fourth Tab !!!"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

    Picture1.Visible = True
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False

End Sub

Private Sub TabStrip1_Click()

If TabStrip1.SelectedItem.Index = 1 Then
    Picture1.Visible = True
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False

ElseIf TabStrip1.SelectedItem.Index = 2 Then
    Picture1.Visible = False
    Picture2.Visible = True
    Picture3.Visible = False
    Picture4.Visible = False
    
ElseIf TabStrip1.SelectedItem.Index = 3 Then
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = True
    Picture4.Visible = False

ElseIf TabStrip1.SelectedItem.Index = 4 Then
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = True
    
End If

End Sub
