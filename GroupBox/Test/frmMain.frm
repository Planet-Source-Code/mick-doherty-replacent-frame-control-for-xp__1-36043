VERSION 5.00
Object = "*\A..\..\..\REPLAC~1\GroupBox\MDGroupBox.vbp"
Begin VB.Form frmMain 
   Caption         =   "XP Form"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin MDGroupBox.GroupBox GroupBox1 
      Height          =   2055
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   3625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&GroupBox1"
      Enabled         =   0   'False
      Begin VB.CommandButton Command4 
         Caption         =   "Command1"
         Height          =   375
         Left            =   5040
         TabIndex        =   33
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Option1"
         Height          =   315
         Left            =   5160
         TabIndex        =   32
         Top             =   780
         Width           =   915
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Option2"
         Height          =   315
         Left            =   5160
         TabIndex        =   31
         Top             =   1080
         Width           =   915
      End
      Begin MDGroupBox.GroupBox GroupBox2 
         Height          =   1395
         Left            =   120
         TabIndex        =   22
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2461
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   900
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option1"
            Height          =   315
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option2"
            Height          =   315
            Left            =   240
            TabIndex        =   20
            Top             =   540
            Width           =   915
         End
      End
      Begin MDGroupBox.GroupBox GroupBox3 
         Height          =   1395
         Left            =   1740
         TabIndex        =   23
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2461
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton Command2 
            Caption         =   "Command1"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   900
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Option1"
            Height          =   315
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Option2"
            Height          =   315
            Left            =   240
            TabIndex        =   25
            Top             =   540
            Width           =   915
         End
      End
      Begin MDGroupBox.GroupBox GroupBox4 
         Height          =   1395
         Left            =   3360
         TabIndex        =   24
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2461
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Begin VB.CommandButton Command3 
            Caption         =   "Command1"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   900
            Width           =   1215
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Option1"
            Height          =   315
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Option2"
            Height          =   315
            Left            =   240
            TabIndex        =   28
            Top             =   540
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Frame1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   2940
      Width           =   6555
      Begin VB.OptionButton Option16 
         Caption         =   "Option2"
         Height          =   315
         Left            =   5160
         TabIndex        =   17
         Top             =   1080
         Width           =   915
      End
      Begin VB.OptionButton Option15 
         Caption         =   "Option1"
         Height          =   315
         Left            =   5160
         TabIndex        =   16
         Top             =   780
         Width           =   915
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command1"
         Height          =   375
         Left            =   4980
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   1395
         Left            =   3240
         TabIndex        =   11
         Top             =   540
         Width           =   1455
         Begin VB.CommandButton Command7 
            Caption         =   "Command1"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   900
            Width           =   1215
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Option1"
            Height          =   315
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Option2"
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   540
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1395
         Left            =   1680
         TabIndex        =   7
         Top             =   540
         Width           =   1455
         Begin VB.CommandButton Command6 
            Caption         =   "Command1"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   900
            Width           =   1215
         End
         Begin VB.OptionButton Option12 
            Caption         =   "Option1"
            Height          =   315
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Option2"
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   540
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   1395
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   1455
         Begin VB.OptionButton Option10 
            Caption         =   "Option2"
            Height          =   315
            Left            =   240
            TabIndex        =   6
            Top             =   540
            Width           =   915
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Option1"
            Height          =   315
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   915
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Command1"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   900
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "Enable"
      Height          =   435
      Left            =   1920
      TabIndex        =   1
      Top             =   2340
      Width           =   1395
   End
   Begin VB.CommandButton cmdVisible 
      Caption         =   "Hide"
      Height          =   435
      Left            =   3480
      TabIndex        =   0
      Top             =   2340
      Width           =   1395
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnable_Click()
    Select Case cmdEnable.Caption
        Case "Enable"
            GroupBox1.Enabled = True
            Frame1.Enabled = True
            cmdEnable.Caption = "Disable"
        Case Else
            GroupBox1.Enabled = False
            Frame1.Enabled = False
            cmdEnable.Caption = "Enable"
    End Select
End Sub

Private Sub cmdVisible_Click()
    Select Case cmdVisible.Caption
        Case "Show"
            GroupBox1.Visible = True
            Frame1.Visible = True
            cmdVisible.Caption = "Hide"
        Case Else
            GroupBox1.Visible = False
            Frame1.Visible = False
            cmdVisible.Caption = "Show"
    End Select
End Sub

