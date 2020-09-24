VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Background Color Change (Move Slider to Change Color or Resize Form)"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   4320
      ScaleHeight     =   1995
      ScaleWidth      =   3435
      TabIndex        =   10
      Top             =   2880
      Width           =   3495
      Begin MSComctlLib.Slider slideblue 
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Max             =   255
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider slidegreen 
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   12
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Max             =   255
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider slidered 
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Max             =   255
         TickStyle       =   3
      End
      Begin VB.Label Label8 
         Caption         =   "Blue"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Green"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Red"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   960
         TabIndex        =   16
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "B O T T O M"
         Height          =   1215
         Left            =   3240
         TabIndex        =   15
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label14 
         Caption         =   "RESULT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   360
      ScaleHeight     =   1995
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      Begin MSComctlLib.Slider slidered 
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Max             =   255
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider slidegreen 
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Max             =   255
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider slideblue 
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Max             =   255
         TickStyle       =   3
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Red"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Green"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Blue"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "T O P"
         Height          =   615
         Left            =   3240
         TabIndex        =   5
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "RESULT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    'INITIALIZATION OF BACKGROUND COLORS
    'USING RGB COLORS
    'INITIAL TOP COLOR    = BLACK : RGB(0,0,0)
    'INITIAL BOTTOM COLOR = RED   : RGB(255,0,0)
    
    slidered(0).Value = 0       'TOP COLOR
    slideblue(0).Value = 0      'TOP COLOR
    slidegreen(0).Value = 0     'TOP COLOR
    slidered(1).Value = 255     'BOTTOM COLOR
    slideblue(1).Value = 0      'BOTTOM COLOR
    slidegreen(1).Value = 0     'BOTTOM COLOR
    
    'TOP COLOR
    Label1.BackColor = RGB(slidered(0).Value, slidegreen(0).Value, slideblue(0).Value)
    
    'BOTTOM COLOR
    Label11.BackColor = RGB(slidered(1).Value, slidegreen(1).Value, slideblue(1).Value)
    
    'bgcolor FROM MODULE1
    bgcolor Me, slidered(0).Value, slidegreen(0).Value, slideblue(0).Value, slidered(1).Value, slidegreen(1).Value, slideblue(1).Value
End Sub

Private Sub Form_Load()
    'IF YOU WISH TO MAXIMIZE FORM ON LOADING
    'REMOVE COMMENT QUOTATIONS
    
    'Me.WindowState = 2
    'Me.Height = 11520
    
    'slidered(0).Value = 0       'TOP COLOR
    'slideblue(0).Value = 0      'TOP COLOR
    'slidegreen(0).Value = 0     'TOP COLOR
    'slidered(1).Value = 255     'BOTTOM COLOR
    'slideblue(1).Value = 0      'BOTTOM COLOR
    'slidegreen(1).Value = 0     'BOTTOM COLOR
    
    'bgcolor Me, slidered(0).Value, slidegreen(0).Value, slideblue(0).Value, slidered(1).Value, slidegreen(1).Value, slideblue(1).Value
End Sub

Private Sub Form_Resize()
    'CHANGING SCALEMODE (GRAPHIC UNITS) OF FORM TO TWIP
    Me.ScaleMode = 1
    
    'MOVING PICTURE2 (BOTTOM COLOR) TO BOTTOM RIGHT
    Picture2.Top = Me.Height - 2800
    Picture2.Left = Me.Width - 4000
    
    'RECOLORING OF BACKGROUND TO SUIT RESIZE
    bgcolor Me, slidered(0).Value, slidegreen(0).Value, slideblue(0).Value, slidered(1).Value, slidegreen(1).Value, slideblue(1).Value
End Sub

Private Sub slideblue_Change(Index As Integer)
    'RECOLORING OF BACKGROUND
    If Index = 0 Then
        Label1.BackColor = RGB(slidered(0).Value, slidegreen(0).Value, slideblue(0).Value)
    Else
        Label11.BackColor = RGB(slidered(1).Value, slidegreen(1).Value, slideblue(1).Value)
    End If
    bgcolor Me, slidered(0).Value, slidegreen(0).Value, slideblue(0).Value, slidered(1).Value, slidegreen(1).Value, slideblue(1).Value
End Sub

Private Sub slidegreen_Change(Index As Integer)
    'RECOLORING OF BACKGROUND
    If Index = 0 Then
        Label1.BackColor = RGB(slidered(0).Value, slidegreen(0).Value, slideblue(0).Value)
    Else
        Label11.BackColor = RGB(slidered(1).Value, slidegreen(1).Value, slideblue(1).Value)
    End If
    bgcolor Me, slidered(0).Value, slidegreen(0).Value, slideblue(0).Value, slidered(1).Value, slidegreen(1).Value, slideblue(1).Value
End Sub

Private Sub slidered_Change(Index As Integer)
    'RECOLORING OF BACKGROUND
    If Index = 0 Then
        Label1.BackColor = RGB(slidered(0).Value, slidegreen(0).Value, slideblue(0).Value)
    Else
        Label11.BackColor = RGB(slidered(1).Value, slidegreen(1).Value, slideblue(1).Value)
    End If
    bgcolor Me, slidered(0).Value, slidegreen(0).Value, slideblue(0).Value, slidered(1).Value, slidegreen(1).Value, slideblue(1).Value
End Sub

