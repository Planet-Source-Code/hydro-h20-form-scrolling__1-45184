VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   3750
   ClientTop       =   1545
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   9855
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   3120
      TabIndex        =   12
      Top             =   480
      Width           =   2895
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1095
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   2295
         Begin VB.OptionButton Option3 
            Caption         =   "Option3"
            Height          =   255
            Left            =   360
            TabIndex        =   17
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Text            =   "Text3"
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3975
      LargeChange     =   200
      Left            =   9600
      Max             =   2836
      SmallChange     =   50
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   10
      Top             =   5040
      Width           =   2295
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      Tag             =   "Vfixed"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Tag             =   "Vfixed"
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastVValue As Integer

Private Sub VScroll1_Change()
    Dim vDIFF As Integer
    Dim i As Integer
    
    On Local Error GoTo CTest
    
    If LastVValue = VScroll1.Value Then Exit Sub
    
    vDIFF = VScroll1.Value - LastVValue
    LastVValue = VScroll1.Value
    
    For i = 0 To Me.Controls.Count - 1
        If Me.Controls(i).Name <> "VScroll1" Then
            If Len(Me.Controls(i).Container) = 0 Then
NoContainer:
                If Me.Controls(i).Tag <> "Vfixed" Then
                    Me.Controls(i).Top = Me.Controls(i).Top - vDIFF
                End If
            End If
        End If
    Next i
    DoEvents
    Exit Sub

CTest:
    Resume NoContainer
End Sub
