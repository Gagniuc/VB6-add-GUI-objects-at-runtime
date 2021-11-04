VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GUI objects at runtime (Ex.1)"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16140
   LinkTopic       =   "Form1"
   ScaleHeight     =   641
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1076
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox v 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   10920
      TabIndex        =   9
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Option"
      Height          =   1455
      Left            =   7080
      TabIndex        =   7
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox RND_Values 
         Caption         =   "Fill random values"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6855
      Begin VB.HScrollBar square_M 
         Height          =   255
         Left            =   240
         Max             =   20
         Min             =   1
         TabIndex        =   6
         Top             =   1080
         Value           =   1
         Width           =   5175
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   20
         Min             =   1
         TabIndex        =   3
         Top             =   360
         Value           =   5
         Width           =   5175
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   240
         Max             =   20
         Min             =   1
         TabIndex        =   2
         Top             =   720
         Value           =   5
         Width           =   5175
      End
      Begin VB.Label MY_ROW 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   5520
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label MX_COL 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   5520
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Timer Real_time_matrix 
      Interval        =   10
      Left            =   9720
      Top             =   120
   End
   Begin VB.TextBox M 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   10320
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ________________________________                          ____________________
'  /            Matrix              \________________________/       v1.00        |
' |                                                                               |
' |            Name:  Matrix V1.0                                                 |
' |        Category:  Open source software                                        |
' |          Author:  Paul A. Gagniuc                                             |
' |           Email:  paul_gagniuc@acad.ro                                        |
' |  ____________________________________________________________________________ |
' |                                                                               |
' |    Date Created:  September 2013                                              |
' |       Tested On:  WinXP, WinVista, Win7, Win8                                 |
' |             Use:  generate objects at runtime                                  |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|
'

Option Explicit


Dim Old_HSX As Integer
Dim Old_HSY As Integer

Dim MX_pos As Variant
Dim MY_pos As Variant

Dim space_x As Integer
Dim space_y As Integer

Private Sub Form_Load()
    MX_pos = 50
    MY_pos = 100
    
    space_x = 3
    space_y = 3
End Sub


Function Load_M(ByVal X As Integer, ByVal Y As Integer)
Dim c, r, a, i As Integer
Dim xx, yy As Variant

For i = 1 To M.UBound
    Unload M(i)
Next i


a = 0

For r = 1 To Y
    
    For c = 1 To X

        a = a + 1

        Load M(a)
        
        xx = MX_pos + ((M(0).Width + space_x) * c)
        yy = MY_pos + ((M(0).Height + space_y) * r)

        M(a).Move xx, yy
        M(a).Visible = True

        If RND_Values.Value = 1 Then M(a).Text = Round(Rnd(10), 2)

        M(a).Refresh

    Next c

Next r
End Function


Private Sub Real_time_matrix_Timer()
Dim MXX, MYY As Integer

MXX = HScroll1.Value + square_M.Value
MYY = HScroll2.Value + square_M.Value

If Old_HSX <> MXX Or Old_HSY <> MYY Then
    Call Load_M(MXX, MYY)
    Old_HSX = MXX
    Old_HSY = MYY
    
    MX_COL.Caption = Old_HSX & " cols"
    MY_ROW.Caption = Old_HSY & " rows"
End If

End Sub


