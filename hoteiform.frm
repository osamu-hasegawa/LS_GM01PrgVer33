VERSION 5.00
Begin VB.Form hoteiform 
   Caption         =   "hoteiform"
   ClientHeight    =   2268
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4212
   LinkTopic       =   "Form1"
   ScaleHeight     =   2268
   ScaleWidth      =   4212
   StartUpPosition =   3  'Windows Ìùèl
   Begin VB.CommandButton Command1 
      Caption         =   "Û·Eâ~@I¹"
      Height          =   252
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   2772
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3720
      Top             =   840
   End
   Begin VB.Label Label3 
      Alignment       =   2  'µ¦
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "lr oSVbN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   2532
   End
   Begin VB.Label Label2 
      Alignment       =   2  'µ¦
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "lr oSVbN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'µ¦
      Caption         =   "Û·Eâ~@oßÔ"
      BeginProperty Font 
         Name            =   "lr oSVbN"
         Size            =   13.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3732
   End
End
Attribute VB_Name = "hoteiform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  DispCenter Me
  Label3.Caption = "Û·Eâ~@Jn" + Time$
End Sub

Private Sub Timer1_Timer()
   hyouji
End Sub
Private Sub hyouji()
Dim imachi%
Dim fintime!, evtime!, difft!
'
 
'  @Åå@Û·â~@Ô@Ýè@-----------------------------
     evtime = Timer
        imachi = 60 * 60 - 1
'@---@Ò¿Ô\¦
     Do
       DoEvents
'       fintime = Timer2func   'NQDÍ@Timer2func@gp
       fintime = Timer
       difft = diffTime(fintime, evtime)
       If (difft < imachi) Then
          If (Int(difft) <> Int(diffTold)) Then
             Label2.Caption = " oßÔ " + Format(Int(difft / 60), "0ª") + Format(Int(difft) Mod 60, " 0b")
              diffTold = difft
          End If
       Else
          Exit Do              '@Û·Eâ~@I¹
       End If
     Loop

Unload Me
'
End Sub
