VERSION 5.00
Begin VB.Form LS21_SC 
   Appearance      =   0  'ﾌﾗｯﾄ
   BackColor       =   &H00C0C0C0&
   Caption         =   "連続成形"
   ClientHeight    =   8532
   ClientLeft      =   132
   ClientTop       =   420
   ClientWidth     =   11844
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8532
   ScaleWidth      =   11844
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Scr.Copy"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   10920
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   123
      Top             =   8160
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   11340
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   122
      Top             =   2280
      Width           =   500
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00E0E0E0&
      Caption         =   "型順"
      ForeColor       =   &H80000008&
      Height          =   1572
      Left            =   10250
      TabIndex        =   110
      Top             =   2760
      Width           =   1575
      Begin VB.Line Line3 
         X1              =   0
         X2              =   1560
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label13 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   600
         TabIndex        =   119
         Top             =   550
         Width           =   396
      End
      Begin VB.Label Label13 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '実線
         Caption         =   "Label13"
         Height          =   240
         Index           =   7
         Left            =   600
         TabIndex        =   118
         Top             =   1320
         Width           =   396
      End
      Begin VB.Label Label13 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "Label13"
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   117
         Top             =   1030
         Width           =   396
      End
      Begin VB.Label Label13 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  '実線
         Caption         =   "Label13"
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   116
         Top             =   720
         Width           =   396
      End
      Begin VB.Label Label13 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  '実線
         Caption         =   "Label13"
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   115
         Top             =   400
         Width           =   396
      End
      Begin VB.Label Label13 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  '実線
         Caption         =   "Label13"
         Height          =   240
         Index           =   3
         Left            =   600
         TabIndex        =   114
         Top             =   150
         Width           =   396
      End
      Begin VB.Label Label13 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  '実線
         Caption         =   "Label13"
         Height          =   240
         Index           =   2
         Left            =   1080
         TabIndex        =   113
         Top             =   150
         Width           =   396
      End
      Begin VB.Label Label13 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '実線
         Caption         =   "Label13"
         Height          =   240
         Index           =   1
         Left            =   1080
         TabIndex        =   112
         Top             =   400
         Width           =   396
      End
      Begin VB.Label Label13 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '実線
         Caption         =   "Label13"
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   111
         Top             =   1030
         Width           =   400
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "5分停止"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   2520
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   106
      Top             =   100
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "保温停止"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   9
      Left            =   3240
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   95
      Top             =   100
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GraphDataSave"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   57
      Top             =   480
      Width           =   1440
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0C0&
      Height          =   768
      Left            =   1920
      TabIndex        =   77
      Top             =   1080
      Width           =   8280
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "強制ソーク"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   8
      Left            =   1800
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   59
      Top             =   100
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "真空到達"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   0
      TabIndex        =   56
      Top             =   2040
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "V エディタ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   120
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   54
      Top             =   840
      Width           =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1920
      Top             =   4200
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   5500
      Left            =   1780
      ScaleHeight     =   5472
      ScaleWidth      =   8376
      TabIndex        =   8
      Top             =   1880
      Width           =   8400
      Begin VB.ListBox List2 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   744
         Left            =   0
         TabIndex        =   93
         Top             =   300
         Width           =   8292
      End
      Begin VB.Label Label14 
         BackStyle       =   0  '透明
         Caption         =   "Label14"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   0
         TabIndex        =   124
         Top             =   1060
         Width           =   8000
      End
      Begin VB.Label Label10 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H00800000&
         BackStyle       =   0  '透明
         Caption         =   "Label10"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   0
         TabIndex        =   94
         Top             =   120
         Width           =   7455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   7
         X1              =   6696
         X2              =   6696
         Y1              =   0
         Y2              =   6436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   6
         X1              =   5040
         X2              =   5040
         Y1              =   0
         Y2              =   6436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   5
         X1              =   3348
         X2              =   3348
         Y1              =   0
         Y2              =   6436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   4
         X1              =   1656
         X2              =   1656
         Y1              =   0
         Y2              =   6436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   3
         X1              =   -360
         X2              =   7992
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   2
         X1              =   0
         X2              =   8352
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   1
         X1              =   0
         X2              =   8352
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '点線
         Index           =   0
         X1              =   0
         X2              =   8352
         Y1              =   4320
         Y2              =   4320
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "終了"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   120
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   5
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "label1(6)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   6
      Left            =   10320
      TabIndex        =   121
      Top             =   2400
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "回数指定："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   10320
      TabIndex        =   120
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   10250
      TabIndex        =   109
      Top             =   2520
      Width           =   1572
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   10250
      TabIndex        =   108
      Top             =   2280
      Width           =   1572
   End
   Begin VB.Label Label12 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   10250
      TabIndex        =   107
      Top             =   2040
      Width           =   1572
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   9
      Left            =   10240
      TabIndex        =   105
      Top             =   7160
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   204
      Index           =   8
      Left            =   10240
      TabIndex        =   104
      Top             =   6884
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   7
      Left            =   10240
      TabIndex        =   103
      Top             =   6596
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   6
      Left            =   10240
      TabIndex        =   102
      Top             =   6320
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   5
      Left            =   10240
      TabIndex        =   101
      Top             =   6044
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   204
      Index           =   4
      Left            =   10240
      TabIndex        =   100
      Top             =   5756
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   3
      Left            =   10240
      TabIndex        =   99
      Top             =   5480
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   2
      Left            =   10240
      TabIndex        =   98
      Top             =   5204
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   204
      Index           =   1
      Left            =   10240
      TabIndex        =   97
      Top             =   4916
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   0
      Left            =   10240
      TabIndex        =   96
      Top             =   4640
      Width           =   200
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   280
      Left            =   4000
      TabIndex        =   92
      Top             =   90
      Width           =   1200
   End
   Begin VB.Label Label9 
      Caption         =   "  Z3補正"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   1
      Left            =   11170
      TabIndex        =   91
      Top             =   4400
      Width           =   580
   End
   Begin VB.Label Label9 
      Caption         =   "  Ｔ係数"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   0
      Left            =   10460
      TabIndex        =   90
      Top             =   4400
      Width           =   612
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   0
      Left            =   11200
      TabIndex        =   89
      Top             =   4640
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   1
      Left            =   11200
      TabIndex        =   88
      Top             =   4916
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   2
      Left            =   11200
      TabIndex        =   87
      Top             =   5204
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   3
      Left            =   11200
      TabIndex        =   86
      Top             =   5480
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   4
      Left            =   11200
      TabIndex        =   85
      Top             =   5756
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   5
      Left            =   11200
      TabIndex        =   84
      Top             =   6044
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   6
      Left            =   11200
      TabIndex        =   83
      Top             =   6320
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   7
      Left            =   11200
      TabIndex        =   82
      Top             =   6596
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   8
      Left            =   11200
      TabIndex        =   81
      Top             =   6884
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   9
      Left            =   11200
      TabIndex        =   80
      Top             =   7160
      Width           =   540
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   216
      Index           =   1
      Left            =   11200
      TabIndex        =   79
      Top             =   7500
      Width           =   540
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   216
      Index           =   0
      Left            =   10500
      TabIndex        =   78
      Top             =   7500
      Width           =   540
   End
   Begin VB.Label Label5 
      Caption         =   "cc3-2"
      Height          =   252
      Index           =   6
      Left            =   10320
      TabIndex        =   76
      Top             =   1440
      Width           =   1380
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   9
      Left            =   10500
      TabIndex        =   75
      Top             =   7160
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   8
      Left            =   10500
      TabIndex        =   74
      Top             =   6884
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   7
      Left            =   10500
      TabIndex        =   73
      Top             =   6596
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   6
      Left            =   10500
      TabIndex        =   72
      Top             =   6320
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   5
      Left            =   10500
      TabIndex        =   71
      Top             =   6044
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   4
      Left            =   10500
      TabIndex        =   70
      Top             =   5756
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   3
      Left            =   10500
      TabIndex        =   69
      Top             =   5480
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   216
      Index           =   2
      Left            =   10500
      TabIndex        =   68
      Top             =   5204
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   216
      Index           =   1
      Left            =   10500
      TabIndex        =   67
      Top             =   4916
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   0
      Left            =   10500
      TabIndex        =   66
      Top             =   4640
      Width           =   540
   End
   Begin VB.Label Label5 
      Caption         =   "cc3"
      Height          =   252
      Index           =   5
      Left            =   8640
      TabIndex        =   65
      Top             =   4560
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "cc2"
      Height          =   252
      Index           =   4
      Left            =   8640
      TabIndex        =   64
      Top             =   4080
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "cc1"
      Height          =   252
      Index           =   3
      Left            =   8640
      TabIndex        =   63
      Top             =   3480
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "ct2"
      Height          =   252
      Index           =   2
      Left            =   8640
      TabIndex        =   62
      Top             =   3120
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "ct1"
      Height          =   252
      Index           =   1
      Left            =   10320
      TabIndex        =   61
      Top             =   1100
      Width           =   1380
   End
   Begin VB.Label Label5 
      Caption         =   "cp1"
      Height          =   252
      Index           =   0
      Left            =   10320
      TabIndex        =   60
      Top             =   1780
      Width           =   1380
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   14
      Left            =   6720
      TabIndex        =   58
      Top             =   7800
      Width           =   4980
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   10920
      TabIndex        =   55
      Top             =   75
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "コマンド："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   11
      Left            =   120
      TabIndex        =   53
      Top             =   8160
      Width           =   1290
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   1428
      TabIndex        =   52
      Top             =   8160
      Width           =   5040
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   6720
      TabIndex        =   51
      Top             =   8160
      Width           =   4140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "ショット数："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   8100
      TabIndex        =   50
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "サイクルタイム："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   8400
      TabIndex        =   49
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   9840
      TabIndex        =   48
      Top             =   75
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   10200
      TabIndex        =   47
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   1440
      TabIndex        =   46
      Top             =   7800
      Width           =   5040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "コマンド："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   45
      Top             =   7800
      Width           =   1290
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   8400
      TabIndex        =   44
      Top             =   780
      Width           =   3312
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   5280
      TabIndex        =   43
      Top             =   3360
      Visible         =   0   'False
      Width           =   4872
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   3480
      TabIndex        =   42
      Top             =   3480
      Visible         =   0   'False
      Width           =   3432
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "成形状態："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   228
      Index           =   1
      Left            =   2040
      TabIndex        =   41
      Top             =   3240
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(分)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   31
      Left            =   9360
      TabIndex        =   40
      Top             =   7560
      Width           =   465
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "経過時間"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   30
      Left            =   7275
      TabIndex        =   39
      Top             =   7560
      Width           =   870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   27
      X1              =   10200
      X2              =   10200
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   26
      X1              =   8520
      X2              =   8520
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   25
      X1              =   6840
      X2              =   6840
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   24
      X1              =   5160
      X2              =   5160
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   23
      X1              =   3480
      X2              =   3480
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   29
      Left            =   9930
      TabIndex        =   38
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   28
      Left            =   8355
      TabIndex        =   37
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   27
      Left            =   6660
      TabIndex        =   36
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   26
      Left            =   4965
      TabIndex        =   35
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   25
      Left            =   3270
      TabIndex        =   34
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   24
      Left            =   1650
      TabIndex        =   33
      Top             =   7485
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   22
      X1              =   1800
      X2              =   1800
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   21
      X1              =   10200
      X2              =   1800
      Y1              =   7380
      Y2              =   7380
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "型温度"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Index           =   23
      Left            =   1230
      TabIndex        =   32
      Top             =   1260
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(℃)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   216
      Index           =   22
      Left            =   1200
      TabIndex        =   31
      Top             =   1515
      Width           =   468
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   20
      X1              =   1620
      X2              =   1764
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   19
      X1              =   1620
      X2              =   1764
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   18
      X1              =   1680
      X2              =   1824
      Y1              =   4056
      Y2              =   4056
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   17
      X1              =   1620
      X2              =   1764
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   16
      X1              =   1620
      X2              =   1764
      Y1              =   6270
      Y2              =   6270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   15
      X1              =   1620
      X2              =   1764
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   14
      X1              =   1776
      X2              =   1776
      Y1              =   1856
      Y2              =   7380
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "####"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   216
      Index           =   21
      Left            =   1200
      TabIndex        =   30
      Top             =   1800
      Width           =   492
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Index           =   20
      Left            =   1290
      TabIndex        =   29
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Index           =   19
      Left            =   1290
      TabIndex        =   28
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   216
      Index           =   18
      Left            =   1320
      TabIndex        =   27
      Top             =   5076
      Width           =   372
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Index           =   17
      Left            =   1290
      TabIndex        =   26
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Index           =   16
      Left            =   1290
      TabIndex        =   25
      Top             =   7230
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "型締圧"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   15
      Left            =   540
      TabIndex        =   24
      Top             =   1260
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(kg)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   216
      Index           =   14
      Left            =   600
      TabIndex        =   23
      Top             =   1512
      Width           =   492
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   13
      X1              =   1005
      X2              =   1149
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   12
      X1              =   1005
      X2              =   1149
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   11
      X1              =   1005
      X2              =   1149
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   10
      X1              =   1005
      X2              =   1149
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   9
      X1              =   1005
      X2              =   1149
      Y1              =   6270
      Y2              =   6270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   8
      X1              =   1005
      X2              =   1149
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   7
      X1              =   1152
      X2              =   1152
      Y1              =   1856
      Y2              =   7356
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   13
      Left            =   645
      TabIndex        =   22
      Top             =   1770
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   12
      Left            =   645
      TabIndex        =   21
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   11
      Left            =   645
      TabIndex        =   20
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   10
      Left            =   645
      TabIndex        =   19
      Top             =   5070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   9
      Left            =   645
      TabIndex        =   18
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   8
      Left            =   645
      TabIndex        =   17
      Top             =   7230
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "座標"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   7
      Left            =   30
      TabIndex        =   16
      Top             =   1260
      Width           =   450
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(mm)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   6
      Left            =   30
      TabIndex        =   15
      Top             =   1515
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   6
      X1              =   390
      X2              =   534
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   5
      X1              =   390
      X2              =   534
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   4
      X1              =   390
      X2              =   534
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   3
      X1              =   390
      X2              =   534
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   2
      X1              =   390
      X2              =   534
      Y1              =   6270
      Y2              =   6270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   1
      X1              =   390
      X2              =   534
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   0
      X1              =   540
      X2              =   540
      Y1              =   1856
      Y2              =   7356
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   5
      Left            =   30
      TabIndex        =   14
      Top             =   1770
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   4
      Left            =   30
      TabIndex        =   13
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   3
      Left            =   30
      TabIndex        =   12
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   2
      Left            =   30
      TabIndex        =   11
      Top             =   5070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   1
      Left            =   30
      TabIndex        =   10
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   0
      Left            =   30
      TabIndex        =   9
      Top             =   7230
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "コメント："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   3240
      TabIndex        =   6
      Top             =   780
      Width           =   4930
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3960
      TabIndex        =   4
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "制御ファイル名："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   7
      Left            =   1950
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "分"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   7695
      TabIndex        =   2
      Top             =   130
      Width           =   225
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   0
      Left            =   6600
      TabIndex        =   1
      Top             =   72
      Width           =   1000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "測定時間："
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   3
      Left            =   5400
      TabIndex        =   0
      Top             =   90
      Width           =   1275
   End
End
Attribute VB_Name = "LS21_SC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    LS21_SC
'            update: 2002.6.28 s.f  private sub cal_pid　削除
'            update: 2002.6.28 s.f difftime　書き換え
'            update: 2002.7.10 s.f "DC","HC" 新規追加
'            update: 2002.8.10 s.f roz(0),roz(1)を突当成形のﾊﾟﾗﾒｰﾀへ max.180
'            update: 2002.8.15 s.f Veditcol 追加
'            update: 2002.8.18 s.f タクトタイム表示 int(stime/60)へ
'                                  "HC" 修正し、完成へ
'                                  "DC" 時　成形回数　戻し（i_s=i_s-1)
'
'            update: 2002.8.22 s.f 座標軸　黄色へ
'            update: 2002.8.24 s.f 暫定変更　「VEditが　毎回必ず入る」
'            update: 2002.8.25 s.f 成形回数　save　　InitDTsave　を　成形後へ移動
'            update: 2002.8.29 s.f cp,ct,ccデータ表示'
'            update: 2002.9.06 s.f 成形回数　表示　idcflg追加
'            update: 2002.9.26 s.f ic(i)=10 で　終了判断　に　訂正
'            update: 2002.10.1 s.f 軸制御モード２へ、　CtlDisp  'DioOut 12,1  位置制御 '  02.10.1 追加
'            update: 2002.10.1 s.f 軸制御　エラー表示　Label2(4)からLabel2(3)へ変更
'            update: 2002.10.2 s.f 軸制御スタート時間表示
'            update: 2002.10.5 s.f タイムアップルーチン見直し（ｾｸﾞﾒﾝﾄ飛び対策）
'            update: 2002.10.5 s.f 時間表示変更
'            update: 2002.10.9 KYOCERA タイマー処理、タイムアップ、コメント表示、時間表示変更
'            update: 2002.10.12 s.f ﾀｲﾑｱｯﾌﾟの成立後　goto文　変更
'            update: 2002.10.16 KYOCERA ﾀｲﾑｱｯﾌﾟ処理 <9 を istend に変更
'            update: 2002.10.16 KYOCERA ﾀｲﾑｱｯﾌﾟで次のｽﾃｯﾌﾟ追加
'            update: 2002.10.17 KYOCERA 原点復帰後に初回原点復帰完了ﾌﾗｸﾞgOrgStartFlgをON
'            update: 2002.10.17 KYOCERA ﾀｲﾑｱｯﾌﾟ処理 <istend を 10 に変更
'            update: 2002.10.26 s.f 軸制御　エラー表示　Label2(3)からLabel2(5)へ変更
'            　　　　　　　　　　s.f cc3-cc2表示追加
'                                   SR　の処理変更　0.1秒に１回ｻﾝﾌﾟﾘﾝｸﾞ
'            update: 2002.11.28 s.f 終了受付・解除　変更　（解除可能にする）
'            update: 2002.12.03 s.f 成形記録の表示・ディスク記録　追加
'            update: 2002.12.05 s.f 成形記録の表示・ディスク記録　修正
'            update: 2003.03.22 s.f CTコマンド　誤記訂正　ct=  -> ct_temp(  へ
'            update: 2003.07.10 HND アラーム表示中の　成形プログラム続行
'                                  frmerr_sign, FbiDio, LS21_SC
'            update: 2004. 3. 8 s.f. LS21_SC 変更　成形軸制御モード　’７’追加　（上軸衝突判定付）
'                                    RecEmgDTsave 非常停止メッセージの保存
'
'            update: 2004. 3.12 s.f.  速度指令電圧　表示
'            update: 2004. 4.23 s.f.  timeupで　非常停止
'            update: 2004. 5. 5 s.f   温度係数、肉厚補正ルーチン　追加  PGM_KTD,My_lib,MYEDIT, LS21_SC, LS21_TC
'            update: 2004.5.12  s.f   PGM_KTD　"ｵｰﾊﾞｰﾌﾛｰ"対策　　wTm0!,wTm1!  global化,  LS21_SCと　LS21_TC から　dim削除
'            update: 2004.5.17  s.f   'S'ｺﾏﾝﾄﾞ　バグ対策
'            update: 2004.5.18  s.f    T係数表示
'            update: 2004.8.17  s.f   ｵｰﾊﾞｰﾌﾛｰ"対策  p(ist0)をppへ  ”：”複数の行を無くす
'                                     List1.Enabled = True or False 追加
'            update: 2004.8.27 - 10.30  s.f   T係数関数変更、　　「ＤＣ　０」コマンド　成形前に型在否チェックセンサーのチェック機能追加
'            update: 2005. 5.25 s.f    Version No表示追加
'            update: 2005. 7.18 s.f    加圧時間　平均値表示
'            update: 2005. 7.25 s.f   加圧時間制御デバッグ    List2.Enabled = True or False 追加
'            update: 2005. 9.27 s.f    保温停止モード追加  成形終了時　軸が下がらずに保温して停止
'            update: 2005. 9.28 s.f   T係数　表示色変更
'            update: 2005.11. 4 s.f 　 LS21_SC　表示変更。速度制御電圧表示削除。T係数、Z３補正表示部変更、　加圧時間制御バグ修正
'            update: 2005.11.22 s.f   Melec C-870 counter動作バグ修正　コンペアカウンタ値セット時　符号反転　　setcm1
'                                     C870sts(3) 周り　バグ修正、右横データ順序変更
'            update: 2005.11.23 s.f   11/22 変更のバグ修正　成形軸制御　「C870sts　resetするまで　読み飛ばす」を　復活
'　　　　　　　　　　　　　　　　　　画面下表示　シンプル化　（スピード低下防止の為）
'            update: 2005.11.26 s.f   すべての　function　に　型宣言をつける　　　overflow対策
'            update: 2005.12.17 s.f   Do-Loop 外の　DoEvent削除 OverFlow 対策 s.f.
'                                     コマンドの　evtime　取り込みを　コマンド開始時へ変更
'　　　　　　　　　　　　　　　　　　DCコマンド　LAコマンド　再チェック修正
'　　　　　　　　　　　　　　　　　　連続前コマンド　evtime　と　fintime　表記入れ替え
'            update: 2005.12.23 s.f
'            update: 2006. 2.18 s.f
'            update: 2006. 3. 3 s.f  edit 使用時　do　loopから抜ける
'　　　　　　　　　　　　　　　　　　DCｺﾏﾝﾄﾞへ　fintime=timer　を　設置
'            update: 2006. 4.14 s.f  on error goto,  sts as long
'            update: 2006. 4.15 s.f  error 表示
'            update: 2006. 5. 9 s.f  O.F.error 表示　軸制御　end3　追加,  tstime=0#
'            update: 2006. 5.14 s.f 　r_pres()の　DoEvents 　 forの外へ移動　s.f  ものすごく効く
'　　　　　　　　　　　　　　　　　  すべて抜くと　LS_TC　プログラム暴走する（LS_SCは　OK)’
'            update: 2006. 5.15 s.f  5分間保温停止　追加
'            update: 2006. 5.18 s.f 　r_pres()の　DoEvents 　削除、　”J"、１秒に1回　Doevents　追加
'                                    非常停止　表示追加
'            update: 2006. 7.12 s.f  加圧時間自動調整　’有効’へ
'
'       Ver.3.33R_061221 2006.12.21 s.f  LS-33改　対応　　VacuumON、VacuumOFF　を廃止、SeikeiON,SeikeiOFF新設　DO3　割り当て変更
'       Ver.3.33R_070827 2007.08.27 s.f  非常停止時の　処置追加
'       Ver.3.33R_070927 2007.09.27 s.f  Z補正　指定したｾｸﾞﾒﾝﾄNo.へ　できるようにする
'       Ver.3.33R_071112 2007.11.13 s.f  「強制ソーク」復活、　「1回成形」enable=Falseへ
'       Ver.3.33R_071119 2007.11.19 s.f  加圧時間制御　バグ修正（edit時、データ継承）、平均値AND最新値で　更新判定へ
'       Ver.3.33R_071120 2007.11.20 s.f  バグ修正、　空成形-排出　追加、　連続成形再開　追加
'       Ver.3.33R_071121 2007.11.21 s.f  加圧制御　平均値計算　今回の加圧時間　重み2.0へ
'       Ver.3.33R_071122 2007.11.22 s.f  型順　表示バグ修正
'       Ver.3.33R_071127 2007.11.27 s.f  型順　表示ポインター式へ変更
'       Ver.3.33R_071210 2007.12.10 s.f  終了時　T係数を格納して　終了する様変更（　save　追加　）
'       Ver.3.33R_080817A 2008. 7.17 s.f  型数　変更版へ katano
'       Ver.3.33-12R-100304R 2010. 3. 4 s.f  初回ポインターカウントアップバグ対策　　成形有効無効判断から　i_s=0のときpassを削除。
'       Ver.3.33-12R-100310R 2010. 3.10 s.f  Timer値異常時　timeup処理の　skip実施。
'       Ver.3.33-12R-100412R 2010. 4.12 s.f  timeup処理の　skip ifに　ic(ist0)<10を追加。ﾎﾟｲﾝﾀｰｶｳﾝﾄｱｯﾌﾟのバグ修正のバグ修正、「成形の有効性判定」を　初回は別枠にする。
'       Ver.3.33-12R-100416R 2010. 4.16 s.f  成形回数指定の追加。 seikeiKaisu
'           '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Option Explicit
Dim lGphNo%
Dim lGphNo0%
Dim EditFlg As Long
Dim lViewFlg      '前の画面番号
Dim NextView%
Dim NextViewBUp%  'NextViewの内容backup
Dim lDtSaveFlg%   'データ保存
Dim iDtSaveCount%  'データ保存回数　　max=14　20190428追加
Dim idcflg%(0 To 3)        ' DCフラグ　形無=1　型有=0
Dim SokuCor!(0 To 1)  '強制ソークタイムのコマンド釦の色
Dim TKatBackCol!(0 To 1)  '加圧時間補正　上限加減　表示のbackColor
Dim lEmgFlg As Long       '非常停止
Dim iflghoonStop As Long, iHoonStopNo As Long  '保温停止フラグ、保温停止回数カウンター
Dim iflg5Stop As Long    '5分間保温停止フラグ
Dim iHoteishuryo As Long
Dim iflgSCopy As Boolean   ' ScreenCopy フラグ
'スクリーンのスナップショットをクリップボードに保存及び印刷　　変数宣言部　　（273） '

Private Declare Sub keybd_event Lib "user32.dll" _
        (ByVal bVk As Byte, ByVal bScan As Byte, _
         ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Const VK_SNAPSHOT = &H2C            'PrintScreen キー(P1051)
Private Const VK_LMENU = &HA4               'Altキー
Private Const KEYEVENTF_KEYUP = &H2         'キーはアップ状態
Private Const KEYEVENTF_EXTENDEDKEY = &H1   'スキャンコードは拡張コード



Private Sub Command1_Click()
    If iflghoonStop = True Then
        iHoteishuryo = 1
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
'Case 0  'キャンセル
'  lGphNo = 0
'  MoniGraph Me.Picture1, 0, lGphNo
Case 1  '終了
   If FrmMenuFlg = True Then
          FrmMenuFlg = False          '終了受付
          NextViewBUp = NextView
          NextView = 1
          Command2(1).BackColor = CmndColon(1)
    Else
          FrmMenuFlg = True           '終了受付解除
          NextView = NextViewBUp
          Command2(1).BackColor = CmndColoff(1)
  End If
'Case 2  'グラフ再描画
'  lGphNo = lGphNo + 100
'  MoniGraph Me.Picture1, 0, lGphNo
''
Case 2
'''アクティブウインドウをクリップボードにコピー印刷する。　True に設定
  If iflgSCopy = True Then
          iflgSCopy = False          'ScreenCopy　受付解除
          Command2(2).BackColor = CmndColoff(1)
    Else
          iflgSCopy = True      'ScreenCopy　受付
          Command2(2).BackColor = CmndColon(1)    ' on 1=red
''        Call SaveWindowPic(True, False)     'Active Windowの保存
  End If
Case 3                        'edit　の　'02/8暫定変更(s.f)
  If EditFlg = True Then
          EditFlg = False          'エディタ起動解除
          Command2(3).BackColor = CmndColoff(3)
    Else
          EditFlg = True      'エディタ起動
          Command2(3).BackColor = CmndColon(1)   ' 1=red
  End If
'
Case 4      '真空到達
  gVumFlg = 1
'真空到達=1
'
Case 5      '"Save" ;データセーブ
'
    If lDtSaveFlg = True Then
          lDtSaveFlg = False          'データセーブ　受付解除
          Command2(5).BackColor = CmndColoff(1)    ' off gray
          Command2(5).Caption = "GraphDataSave"
    Else
          lDtSaveFlg = True           'データセーブ　受付
          Command2(5).BackColor = CmndColon(1)   ' on 1= red
          Command2(5).Caption = "DataSave中"
          iDtSaveCount = 14
  End If
'
Case 8      '強制ソークタイム
  If lSokuFlg = True Then
          lSokuFlg = False          '強制ソークタイム　受付解除
          Command2(8).BackColor = SokuCor(0)
    Else
          lSokuFlg = True           '強制ソークタイム　受付
          Command2(8).BackColor = SokuCor(1)
  End If
Case 9     '保温停止  成形終了時　軸が下がらずに保温して停止
  If iflghoonStop = True Then
          iflghoonStop = False          '保温停止　受付解除
          iHoteishuryo = 1
          Command2(9).BackColor = CmndColoff(9)
    Else
          iflghoonStop = True      '保温停止　受付
          iHoteishuryo = 0
          Command2(9).BackColor = CmndColon(1)    ' on 1=red
          iflg5Stop = False        '5分間保温停止　受付解除
          Command2(0).BackColor = CmndColoff(0)
  End If
  If (KataChk() < 3) Then  '型が無い
          iflghoonStop = False          '保温停止　受付解除
          Command2(9).BackColor = CmndColoff(9)
  End If
Case 0     '5分間保温停止
  If iflg5Stop = True Then
          iflg5Stop = False          '5分間保温停止　受付解除
          Command2(0).BackColor = CmndColoff(0)
    Else
          iflg5Stop = True      '5分間保温停止　受付
          Command2(0).BackColor = CmndColon(1)    ' on 1=red
          iflghoonStop = False  '保温停止　受付解除
          Command2(9).BackColor = CmndColoff(9)
  End If
  If (KataChk() < 3) Then  '型が無い
          iflg5Stop = False          '5分間保温停止　受付解除
          Command2(0).BackColor = CmndColoff(0)
  End If
'
End Select
DoEvents
End Sub

Private Sub SetData()
  Label2(0) = Format(ptime, "###0")  '測定時間
  Label2(2) = gcoxFlName             '制御ファイル名
  Label2(3) = hcomm(2)               'コメント
' -----------------------------------
  DispGphScale
End Sub

Private Sub Form_Load()
  DispCenter Me
  LS21_SC.Caption = LS21_SC.Caption + "     " + versionNo
  Me.Top = 0
  SokuCor(0) = &H8000000F     '強制ソークタイムのコマンド釦の色
  SokuCor(1) = &HFF&          '強制ソークタイムのコマンド釦の色 押されたとき
  lDtSaveFlg = False      'データ保存
  iDtSaveCount = 0        'データ保存回数　初期値=0
'
  If lSokuFlg = False Then
          Command2(8).BackColor = SokuCor(0)
    Else
          Command2(8).BackColor = SokuCor(1)
  End If
  lViewFlg = ViewFlg      '前の画面番号
  ViewFlg = 2             '画面番号
  FrmMenuFlg = True       'メニューから抜けるときfalse
  EditFlg = False        'エディタ起動解除
  Command2(1).BackColor = CmndColoff(1)     '終了コマンド釦の色
  Command2(3).BackColor = CmndColoff(3)     'Vエディトのコマンド釦の色
  Command2(9).BackColor = CmndColoff(9)     '保温停止コマンド釦の色
    TKatBackCol!(0) = &H8000000F      '加圧制御　ＯＦＦのとき
    TKatBackCol!(1) = &HC0C0FF      '加圧制御　ＯＮのとき
    lEmgFlg = False         '非常停止
  SetData
  Timer1.Enabled = True
  iflghoonStop = False
  iHoonStopNo = 0
End Sub


Private Sub DispGphScale()
Dim i%, p%, max!, min!, def!, dev%
  '
  GphXSet           '時間軸の時間をセット
  '
  dev = 5
  '
  min = InitDat(1)  'グラフスケール座標 (Min)
  max = InitDat(2)  'グラフスケール座標 (Max)
  def = (max - min) / dev
  p = 0
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(3)  'グラフスケール型締圧 (Min)
  max = InitDat(4)  'グラフスケール型締圧 (Max)
  def = (max - min) / dev
  p = 8
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(5)  'グラフスケール型温度 (Min)
  max = InitDat(6)  'グラフスケール型温度 (Max)
  def = (max - min) / dev
  p = 16
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(7)  'グラフスケール経過時間 (Min)
  max = InitDat(8)  'グラフスケール経過時間 (Max)
  def = (max - min) / dev
  p = 24
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
'
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  LS21S_MAIN
End Sub
Public Sub LS21S_MAIN()
Dim i%, j%, js%, l%, ist0%, ist1%, iflg%, isflg%
Dim ied%, ips%, i_s%, I_s0%, irei%, r_ch%, ix%, ix0%, iy%, isp%, i_s_do%
Dim stime%, ii%, iii%, istend%
Dim ie02%, ie03%, ie04%, ituflg%, iSRcount%, iki%, ikii%
Dim ie%, ie0%, ie1%, ie2%, ie3%, ie4%, ie5%, iFlg_hijyou%, iflghsmsg%
Dim m_l%, sv%, zch%
Dim ivd%, id_0%, id_1%, id_2%
Dim ct_dummy!, iz3%, itc%, ict%, ikat%
Dim idmy%, ch%, hdt%, flindex%, imax%, sts1%, sts2%, ch1%, ch2%
Dim sts As Long                                     '2006.4.14
Dim it_ts%, i_ts%
Dim dmy$, sdt$, c$, com$, tdate$, ttime$, kjdisp$
Dim isub As Long, jsub As Long, ksub As Long
Dim flg As Long, cnt As Long
Dim iwt!, S_StartTime!
Dim sdata!    '  05.11.26 s.s. overflow 対策
Dim ndata!, mdata!, ntemp!, mtemp!, ntemp0!, mtemp0!, htemp!
Dim imachi!, hs5_fintime!, hs5_sttime!, hs5_difft!, hs5_diffTold!
Dim st!, ev!, sev!, fin!, it!, it0!         '/* 時間用データ */
Dim btemp!(0 To 4), bposi!, bpre! '/* 温度　位置　圧力 の前データ */
Dim stTime!, evtime!, fintime!, sevTime!, mTime!, tsTime!, endTime!
Dim dt!(0 To 4)
Dim diTime!, diTime1!, diTime2!, diTimeSR!, pdt!, pp!, pml!
Dim x1dt!, x2dt!, pos!
Dim r_z_now!, r_z_ave!, r_z_dum!(0 To 180)    ' /* 2002.7.10　追加　突当成形　*/
Dim epsilon!
Dim cp_z!, cc_time!(0 To 3), ct_temp!(0 To 2)   ' CP , CT 用
Dim ct_t!(0 To 10)
Dim avekatJ!(0 To 10), katJ!
Dim zclear!
Dim tudiffTime!     '  '100310 追加
Dim dumlbl14$      ' 成形ショット数の画面表示用　ダミー190428 追加
'
 On Error GoTo errHandler:
' ---  init  val-----------------
  ppos = "SC"   'LS21_SC  現在位置
  ips = 1
    If Saikaiflg = True Then            ' i_s:　成形回数カウンタ
            i_s = 0                     '再開時は、初回からカウント
        Else
            i_s = -1
    End If
'  i_s = -1            'i_s: 成形回数    初期値=-1　loop内で　最初に　カウントアップするため
  iFlg_hijyou = 0
  For ii = 0 To 3: idcflg(ii) = 0: Next ii
  For ii = 0 To 10: ct_t(ii) = 0: Next ii
  c = "0"
  ivd = 0:   id_0 = 0: id_2 = &H8
  For ii = 1 To 180: r_z_dum(ii) = 0#: Next ii
  For i = 0 To 5: For ii = 0 To 10: kaatsuJ(ii, i) = 0#: Next ii: Next i
  For ii = 0 To 10: avekatJ(i) = 0#: Next ii
  Label10.Caption = "  No   SL   Ave.   0   -1   -2   -3   New-T Old-T"
  tsTime = 0#
'
  Label12(0).Visible = False
  Label12(1).Visible = False
  Label12(2).Visible = False
'
'----------------------- 連続成形メインプログラム
  C870Stop
  ServoON       '/* サーボｏｎ */
  CtlDisp       '位置制御
  TrnsReqOFF    '搬送依頼信号OFF
  SeikeiON         '成形ON　連続又は１回成形中
'/***********     ﾒﾚｯｸ　C-853ボード初期設定　　　*************/
'/* SPEC INITIALIZE CMD OUT */
'/* カウンタボードの初期設定 */
  InitDat(10) = 0
'/* 加減速ﾚｰﾄｾｯﾄｺﾏﾝﾄﾞ */
  C870AccRate
'/* 速度設定 */
  C870LSPDSet 300    '/* 300 pps 0.066mm/sec */
'/* ディレータイム設定 */
  C870DelayTime
  rstcm1   '  compareter reset
'/***********     ﾒﾚｯｸ　C-853ボード初期設定　終了  *************/
'/* ＡＴＣ温度リセット */
'/* ロボットデータのフロッピーからの読みとり */
  rozFileLoad
'
'/* 成形データ保存ファイルの作成　*/
  RecDtSave0 InitDat(11)
'
'
  it_ts = Int(roz(1))   ' 10       '/* 突き当て達成　ﾁｪｯｸ　平均する回数 */
  epsilon = roz(0)    ' 0.0005   '/* 突当　許容幅　　mm　　*/
    i_s_do = -1   ' Do Loop の　回数   '　成形　Do　Loop(本体のDo Loop）の回数　　　　edit でキャンセルされないように　ここへ移動。 2007.11.26
    kataNoPnt = 0  ' 型No の　ポインター　初期設定
'
'-------------------------------------------------------------------------------------
st:
  If ied = 2 Then GoTo st2:             '  この文　気になる！！　ied=2　は　無い！！　　editの時は、ied=1　　それ以外は、ied=0
'  ---　2007.11.27　追加　kataNo表示  更新
    For iii = 0 To katamax
        kataNoHyj(iii) = kataNo(iii)
        kataNoHyj(iii + katamax + 1) = kataNo(iii)
        kataNoHyj(iii + (katamax + 1) * 2) = kataNo(iii)
        kataNoHyj(iii + (katamax + 1) * 3) = kataNo(iii)
    Next iii
' --- label13(8) へ　katamax（ステーション数）を表示
    Label13(8) = katamax
'
'/*  制御ファイルのオープン */
  coxDtRead gcoxFldir & gcoxFlName
  Label2(0).Caption = Format(ptime, "0")
  '/* グラフィック画面の初期化 */
  InitDat(8) = ptime  'グラフスケール経過時間 (Max)
  SetData
  lGphNo0 = 0
  lGphNo = 0
  MoniGraph Me.Picture1, lGphNo0, lGphNo
  For itc = 0 To 9
    Label4(itc).Caption = Format(T_keisu(itc), "0.000")
    Label6(itc).Caption = Format(Z3_Hosei(itc), "0.000")
    If itc < T_keisuCont(0) Then
         Label4(itc).BackColor = T_keisuCol!(1)
         Label6(itc).BackColor = T_keisuCol!(1)
         Label11(itc).Caption = kataNo(itc)
       Else
         Label4(itc).BackColor = T_keisuCol!(0)
         Label6(itc).BackColor = T_keisuCol!(0)
         Label11(itc).Caption = " "
    End If
    If (iflgKataTorF(itc) = False) Then
         Label4(itc).BackColor = T_keisuCol!(4)
         Label6(itc).BackColor = T_keisuCol!(4)
    End If
  Next itc
  If (katCflag = True) Then
         Label7(0).BorderStyle = 1  '  枠有り
         Label7(1).BorderStyle = 1  '  枠有り
    Else
         Label7(0).BorderStyle = 0  '  枠なし
         Label7(1).BorderStyle = 0  '  枠なし
  End If
''/* 予備加熱温度設定 */
'/* 軸駆動制御コマンドのファイルからの読み取り */
  i = 0
  Do
    sdt = Right("     " & Format(i, "0"), 4)
    sdt = sdt & "  " & Right("     " & Format(seg_num(i), "0"), 4)
    sdt = sdt & "  " & Right("     " & Format(ic(i), "0"), 4)
    sdt = sdt & "  " & Right("         " & Format(z(i), "0.000"), 7)
    sdt = sdt & "  " & Right("         " & Format(vel(i), "0.0"), 7)
    sdt = sdt & "  " & Right("       " & Format(pres(i), "0"), 6)
    sdt = sdt & "  " & Right("     " & Format(t0(i), "0.0"), 4)
    sdt = sdt & "  " & Right("     " & Format(p(i), "0.0"), 4)
    Label2(12).Caption = sdt
    If pres(i) >= 1000 Then ips = 2    '/* ﾌﾟﾚｽ圧が1ton以上で軸変更 */
    i = i + 1                          '/*軸自動描画時のスケール変更用*/
    If ic(i - 1) = 9 Then Exit Do
  Loop

  istend = i   '  /* コマンド数の最大値 */
  ic(i) = 10
  ic(i + 1) = 10 '  /* 軸制御方式　終了の意味　だめ押し*/
'
''
'/* 表題の表示 */
  Label2(2).Caption = gcoxFlName
'/* 原点出し */
  Label2(6).Caption = "原点出し実行"
  genten
  Ready_Wait
  Counter0
  Label2(6).Caption = "原点出し完了"
'/* カウンタにゼロを書き込む */
  C870CntPreSet 0   'ＣＯＵＮＴＥＲ ＰＲＥＳＥＴ ＣＯＭＭＡＮＤ
  pos = r_z()
  GCnt0 = 0
  GCnt1 = 0
'
'
'/* 自動運転認識 */
  ch1 = 1            'システムレディー
  ch2 = 3            '自動
  Do
    DoEvents
    If FrmMenuFlg = False Then GoTo eend:            'メニューから抜けるときfalse
    '
    DioInput ch1, sts1
    DioInput ch2, sts2
    If sts1 = 1 And sts2 = 1 Then Exit Do
  Loop
'/* 成形プロセス開始　連続前コマンド */

  flindex = 0      '制御コマンドファイルの位置
  Do
    DoEvents
    '-------------- ピラニ計読み
'    LS21S_Monitor    '2006.12.21 削除 s.f
    'flindex = flindex + 1
    com = Left(scom(flindex), 1)
    isub = sisub(flindex)
    sdt = Right("    " & scom(flindex), 2)
    sdt = sdt & Right(Space(15) & Format(isub, "0"), 15)
    If (com = "S") Or (com = "L") Then
      jsub = sjsub(flindex)
      ksub = sksub(flindex)
      sdt = sdt & Right(Space(15) & Format(jsub, "0"), 15)
      sdt = sdt & Right(Space(15) & Format(ksub, "0"), 15)
    End If
    Label2(7).Caption = sdt
    flindex = flindex + 1
    i = 10
    '
    If ied <> 0 Then GoTo jp0:
    '
    Select Case com
      Case "B"
      Case "N"    '/* 窒素ガスの注入 */
        If Mid(scom(flindex), 2, 1) = "S" Then
          If isub = 1 Then
            N2Open
          End If
          If isub = 0 Then
            N2Close
          End If
        End If
      Case "J"    '/* 時間待ち */
        evtime = Timer

        Do
          fintime = Timer
          DoEvents
          If diffTime(fintime, evtime) >= isub Then Exit Do
        Loop
      Case "K"    '/* 加熱 */
        Select Case Int(isub)
        Case 1
          HeatON
        Case 0
          HeatOFF
        End Select
      Case "S"    '/* ＡＴＣ温度設定 */
        evtime = Timer              '待ち初めの時間
        ntemp0 = isub
        mtemp0 = jsub
        ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1))  'ntemp0
        mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1))  'mtemp0
        Do
          DoEvents
          fintime = Timer         '現在時間
          diTime = diffTime(fintime, evtime)
          If ksub <> 0 Then x1dt = diTime / ksub
          ndata = (ntemp0 - ntemp) * x1dt + ntemp
          mdata = (mtemp0 - mtemp) * x1dt + mtemp
          TempSet 2, ndata
          TempSet 3, mdata
          If diTime >= ksub Then Exit Do
        Loop
        ntemp = ntemp0
        mtemp = mtemp0
        TempSet 2, ntemp
        TempSet 3, mtemp
      Case "R"    '/* 冷却 */
        Select Case Int(isub)
        Case 0    '冷却大　ＯＦＦ
          CoolOFF
        Case 1    '冷却大　ＯＮ
          CoolON
        Case 2    '冷却小　ＯＮ
          CoolON
        End Select
    End Select
jp0:
    If i < 24 Then
      i = i + 1
    Else
    End If
    If com = "B" Then Exit Do
  Loop
'/* 成形プロセス連続運転開始 */
'/* データを読み取る */
'/* ブザーを鳴らす */
  'Label2(4).Caption = ""
'-----------------------------------------------------------------------------
st2:
'/* タイトルの表示 */
'/* 型締圧軸の表示 */
'/* 座標値軸の表示 */
'/* 搬送用Ｚ軸位置変更枠表示 */
'  Label2(5).Caption = Format(roz(0), "0.0000")     '/* 突当成形para　幅 */
  Label2(6).Caption = Format(roz(0), "0.0000") & Format(roz(1), "0.0")     '/* 突当成形para　時間 */
'------------------------------------------------------------------------------
'/* 成形開始 */
'    i_s_do = -1   ' Do Loop の　回数           '  st: の　前へ移動 2007.11.26
  Do        '----------------- DO LOOP
    DoEvents
    I_s0 = i_s                                '　I_s0：　1回前の成形回数　保持
    i_s = i_s + 1                              '　i_s：　成形回数　　loopんｐ頭でカウントアップ
    i_s_do = i_s_do + 1
    js = 0
    ist0 = -1
    ist1 = -1
    ie0 = 0
    ie1 = 0
    ie2 = 0
    ie3 = 0
    S_StartTime = Timer
    stTime = Timer
    sevTime = Timer
    diTimeSR = -9999.99                        ' 温度設定　ＳＲの初期化
    iSRcount = 1                               ' 温度設定　ＳＲの初期化
    For ii = 0 To 10
      ct_t(ii) = 0
    Next ii    ' 温度設定　ＳＲの初期化
'
    Label4(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(3)    '  文字　ピンク(ポインター）
    Label6(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(3)    '  文字　ピンク(ポインター）
    Label11(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(3)    '  文字　ピンク(ポインター）
    Label4(T_keisuCont(1) - 1).BorderStyle = 1  '  枠付きにする
    Label6(T_keisuCont(1) - 1).BorderStyle = 1  '  枠付きにする
    Label11(T_keisuCont(1) - 1).BorderStyle = 1  '  枠付きにする
    Command1.Visible = False
'
    iz3 = Z3_HoseiCont(2)   ' Z補正　を実施する　ZNo.　　　’07.9.27　追加
    z(iz3) = z(iz3) + Z3_Hosei(T_keisuCont(1) - 1) '  ”Z3"の補正値set
'/*  制御ファイル名と　保温停止回数　表示
  Label2(2).Caption = gcoxFlName + " -" + Format(iHoonStopNo, "0000")
'/* カウンタへの出力ｕｐ */

    If i_s <> 0 Then                '　1回目は"0"のためカウントアップしない
      InitDat(11) = InitDat(11) + 1  '成形カウンタトウタル
'      InitDtSave                   ' E  成形後にsave　02.8.25 s.f.
      Label2(13).Caption = Str(InitDat(11))
    End If
'/* 成形枠の表示 */　　　-------　画面表示の　最初
ejs1:
  lGphNo0 = 0
  lGphNo = 0
  MoniGraph Me.Picture1, lGphNo0, lGphNo
'/* Ｘ軸の表示 */
'/* Ｙ軸の表示 */
'/* ｼｮｯﾄ数ｻｲｸﾙﾀｲﾑ枠表示 */
    sdt = Format(Int(stime / 60), "0") & "分" & Format(Int(stime) Mod 60, "0") & "秒"
    Label2(8).Caption = sdt
    Label2(9).Caption = Format(i_s, "0")
    InitDat(10) = i_s               '成形カウンタ
'
''    加圧時間制御　下限、上限の表示       for no uchigawa he idou
     Label7(0).Caption = Format(DkatJ(0), "0.0")
     Label7(1).Caption = Format(DkatJ(1), "0.0")
    If (katCflag = True) Then
         Label7(0).BackColor = TKatBackCol(1)
         Label7(1).BackColor = TKatBackCol(1)
     Else
         Label7(0).BackColor = TKatBackCol(0)
         Label7(1).BackColor = TKatBackCol(0)
    End If
''
'/* カウンタへの出力ダウン */
'/* データの取り込み */'
    evtime = Timer
    iflg = 1
    ied = 0
    ttime = Time
    mTime = Timer
'-----------------------------------------------------------------------------------
'----------------------------- For Loop i　　先頭
    imax = ptime * 60
    For i = 1 To imax
start:
      DoEvents               '2005.12.17 OverFlow 対策 s.f.  2006.3.3 復活 s.f.
      ituflg = 0            '　タイムアップflgのリセット10/5
'
      If seikeiKaisu <> 0 Then
         Label1(4).Visible = True
         Label1(6).Caption = s_kaisu
      Else
         Label1(4).Visible = False
         Label1(6).Visible = False
      End If
'
'/* 成形軸のドライブ*/　　　’　ist0　＝　現在の軸コマンドNo.　　　それぞれの軸コマンド終了時にカウントUP
        If ist0 > 0 Then
          If ic(ist0 - 1) = 10 Then      '  /* ic(ist0-1)=10 終了の意味　*/
            ist0 = ist0 - 1
          End If
        End If
          sdt3 = DispSegm(ist0)
          Label2(12).Caption = sdt3
        If ist0 <> ist1 Then             '　新ｾｸﾞﾒﾝﾄ開始条件
            gOrgFlg = False                '原点復帰完了=TRUE
            ist1 = ist0
            sevTime = Timer              '軸制御セグメント開始時間
'
            If (ist0 > 0 And ist0 < 11) Then   '　開始時間の表示　ｄｅｂｕｇ用
               diTime1 = diffTime(sevTime, stTime)          '2002.10.09 KYOCERA
               sdt = Format(ist0, "0") & "=" & Format(Int(diTime1 / 60), "0") & ":" & Format(Int(diTime1) Mod 60, "00")       '2002.10.09 KYOCERA
            End If
'
            Select Case ic(ist0)  '-------- 軸制御モード番号
            Case 0, 8   '-------------------- 位置制御
              List1.Enabled = True
              List2.Enabled = True
              ppos = "SC JikuStart 0 8"
              Ready_Wait    '
              CtlDisp     'outp(DIO_P+3,9); サーボON & 速度上限S12
              s_drive z(ist0), vel(ist0)
            Case 1, 3, 7   '-------------------- 速度制御  '2004.3.8 sf
              ppos = "SC JikuStart 1 3 7"
              List1.Enabled = False
              List2.Enabled = False
              m_l = vel(ist0)
              If m_l > 50 Then m_l = 50
              setcm1 z(ist0)
              Ready_Wait    '
              CtlVelo       'outp(DIO_P+3,5);  速度制御へ切り替え
              Do       ' 「カウンター一致」状態脱出用
                DoEvents
                sts = C870Sts(3)   'sts=1の時　成立＝＞「-1」　sts=0の時不成立＝＞「0」
                If (sts And &H1) = 0 Then Exit Do   '「PULSE と COMPARE が一致状態」時loop
              Loop
            Case 2    '-------------------- ダミー
              ppos = "SC JikuStart 2"
              List1.Enabled = True
              List2.Enabled = True
              Ready_Wait    '
              CtlDisp     'DioOut 12,1  位置制御 '  02.10.1 追加
              Ready_Wait    '
              ServoON     'outp(DIO_P+3,1);
            Case 9    '-------------------- 終了
              ppos = "SC JikuStart 9"
              List1.Enabled = True
              List2.Enabled = True
              Ready_Wait    '
              CtlDisp     'outp(DIO_P+3,9);
              genten
              'Ready_Wait
              For ii = 1 To 180          '/* 制御３用の初期化 */
                r_z_dum(ii) = 0#
              Next ii
              i_ts = 1
              r_z_ave = 0#
            End Select
        End If
'
        fintime = Timer         '2002.10.09 KYOCERA   fintime:現在時間
'
'/* タイムアップ処理 */
      '2002.10.09 KYOCERA
        If ist0 < 0 Then GoTo sj1:
' ----　　timerの異常値検出skip部
          fintime = Timer       '　現在時刻の取得　　 ' 2010.3.10 新設　LongTime判断追加
          tudiffTime = diffTime(fintime, sevTime)     ' 2010.3.10 新設
          If ((ic(ist0) < 10) And (tudiffTime > (t0(ist0)) * 1.2)) Then ' 2002.10.17 KYOCERA '10/3/10 新設 '10/4/12 ic(ist0>10 追加
             sdt = "ﾀｲﾑｱｯﾌﾟskip  " & Format(tudiffTime, "0.0")   ' 2010.3.10 新設
             Label2(6).Caption = sdt                  ' 2010.3.10 新設
             GoTo TimeUpEnd:   '' 2010.3.10 新設 　　設定時間の1.5倍より大きかったら異常値と判断してtimeupルーチンをスキップ
          End If                                      ' 2010.3.10 新設
'　----　通常のタイムアップ判断
        If ((ic(ist0) < 10) And (tudiffTime > t0(ist0))) Then '2002.10.16 KYOCERA 2002.10.17 KYOCERA     '10/4
             ituflg = 1
             sdt = "ﾀｲﾑｱｯﾌﾟ" & Format(tudiffTime, "0.0")
             sdt = sdt & " " & Format(t0(ist0), "0.0") & " " & Format(ist0 + 1, "0")
             Label2(6).Caption = sdt
'
                RecEmgDtSave sdt3, sdt1, sdt2
                gemgmsg = "ﾀｲﾑｱｯﾌﾟ"
                hijyou        '非常停止処理
                iFlg_hijyou = 1     '   タイムアップ
                GoTo eend:
'
        End If
TimeUpEnd:
'
'/* 終了信号の処理 */
        Select Case ic(ist0)
        Case 0, 8   '/* 位置制御の場合 */
          ppos = "SC JkE 0 8"
          If (C870Sts(1) And 1) = 0 Then
             ist0 = ist0 + 1
          End If
        Case 1    '/* 速度制御の場合 */
            ppos = "SC JkE1"
          pdt = pres(ist0)
          pp = p(ist0)
          pml = m_l
            ppos = "SC JkE1 -1cal"
          cal_pid pdt, pp, pml
            ppos = "SC JkE1 cal_pid"
          sts = C870Sts(3)  'status3 を読む
             ppos = "SC JkE1 sts=C870"
         If (sts And &H1) <> 0 Then      ' 成立で「-1」　　不成立で「0」
            ist0 = ist0 + 1             '/* 位置達成で終了 */
            Label2(6).Caption = "位置 pass CNT " & Str(ist0)   '11/2 sf
            rstcm1   '  compareter reset
         Else                       ' 2008.2.21  変更　１秒に１回行き過ぎを確認へ
'
           If Int(mTime) = Int(Timer) Then
             If r_z() >= z(ist0) Then
               ist0 = ist0 + 1             '
               Label2(6).Caption = "位置 pass PC " & Str(ist0)
             End If
           End If
         End If
         ppos = "SC JkE1 r_z -1"
        Case 3    '/* 速度制御　突当成形の場合  2002.7.10 ls21_tcよりコピー */
           ppos = "SC JkE3"
          pdt = pres(ist0)
          pml = m_l
          pp = p(ist0)
           ppos = "SC JkE3 -1cal"
          cal_pid pdt, pp, pml
           ppos = "SC JkE3 cal_pid"
          sts = C870Sts(3)  'status3 を読む
           ppos = "SC JkE3 sts=C870"
          If (sts And &H1) <> 0 Then
            ist0 = ist0 + 1             '/* 位置達成で終了 */
            Label2(6).Caption = "位置 pass CNT " & Str(ist0)   '11/2 sf
            rstcm1   '  compareter reset
         Else                       ' 2008.2.21  変更　１秒に１回行き過ぎを確認へ
           If Int(mTime) = Int(Timer) Then
             If r_z() >= z(ist0) Then
               ist0 = ist0 + 1             '
               Label2(6).Caption = "位置 pass PC " & Str(ist0)
             End If
           End If
         End If
'
          If r_z() < z(ist0) Then
'            r_z_now = r_z()                    '2008.2.23 移動
              ppos = "SC JkE3 r_z -2"
            If Int(tsTime) <> Int(mTime) Then
              tsTime = mTime                  '/* １秒前と、２秒前と */
              r_z_now = r_z()                    '2008.2.23 移動
              If Abs(r_z_now - r_z_ave) < epsilon Then
                ist0 = ist0 + 1               '/* it_ts回連続　epsilon以下 */
              Else                            '/* で　突当達成で終了 */
                r_z_dum(i_ts) = r_z_now
                r_z_ave = 0#
                For ii = 1 To it_ts
                   r_z_ave = r_z_ave + r_z_dum(ii)
                Next ii
                r_z_ave = r_z_ave / it_ts
                i_ts = i_ts + 1
                If i_ts > it_ts Then i_ts = 1
              End If
            End If
          End If
        Case 7    '/* 速度制御　上軸衝突判定付　　　　　　　　　2004.3.8 s.f. 軸制御「７」追加　　ここから　*/
'　　　  　　　　/*　指定圧力より高い圧力が３秒以上続いたら非常停止　　*/
          ppos = "SC JkE7"
          pdt = pres(ist0)
          pp = p(ist0)
          pml = m_l
          cal_pid pdt, pp, pml
          sts = C870Sts(3)  'status3 を読む
          If (sts And &H1) <> 0 Then
            ist0 = ist0 + 1             '/* 位置達成で終了 */
            Label2(6).Caption = "位置 pass CNT " & Str(ist0)   '11/2 sf
            rstcm1   '  compareter reset
          Else                       ' 2008.2.21  変更　１秒に１回行き過ぎを確認へ
            If Int(mTime) = Int(Timer) Then        '　１秒に1回チェック
              If r_z() >= z(ist0) Then
                ist0 = ist0 + 1             '
                Label2(6).Caption = "位置 pass PC " & Str(ist0)
              End If
            End If
          End If
'
          If Int(tsTime) <> Int(mTime) Then '2008.2.23 変更 1秒に1回チェック
             tsTime = mTime                  '/* １秒前と比較 */
             bpre = r_pres()
             If bpre > pdt Then                ' 2008.2.18 変更
               i_ts = i_ts + 1               '/* i_ts回連続して　圧力が指定値以上 */
                If i_ts > 3 Then
                  gemgmsg = "軸制御　７"
                  hijyou        '非常停止処理
                  'getch
                  iFlg_hijyou = 2    '    軸制御 7　error
                  GoTo eend:
                End If
             End If
          End If                                 '/*     '2004.3.8　ここまで　*/
       Case 9    '終了
          ppos = "SC JkE9"
          sts = C870Sts(1)
          If (sts And 1) = 0 Then
            ist0 = ist0 + 1     '/* 完了 */
            If Abs(r_z()) > 0.1 Then
              Label2(6).Caption = "原点不良"
              ist0 = ist0 - 1
              genten              '原点出し
            End If
          Else
            '/* カウンタにゼロを書き込む */
            Ready_Wait
            Counter0
          End If
        End Select
''                                                  ' 2007.12.21 delete  速度制御値の表示
jscmdend:                               '軸制御コマンド　ｅｎｄ  10/4 sf
'
'/* エラー表示 */
      If ArmChk <> 0 Then               'アラームメッセージ
        frmerr_sign.Show   'ALM出力
      Else
        Unload frmerr_sign
      End If
'/* プロセス実行 */
sj1:
      If iflg = 1 Then                          '　iflg=1　前のｺﾏﾝﾄﾞ終了のフラグ
        com = scom(js + flindex)                '　js　は　コマンドのNo.
        isub = sisub(js + flindex)
        jsub = sjsub(js + flindex)
        ksub = sksub(js + flindex)
        js = js + 1                              '　jsを　次ぎ用に　１進めておく
'
        evtime = Timer                  '  '05.12.17 evtime カウント開始をここへ設置　s.f.
'
        sdt = com & Right(Space(7) & Format(isub, "0"), 7)    ' ｺﾏﾝﾄﾞの表示
'
        If ((Left(com, 1) = "S") Or (Left(com, 1) = "L")) Then
          sdt = sdt & Right(Space(7) & Format(jsub, "0"), 7)
          sdt = sdt & Right(Space(7) & Format(ksub, "0"), 7)
        Else
          sdt = sdt
        End If
        Label2(7).Caption = sdt
      End If
        'システムレディ/* 非常停止の場合は成形中止 */
          sts1 = SystemReadyChk()   'システムレディ or 非常停止
          sts2 = AutoChk()          '自動状態？
          If sts1 = 0 Or sts2 = 0 Then
            gemgmsg = gemgmsg + ArmEmgMsgChk$()         ' '08.7.15  gemgmsg + 追加
            iFlg_hijyou = 10              '非常停止ﾒｯｾｰｼﾞのｓａｖｅ
            FrmMenuFlg = False              'メニューから抜けるときfalse
            NextView = 1
            Exit Do                         '　Loopから飛び出す＊＊＊
          End If
        '
          Select Case Left(com, 1)
          Case "D"    '------------ 成形室の型の有無   0:成（無）予（無）、1:成（有）予（無）、　2:成（無）予（有）、　3:成（有）予（有）
             ppos = "SC Proc D"
             If (isub = 0) Then     '在否センサーチェック
               If (KataChk() > 0) Then                '  2004.10.30  型在否チェック用センサの動作確認用
                 sdt = "DC　在否センサー異常（型有り！！）"
                 Label2(6).Caption = sdt
'
                  sdt2 = sdt2 & sdt
                  RecEmgDtSave sdt3, sdt1, sdt2
                  gemgmsg = "DC 型有り"
                  hijyou        '非常停止処理
                  iFlg_hijyou = 3          '　DC　error　型有り
                  GoTo eend:
               Else
                  GoTo scend:
               End If
            End If                                 '  2004.10.30  型在否チェック用センサの動作確認用
'
'            If (KataChk() < 3) Or (Karauchiflg = True) Then '成形室,または、予備加熱 に型が無い　　'08.6.8 '08.7.14 削除　ＬＳ12改造中専用
            If (KataChk() < 2) Or (Karauchiflg = True) Then '成形室に型が無い　　'06.12.21  '08.7.14 復活
'            If KataChk() < 3 Then '型が無い
 '             Label2(4).Caption = "CASE D 成形室型無し DO2"
               fintime = Timer       ' 現在時間　　　　'2006.3.3　　追加　s.f.
              If (diffTime(fintime, evtime) < isub) Then
                 iflg = 0             ' 時間未達の場合
              Else
                 idmy = js            '　時間待ち終了の場合　　js　=　次のコマンドのNo.　　(最初に読み取るため、値は1個進んでいる）
                 Do
                   DoEvents
                   dmy = scom(idmy + flindex)          '　次のコマンドを読み取る
                   If "LA" = dmy Then  '----- コマンドLAまで進める
                     js = idmy          '　　LAが見つかったら　次のコマンドNo.を　LAの　No.にセット
                     '------------- LAが見つかったら次に、セグメントをモード８まで（9の２つ前まで）進める
                     Do
                       DoEvents
                       If ic(ist0) = 8 Then
                         ist0 = ist0 - 1
                         sevTime = Timer        '  2005.12.17 Timeup防止 念のため s.f.
                         Exit Do
                       End If
                       ist0 = ist0 + 1
                       If ist0 > 50 Then   'エラー
'
                         sdt = "DCｺﾏﾝﾄﾞ ist0 > 50 ｴﾗｰ"
                         Label2(6).Caption = sdt
                         RecEmgDtSave sdt3, sdt1, sdt2
                         gemgmsg = "DC　エラー　4"
                         hijyou        '非常停止処理
                         iFlg_hijyou = 4        '　DC　コマンドエラー
                         GoTo eend:
'
                       End If
                     Loop
                   '
                     Exit Do
                   End If
                   idmy = idmy + 1
                   If idmy > 50 Or "EN" = dmy Then 'エラー
'
                         sdt = "DCｺﾏﾝﾄﾞ ist0 > 50 ｴﾗｰ"
                         Label2(6).Caption = sdt
                         RecEmgDtSave sdt3, sdt1, sdt2
                         gemgmsg = "DC　エラー　5"
                         hijyou        '非常停止処理
                         iFlg_hijyou = 5          '　　DCコマンドエラー
                         GoTo eend:
'
                   End If
                 Loop
'
                 iflg = 1                    '　ｺﾏﾝﾄﾞ終了処理
                 idcflg(1) = 1               '  DCフラグ　型無=1　型有=0
                  sevTime = Timer             ' 2005.12.17 念のため
              End If
            Else
              idcflg(1) = 0             '  型がある場合　idcflg=0にして抜ける
            End If                    '　型がある場合はそのまま抜ける
'
          Case "L"    '------------ 成形室に型が無かった時の飛び先番地
             ppos = "SC Proc L"
             If (KataChk() < 3) Then GoTo caselend: '型が無い
             If (iflghoonStop = False) And (iflg5Stop = False) Then GoTo caselend:
'                      ------------  型があり、かつ　保温停止フラグ　ONの時の処理
             iflg = 0
             Command2(0).Enabled = False
             Command2(9).Enabled = False
'　　　　　　　　　　　------------　温度データの取り込み
              If (iflghoonStop = True) Then
                 htemp = isub
              End If
              If (iflg5Stop = True) Then
                 htemp = jsub
              End If
'
              ntemp0 = htemp
              mtemp0 = htemp
              ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1))  'ntemp0
              mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1))  'mtemp0
              TempSet 2, ntemp0
              TempSet 3, mtemp0
         ''  　保温停止　-----------------------------
              If (iflghoonStop = True) Then
                 Label12(0).Visible = True
                 Label12(1).Visible = True
                 Label12(2).Visible = True
                 Command1.Visible = True
                 Label12(0).Caption = "保温停止中"
                 Label12(1).Caption = " 経過時間"
'
         ''  　保温停止　時間待ち　-----------------------------
                 hs5_sttime = Timer
                 imachi = 60 * 60 - 1          '  待ち時間　60分決定
                 Do
                   DoEvents
                   hs5_fintime = Timer
                   hs5_difft = diffTime(hs5_fintime, hs5_sttime)
                   If (hs5_difft < imachi) And (iHoteishuryo = 0) Then
                      If (Int(hs5_difft) <> Int(hs5_diffTold)) Then
                          Label12(2).Caption = Format(Int(hs5_difft / 60), " 00分") + Format(Int(hs5_difft) Mod 60, " 00秒")
                          hs5_diffTold = hs5_difft
                          End If
                       Else
                          Exit Do              '　時間待ち終了
                       End If
                 Loop
'
                 Label12(0).Visible = False
                 Label12(1).Visible = False
                 Label12(2).Visible = False
                 Command1.Visible = False
                 iHoteishuryo = 1
  '
                  iflg = 1
                  GoTo caselend2:
              End If                                '　2012.1.8　　5分止め効かないバグ修正。　この「endif」が、2行上にあった。
'
         ''  　5分間停止　-----------------------------
              If (iflg5Stop = True) Then
                 Label12(0).Visible = True
                 Label12(1).Visible = True
                 Label12(2).Visible = True
                 Label12(0).Caption = "5分停止中"
                 Label12(1).Caption = " 再開まで "
'
         ''  　5分間保温停止　時間待ち　-----------------------------
                 hs5_sttime = Timer
                 imachi = 5 * 60 - 1          '  待ち時間　５分決定
                 Do
                   DoEvents
                   hs5_fintime = Timer
                   hs5_difft = diffTime(hs5_fintime, hs5_sttime)
                   If (hs5_difft < imachi) Then
                      If (Int(hs5_difft) <> Int(hs5_diffTold)) Then
                          Label12(2).Caption = Format(Int((imachi - hs5_difft) / 60), "  0分") + Format(Int((imachi - hs5_difft)) Mod 60, " 0秒")
                          hs5_diffTold = hs5_difft
                          End If
                       Else
                          Exit Do              '　時間待ち終了
                       End If
                 Loop
'
                 Label12(0).Visible = False
                 Label12(1).Visible = False
                 Label12(2).Visible = False
              End If
  '
'　　　　　　　　　　　-------------　終了の処理
caselend2:    TempSet 2, ntemp    ' 元の温度に戻して終了
              TempSet 3, mtemp
              If (iflghoonStop = True) Then
                iHoonStopNo = iHoonStopNo + 100  ' 保温停止回数のカウントアップ
                iflghoonStop = False   ' フラグをリセット
                Command2(9).BackColor = CmndColoff(9)    'コマンドボタンの色を戻す
              End If
              If (iflg5Stop = True) Then
                iHoonStopNo = iHoonStopNo + 1  ' 保温停止回数のカウントアップ
                iflg5Stop = False   ' フラグをリセット
                Command2(0).BackColor = CmndColoff(0)    'コマンドボタンの色を戻す
              End If
              
             Command2(0).Enabled = True
             Command2(9).Enabled = True

'
              sevTime = Timer     '　軸制御コマンドがタイムアップしないように　sevtimeのリセット
              evtime = Timer      '  2005.12.17  念のため  s.f.
'
caselend:     iflg = 1            'これを抜けると終了
'
          Case "H"    ' 強制ソーク　　　”ＨＣ”
             ppos = "SC Proc H"
             fintime = Timer      ' 現在時間　　　'　2006.3.3　追加　s.f.
             If (lSokuFlg = True And diffTime(fintime, evtime) < isub) Then
               iflg = 0
             Else
               iflg = 1
               lSokuFlg = False
               Command2(8).BackColor = SokuCor(0)
             End If
'
          Case "S"    '/* ＡＴＣ温度設定 */
             ppos = "SC Proc S"
            If Mid(com, 2, 1) = "R" Then             ' SRの場合  注：関連初期化　Do　Loop　Topにあり
               fintime = Timer
               diTime = diffTime(fintime, stTime)    ' 0.1秒に１回温度取り込み（５回実施）
               If ((diTime - diTimeSR) > 0.1) Then
                   ct_dummy = TempRdMoldTop()    '温度読込
                   ct_dummy = T_keisu_cread(ct_dummy, T_keisu(T_keisuCont(1) - 1))
                   ct_t(0) = ct_t(0) + ct_dummy '温度読込
                   iSRcount = iSRcount + 1
                   diTimeSR = diTime
                   iflg = 0
                   If iSRcount > 5 Then
                      ct_t(0) = ct_t(0) / 5
                      ntemp0 = isub
                      ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1)) 'ntemp0
                      mtemp0 = jsub
                      mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1)) 'mtemp0
                      ntemp0 = ct_t(0) + ntemp0
                      mtemp0 = ct_t(0) + mtemp0
                      ntemp = ntemp0
                      mtemp = mtemp0
                      TempSet 2, ntemp
                      TempSet 3, mtemp
                      ct_t(0) = 0
                      Label2(6).Caption = "SR= " & Format(Int(ntemp), "000") & Format(Int(mtemp), "  000")
                      iSRcount = 1
                      iflg = 1
                   End If
               End If
            Else
             ppos = "SC Proc SA"
              fintime = Timer
              diTime = diffTime(fintime, evtime)        'SAの場合
             ppos = "SC Proc SA af dev"
              If ksub <> 0 Then x1dt = diTime / ksub
              ntemp0 = isub
              mtemp0 = jsub
              ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1))  'ntemp0
              mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1))  'mtemp0
              ndata = (ntemp0 - ntemp) * x1dt + ntemp
              mdata = (mtemp0 - mtemp) * x1dt + mtemp
              TempSet 2, ndata
              TempSet 3, mdata
              If diTime >= ksub Then
                iflg = 1
                ntemp = ntemp0
                mtemp = mtemp0
                TempSet 2, ntemp
                TempSet 3, mtemp
              Else
                iflg = 0
              End If
            End If
          Case "P"    '/* 移動軸制御の駆動 */
             ppos = "SC Proc P"
            If Mid(com, 2, 1) = "W" Then
              Beep
              ist0 = ist0 + 1
              sevTime = Timer          '2005.12.17　念のため　s.f.
            End If
            If Mid(com, 2, 1) = "R" Then
              iflg = 0
              If ist0 <> ist1 Then iflg = 1
              If isub = 4 And ist0 = 0 Then iflg = 1
              If iflg = 1 Then sevTime = Timer             '2005.12.17　s.f.
             End If
          Case "K"    '/* 加熱 */
             ppos = "SC Proc K"
            Select Case isub
            Case 1
              HeatON
            Case 0
              HeatOFF
            End Select
          Case "N"
             ppos = "SC Proc N"
            If Mid(com, 2, 1) = "S" Then
              If isub = 1 Then hdt = hdt
              If isub = 0 Then hdt = hdt
            End If
          Case "R"    '/* 冷却 */
             ppos = "SC Proc R"
            Select Case isub
            Case 2
              CoolON
            Case 1
              CoolON
            Case 0
              CoolOFF
            End Select
          Case "T"    '/* ＡＴＣ１の温度の読み取り */
             ppos = "SC Proc T"
            sdata = TempRdMoldTop()    '上モールド温度
            sdata = T_keisu_cread(sdata, T_keisu(T_keisuCont(1) - 1)) 'ndata
            If (Mid(com, 2, 1) = "L" And sdata > isub) Or (Mid(com, 2, 1) = "G" And sdata < isub) Or (Mid(com, 2, 1) = "E" And (sdata > (isub + 20) Or sdata < (isub - 20))) Then
              iflg = 0
            Else
              If iflg = 2 Then iflg = 1 Else iflg = 2
            End If
          Case "J"    '/* 時間待ち */
             ppos = "SC Proc J"
            DoEvents             ' 2006.5.18  追加　s.f
            fintime = Timer      ' 現在時間　　　　’2006.3.3　追加　s.f.
            diTime1 = diffTime(fintime, stTime)
            diTime2 = diffTime(fintime, evtime)
            If (Mid(com, 2, 1) = "S" And diTime1 >= isub) Or (Mid(com, 2, 1) = "C" And diTime2 >= isub) Then
              iflg = 1
            Else
              iflg = 0
            End If
          Case "C"
             ppos = "SC Proc C"
            Select Case Mid(com, 2, 1)
            Case "P"    '成形終了位置　チェック
              cp_z = r_z()
              Label5(0).Caption = " cp=   " & Format(cp_z, "0.000")
            Case "C"    '　時間チェック
              If isub > 3 Then
                  ict = 5
              Else
                ict = isub + 2
              End If
              fintime = Timer         '現在時間
              cc_time(isub) = diffTime(fintime, stTime)
              sdt = " cc" & Format(isub, "0") & "= " & Format(Int(cc_time(isub) / 60), "0") & ":" & Format(Int(cc_time(isub)) Mod 60, "00")        '2002.10.09 KYOCERA
              Label5(ict).Caption = sdt
              If isub = 3 Then
                diTime1 = diffTime(cc_time(isub), cc_time(isub - 1))
                katJ = diTime1
                sdt = " cc3-2= " & Format(Int(diTime1 + 0.5), "0") & "s"
                Label5(6).Caption = sdt
              End If
'
          Case "T"    '　温度チェック
            If isub > 2 Then
                ict = 2
              Else
                ict = isub
            End If
            ct_temp(isub - 1) = TempRdMoldTop() '温度 0V-300℃ 1V-1300℃           ' v3.30322 誤記訂正
            ct_temp(isub - 1) = T_keisu_cread(ct_temp(isub - 1), T_keisu(T_keisuCont(1) - 1))
            sdt = " ct" & Format(isub, "0") & "=   " & Format(ct_temp(isub - 1), "0.0") & "℃"
            Label5(ict).Caption = sdt
          End Select
          Case "X"    '搬送終了信号（成形開始）
             ppos = "SC Proc X"
            Select Case Mid(com, 2, 1)
              Case "R"    '成形開始 [搬送終了まで待つ]
            '
                TrnsReqON  '搬送依頼信号Ch21出力 (搬送終了解除)
            '
                Do
              '-------------- ピラニ計読み
                  sts = TrnsFinChk()      '搬送終了？
                  If sts = 1 Then
                    TrnsReqOFF            '搬送依頼信号ＯＦＦ
                    Exit Do
                  End If
                  DoEvents           '  注意　このDoEventsを　Do　直後に移すと　誤動作する。　搬送終了2回待ちになる！！
                Loop
'
'            --- 型　No.の表示　一回送り　---
                kataNoPnt = kataNoPnt + 1
                If kataNoPnt > katamax Then kataNoPnt = 0
'
                For iii = katamax To 0 Step -1
                    Label13(iii).Caption = kataNoHyj(katamax - iii + kataNoPnt + katamax + 1 + Val(kataNo(10)))
                Next iii
'
                If (i_s_do) < katamax - 1 Then
                    For iii = kataNoPnt + 1 To katamax
                        Label13(iii).Caption = "空"
                    Next iii
                End If
'
                If katamax <> 7 Then
               '    --- 6st,4st のときは、1 は　素通し　---
                    Label13(2).Caption = Label13(1).Caption
                    Label13(1).Caption = " "
               '    --- 4st のときは、３，４　素通し　---
                    If katamax = 4 Then
                        For iii = 3 To 4
                            Label13(iii + 1).Caption = Label13(iii).Caption
                            Label13(iii).Caption = " "
                        Next iii
                    End If
                End If
'
' ---           型Ｎｏ．　１回送り完了
              Case "W"    '成形終了
              End Select
          Case "E"    '/* 終了　ロボット搬送 */
             ppos = "SC Proc E"
             DoEvents
            If iflg <> 99 Then
              iflg = 0
              If r_z() > 2 Then
                genten
              End If
              TrnsReqON       '搬送依頼信号Ch21出力
              WaitSec 1.5     '
              '搬送表示信号Ch15を待つ
                iflg = 99
              isp = 0
            Else
             'DioInput 13, sts    '搬送終了信号Ch13を待つ
              sts = TrnsFinChk()      '搬送終了？
              If sts = 1 Then
                TrnsReqOFF        '搬送依頼信号OFF
                GoTo send:
              Else
              End If
            End If
scend:
          End Select
cjump:
'
  '-------------- ピラニ計読み
'          LS21S_Monitor　　　　　2005.6.4　削除s.f.
'
'          DoEvents
          lEmgFlg = SystemReadyChk()  '非常停止の確認
          If Int(mTime) = Int(Timer) And lEmgFlg <> 0 Then GoTo start:
           mTime = Timer
'
'                    start: から　ここまで　高速にループ
' ---------------- /* 1秒に1回下に抜ける 画面表示出力*/  ------------------------
'
          ppos = "SC 1sec Disp 1"
'           /* 圧力　ＰＩＤ制御　Ｐ＞１５　なら速度　ゼロ */
          If ist0 >= 0 Then
            If p(ist0) > 15 Then
              DaVoltOut 1, 0        ' 0V D/A ch=1
            End If
          End If
          
'/*　経過時間　*/
          KeikaTime(i) = i
'/*　温度取り込み */
'          DoEvents               '2005.12.17 OverFlow 対策 s.f.
          atemp(i, 0) = TempRdMoldTop()   '上モールド温度 0V-300℃ 1V-1300℃
          atemp(i, 0) = T_keisu_cread(atemp(i, 0), T_keisu(T_keisuCont(1) - 1))
'         atemp(i, 1) = 0                 '下モールド温度
'
'* 成形軸位置の取り込み */
          ppos = "SC 1sec Disp 2"
          aposi(i) = r_z()
'/* 型圧力の取り込み */
          ppos = "SC 1sec Disp 3"
          apre(i) = r_pres()
'
'/* 温度分布の表示 */
'/* 型締圧のプロット */
'/* 座標値のプロット */
          lGphNo = i
          GphDataSet lGphNo0, lGphNo
          MoniGraph Me.Picture1, lGphNo0, lGphNo
          lGphNo0 = lGphNo
jo0:
'/* 各種データの画面下表示 １　*/
          DoEvents           '2006.5.18 OverFlow 対策 s.f. 追加
          sdt1 = Format(atemp(i, 0), "  0.0℃     ")
          sdt1 = sdt1 & Format(apre(i), "0.00kgf    ")
          sdt1 = sdt1 & Format(aposi(i), "0.000mm   ")
          Label2(14).Caption = sdt1
'/* 各種データの画面下表示 ２ */
          it0 = Timer                                                          ' 10/5
          it = diffTime(it0, stTime)
          sdt2 = Format(Int(it / 60), "  0分")
          sdt2 = sdt2 & Format(Int(it) Mod 60, " 0秒")      '2002.10.09 KYOCERA
          sdt2 = sdt2 & "     ct " & Format(diffTime(it0, evtime), "0.0")
          sdt2 = sdt2 & "     st " & Format(diffTime(it0, sevTime), "0.0")
'          sdt2 = sdt2 & "tt   " & Format(diffTime(it0, stTime), "0.0")    '2005.11.23 時間削減のため削除
          Label2(11).Caption = sdt2
'
'/* 時刻表示 */
          Label8.Caption = Time$
'
'/* ﾛﾎﾞｯﾄ位置変更　*/
          'If FrmMenuFlg = False Then GoTo eend:
      Next i   '--------------------------------- For Loop　i　　終端　　１秒に１回「観測時間」分回る
      js = js - 1
      GoTo ejs1:      '/* 表示終了で元画面へ */（次回分　画面表示へ）
'
'
' ----------------  1回分の成形終了　--------------------------------------
send:
'    ---- /* タクトタイムの算出　*/ ----
      ppos = "SC 1回end"
      iSeikeiTorF_flg = True
'　　　　　　　　　　　　　成形後　今回の成形の有効性確認（成形回数用）
        idcflg(3) = idcflg(2)          '  idcflg(3) １回前
        idcflg(2) = idcflg(1)          '  idcflg(2) 今回
'
      If i_s > 0 Then      '　1回目は"i_s=0"のためPass　else以降のみチェック。2回目から計算　'　100304削除　LS改造機は、初回空打ちからスタートのため。　初回ポインターカウントアップバグ対策、　'100412　再度バグ修正　初回もｅｌｓｅ以下チェックへ
'
        If idcflg(2) = 1 Then
           i_s = i_s - 1                  '空の時は　成形回数－１
           InitDat(11) = InitDat(11) - 1  '成形カウンタトウタルの戻し
           iSeikeiTorF_flg = False
        Else
          If idcflg(3) = 1 Then
            i_s = i_s - 1               ' ダミーの時は、無効ショット
            InitDat(11) = InitDat(11) - 1  '成形カウンタトウタルの戻し
           iSeikeiTorF_flg = False
          End If
        End If
      Else
        If idcflg(2) = 1 Then
           i_s = i_s - 1                  '空の時は　成形回数－１
           iSeikeiTorF_flg = False
        End If
      End If          '　100304削除　LS改造機は、初回空打ちからスタートのため。　初回ポインターカウントアップバグ対策
      If i_s = 0 Then iSeikeiTorF_flg = False
'
'     stime = i
      endTime = Timer
      stime = diffTime(endTime, stTime)         ' 2002. 10/5
      InitDtSave            '　データsave　（成形回数）
'
'
' --- 加圧時間の平均値計算　　現在の型No＝T_keisuCont(1)-1　、　現在から　４周前までの平均値
'     --- 今回が　ダミー　の場合、　加圧データ(KatJ)をリセット（0へ）
      If iflgKataTorF(T_keisuCont(1) - 1) = False Then
        For ikat = 0 To 3
          kaatsuJ(T_keisuCont(1) - 1, ikat) = 0#
        Next ikat
      End If
'　　----　’　型変更時の取り扱い 型数不変で新規型に入れ替え（０にリセットする）
     If (i_s > 0) And (i_s <> I_s0) Then    '   -----------------加圧時間制御ルーチン　start
                                            '  --------- 有効な成形かどうかの判定
                
'
        kaatsuJ(T_keisuCont(1) - 1, 0) = katJ    '  katJ=今回の加圧時間
' ---                                            ' 加圧時間平均値　今回の加圧時間　重み（ウェイト）2.0へ　　2007.11.21
        avekatJ(T_keisuCont(1) - 1) = (kaatsuJ(T_keisuCont(1) - 1, 0) * 2 + kaatsuJ(T_keisuCont(1) - 1, 1) + kaatsuJ(T_keisuCont(1) - 1, 2) + kaatsuJ(T_keisuCont(1) - 1, 3)) / (4 + 1)
'
        kjdisp = Format(InitDat(11), "000") & "  "
        kjdisp = kjdisp & Format(T_keisuCont(1), "00") & "  "
        kjdisp = kjdisp & Format(avekatJ(T_keisuCont(1) - 1), "000") & "  "
        For ikat = 0 To 3
           kjdisp = kjdisp & Format(kaatsuJ(T_keisuCont(1) - 1, ikat), "000") & "  "
        Next ikat
'     --- 新T係数計算 ---　　平均値と今回の加圧時間で　評価
'       ---　（１）平均値が　上限下限内にあるか？
        If ((avekatJ(T_keisuCont(1) - 1)) > DkatJ(1)) Then
              T_keisu_dum = T_keisu(T_keisuCont(1) - 1) + 0.001      '上限より大きい場合　+0.001          DkatJ(1)=上限値
        Else
             If (avekatJ(T_keisuCont(1) - 1) >= DkatJ(0)) Then
                  T_keisu_dum = T_keisu(T_keisuCont(1) - 1)       ' 上限以下、かつ、下限以上なら　元の値のまま
             Else
                  T_keisu_dum = T_keisu(T_keisuCont(1) - 1) - 0.001  '下限より小さい場合　-0.001      DkatJ(０)=下限値
             End If
        End If
'
'       ---　（２）今回の加圧時間が　上限下限内にあるか？
        If ((katJ <= DkatJ(1)) And (katJ >= DkatJ(0))) Then
              T_keisu_dum = T_keisu(T_keisuCont(1) - 1)             ''今回の加圧時間が　上限と下限内側なら　T係数は　変えない！
        End If
'       ---　（3）今回の加圧時間が　下限以下か？
        If (katJ < DkatJ(0)) Then
              T_keisu_dum = T_keisu(T_keisuCont(1) - 1) - 0.001           ''今回の加圧時間が　上限と下限内側なら　T係数は　変えない！
        End If
'     --- 表示 ---
        kjdisp = kjdisp & Format(T_keisu_dum, "0.000") & "  " & Format(T_keisu(T_keisuCont(1) - 1), "0.000") & "  "
        List2.AddItem kjdisp, 0
'     ---'次回計算用　データ更新 ----
        For ikat = 3 To 0 Step -1
          kaatsuJ(T_keisuCont(1) - 1, ikat + 1) = kaatsuJ(T_keisuCont(1) - 1, ikat)
        Next ikat
      End If                ' ---------------------- 加圧時間制御ルーチン　end
'
'     --- 加圧時間自動制御　実施/pass　---
'
      If ((katCflag = True) And (kaatsuJ(T_keisuCont(1) - 1, 3) > 0) And (iflgKataTorF(T_keisuCont(1) - 1) = True)) Then T_keisu(T_keisuCont(1) - 1) = T_keisu_dum
'      If ((katCflag = True) And (kaatsuJ(T_keisuCont(1) - 1, 3) <> 0) And (iflgKataTorF(T_keisuCont(1) - 1) = True)) Then T_keisu(T_keisuCont(1) - 1) = T_keisu_dum
'
      Label4(T_keisuCont(1) - 1).Caption = Format(T_keisu(T_keisuCont(1) - 1), "0.000")
'
'　 --- /*　現在成形中金型の 型No 確認　20190501 sf  ---
'　　　　　　　　　'　3は成形室　　label13(3)
        ikn = katamax - 3 + kataNoPnt + katamax + 1 + Val(kataNo(10))
'
        For iii = 1 To 4
            If ikn > katamax Then ikn = ikn - (katamax + 1)
        Next iii
         
 '--- /* 　カウントアップ　---/*
        If (kataNo(ikn) <> "") Then ShotSu(ikn) = ShotSu(ikn) + 1
'
 '--- /* 　shot数の画面グラフ内表示　---/*
        dumlbl14 = kataNo(0) & "=" & Format(ShotSu(0), "0") & "  " & kataNo(1) & "=" & Format(ShotSu(1), "0") & "  "
        dumlbl14 = dumlbl14 & kataNo(2) & "=" & Format(ShotSu(2), "0") & "  " & kataNo(3) & "=" & Format(ShotSu(3), "0") & "  "
        dumlbl14 = dumlbl14 & kataNo(4) & "=" & Format(ShotSu(4), "0") & "  " & kataNo(5) & "=" & Format(ShotSu(5), "0") & "  "
        dumlbl14 = dumlbl14 & kataNo(6) & "=" & Format(ShotSu(6), "0") & "  " & kataNo(7) & "=" & Format(ShotSu(7), "0")
        Label14.Caption = dumlbl14
'
'　 --- /*　成形データの表示（リスト表示）　*/  csv 2019.4.28 sf  ---
'        InitDat(11)=成形回数（ショット数）
'
      Rec_of_Mold = Format(InitDat(11), "000")
      Rec_of_Mold = Rec_of_Mold & ", " & kataNo(ikn) & ", " & Format(ShotSu(ikn), "0")
      Rec_of_Mold = Rec_of_Mold & ", " & Format(z(iz3), "000.00")
      Rec_of_Mold = Rec_of_Mold & ", " & Format(Int(ct_temp(0)), "000") & "℃, " & Format(Int(ct_temp(1)), "000") & "℃"
      Rec_of_Mold = Rec_of_Mold & ", " & Format(Int(cc_time(1) / 60), "0") & ":" & Format(Int(cc_time(1)) Mod 60, "00")
      Rec_of_Mold = Rec_of_Mold & ", " & Format(Int(cc_time(2) / 60), "0") & ":" & Format(Int(cc_time(2)) Mod 60, "00")
      Rec_of_Mold = Rec_of_Mold & ", " & Format(Int(cc_time(3) / 60), "0") & ":" & Format(Int(cc_time(3)) Mod 60, "00")
      diTime1 = diffTime(cc_time(3), cc_time(2))
      Rec_of_Mold = Rec_of_Mold & ", " & Format(Int(diTime1 + 0.5), "000") & "s"
      Rec_of_Mold = Rec_of_Mold & ", " & Format(cp_z, "000.000")
      Rec_of_Mold = Rec_of_Mold & ", " & Format(Int(stime / 60), "0") & ":" & Format(Int(stime) Mod 60, "00")
      Rec_of_Mold = Rec_of_Mold & ", " & Format(T_keisu(T_keisuCont(1) - 1), "0.000") & ", " & Format(Z3_Hosei(T_keisuCont(1) - 1), "0.000")
      Rec_of_Mold = Rec_of_Mold & ", " & Format(avekatJ(T_keisuCont(1) - 1), "000") & ", " & Format(iHoonStopNo, "0000")
      List1.AddItem Rec_of_Mold, 0                                                                                            ' ”、0”　追加　2004.8.18
        
      RecDtSave Rec_of_Mold
'
'
'' /* 温度係数、肉厚補正データのカウントアップ
      Label4(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(2)  '  文字色を元に戻す
      Label6(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(2)  '  文字色を元に戻す
      Label11(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(2)  '  文字色を元に戻す
      Label4(T_keisuCont(1) - 1).BorderStyle = 0  '  枠なしに戻す
      Label6(T_keisuCont(1) - 1).BorderStyle = 0  '  枠なしに戻す
      Label11(T_keisuCont(1) - 1).BorderStyle = 0  '  枠なしに戻す
'     *** Z3の値を　戻す
          z(iz3) = z(iz3) - Z3_Hosei(T_keisuCont(1) - 1) '  ”Z3"の補正値reset
'     *** ポインターカウントアップ
      If (i_s > 0) And (i_s <> I_s0) Then
        T_keisuCont(1) = T_keisuCont(1) + 1       ' ポインターのカウントアップ
      End If
      If T_keisuCont(1) > (T_keisuCont(0)) Then T_keisuCont(1) = 1
'
      T_keisuCont(2) = T_keisuCont(1)           ' ** ポインターのBuckup **
      T_keisuCont(3) = T_keisuCont(0)           ' ** 型個数　のBuckup **
'       --- Saikaiflg 　を　false　へ
      Saikaiflg = False
'/* データの保存　*/
      If lDtSaveFlg = True Then
        iDtSaveCount = iDtSaveCount - 1
        If kataNo(ikn) <> "" Then ResDtSave i_s, stime
        If iDtSaveCount <= 0 Then
          lDtSaveFlg = False          'データセーブ　受付解除
          Command2(5).BackColor = CmndColoff(1)    ' off gray
          Command2(5).Caption = "GraphDataSave"
        End If
      End If
'
' ScreenCopy iflgSCopy=True の場合、ScreenCopy
    If iSeikeiTorF_flg = True Then
        If iflgSCopy = True Then
            Call SaveWindowPic(True, False)     'Active Windowの保存
        End If
        iflgSCopy = False          'ScreenCopy　受付解除
        Command2(2).BackColor = CmndColoff(1)
    End If

' ---  成形回数　指定の確認
'
      If (seikeiKaisu <> 0) Then
        If (i_s > 0) Then s_kaisu = s_kaisu - 1
        If s_kaisu = 0 Then EditFlg = True
      End If
'
' ---
' '/* エディとが押されていたら　エディット */
      If FrmMenuFlg = False Then Exit Do            '終了が押されているとメニューから抜けるときfalse
      If EditFlg = True Then 'エディタ起動
         ied = 1             'エディタ起動は　doLoopの外で実施　06.3.3 sf
         Exit Do
      End If
'/* 自動停止状態であれば停止 */
      sts1 = SystemReadyChk()   'システムレディ or 非常停止
      sts2 = AutoChk()          '自動状態？
      If sts1 = 0 Or sts2 = 0 Then
        gemgmsg = gemgmsg + ArmEmgMsgChk$()                 ' '08.7.15 'gemgmsg +' 追加
         
        iFlg_hijyou = 10            '非常停止時の情報セーブ
        FrmEmg.Show 1               '　非常停止表示
        FrmMenuFlg = False              'メニューから抜けるときfalse
        NextView = 1
        SeikeiOFF        '非常停止時の処置 '成形OFF　待機中
        HeatOFF          '非常停止時の処置
        CoolOFF          '非常停止時の処置
        ServoOFF         '非常停止時の処置
        Exit Do
      End If
  Loop    '------------------------------------ DO LOOP　　（一番外のループ）
'/*　ｅｄｉｔのときは　do　Loopから抜ける　変更　060303 s.f
'/*  エディットが押されていたら 　ied=1　*/
  If ied = 1 Then 'エディタ起動
      Command2(3).BackColor = CmndColoff(3)  '色を戻す
      EditFlg = False      'エディタ起動解除
      MYEdit.Show 1
      ied = 0
      c = 0
      GoTo st:             '/* エディットモードであれば　ｓｔにジャンプ */
  End If
'/* エディットモードであれば　ｓｔにジャンプ */
'  If ied <> 0 Then GoTo st:
'
'   そうでなければ終了へ
'/* 予備加熱をゼロにし、ＯＦＦする */
eend:
  If iFlg_hijyou > 0 Then              '非常停止から来た時
    RecEmgDtSave sdt3$, sdt1$, sdt2$ & gemgmsg
  End If
  SeikeiOFF          '成形OFF　待機中
  HeatOFF
  CoolOFF
  ServoOFF
'/* coxデータのＨＤへの書き出し */
'    正常終了時  ｺﾝﾄﾛｰﾙﾃﾞｰﾀのsave
      coxDtSet
      coxDtSave gcoxFldir & gcoxFlName
      '
      RecDtSave999                 '   成形データファイルへ　成形プロセスデータを書き込んで終了
' ---  成形回数指定の　リセット
    seikeiKaisu = 0
'' ---
  If FrmMenuFlg = False Then             'メニューから抜けるときfalse
    FrmMenuFlg = True                    'メニューから抜けるときfalse
    Select Case NextView
    Case 1
      Unload Me
      PGM_Menu.Show
    Case 2 '成形（シングル）
      LS21_SC.Show
    Case 3  '成形（ダブル）
    Case 4  'I O チェック
      IOChk.Show
    Case 5  'スケール変更
      LS21_GphScale.Show
    Case 6  '読み出し
    Case 7  'メモ帳
    Case 8  'edit
      Unload Me
      MYEdit.Show
    Case Else
      Unload Me
      PGM_Menu.Show
    End Select
  End If
  If iFlg_hijyou = 0 Then Unload Me       '非常停止から来た時は、消さない
  PGM_Menu.Show
'
Exit Sub
'
errHandler:
  SeikeiOFF          '成形OFF　待機中
  HeatOFF
  ServoOFF
  CoolOFF
'
  RecEmgDtSave sdt3, sdt1, sdt2
  If Err.Number <> 0 Then
     sdt1 = "エラー番号：" & Err.Number
     sdt2 = "ﾌﾟﾛｼﾞｪｸﾄ名：" & Err.Source & "  " & ppos
     sdt3 = "エラー内容：" & Err.Description
  End If
  RecEmgDtSave sdt1, sdt2, sdt3
  gemgmsg = gemgmsg + Err.Number & Err.Description      '  08.7.15 gemgmsg +  追加
  hijyou        '非常停止処理
'
End Sub
Private Sub genten()
'--------------
  C870Genten
  gOrgFlg = True                       '原点復帰完了=TRUE
  OrgON
  gOrgStartFlg = True   '2002.10.17 KYOCERA
End Sub

Private Sub GphXSet()
Dim i%
  For i = 0 To ptime * 60 + 10
    TPass(i) = i
  Next i
End Sub

Private Sub GphDataSet(i0%, i1%)
Dim i%
  For i = i0 To i1
    Templ(i) = atemp(i, 0)
    Templd(i) = atemp(i, 1)   '下型温度
    Press(i) = apre(i)
    ZAxis(i) = aposi(i)
  Next i
End Sub

Private Function DispSegm$(ist0%)
Dim sdt$
    If ist0 < 0 Then Exit Function
    sdt = Right(Space(2) & Format(ist0, "0"), 2)
    sdt = sdt & Right(Space(4) & Format(seg_num(ist0), "0"), 4)
    sdt = sdt & Right(Space(4) & Format(ic(ist0), "0"), 4)
    sdt = sdt & Right(Space(12) & Format(z(ist0), "0.000"), 12)
    sdt = sdt & Right(Space(7) & Format(vel(ist0), "0.0"), 7)
    sdt = sdt & Right(Space(6) & Format(pres(ist0), "0"), 6)
    sdt = sdt & Right(Space(7) & Format(t0(ist0), "0.0"), 7)
    sdt = sdt & Right(Space(7) & Format(p(ist0), "0.0"), 7)
    DispSegm = sdt
'    Label2(12).Caption = sdt
End Function
Private Function EmgChk%()
Dim ch%, sts%
  ch = 1
  DioInput ch, sts
  If sts = 0 Then
    EmgChk = True
  Else
    EmgChk = False
  End If
End Function

Private Sub Timer2_Timer()
    If r_z > 0.1 Then
        OrgOFF
    Else
        OrgON
    End If
    
    'Label6(0).Caption = "原点 = " & gOrgIL
    'Label6(1).Caption = r_z
End Sub
'スクリーンのスナップショットをクリップボードに保存及び印刷　本体　　　　　（273） '

Private Sub SaveWindowPic(Optional ActWind As Boolean = True, _
                                    Optional PrintOn As Boolean = False)
'スクリーンのスナップショットをクリップボードに保存及び印刷　　　　　　　　　（273） '
'フォームにCommandボタンを２個貼り付けておいて下さい。
'　 Option Explicit　　 'SampleNo=273　WindowsXP VB6.0(SP5) 2003.03.30
'キーストロークをシミュレートする(P1065)

    Dim MyFileName As String, PicData As Picture, OsVer As Single
    Dim sngSt As Single
'
    Clipboard.Clear
    OsVer = CreateObject("SysInfo.SYSINFO").OSVersion

    If ActWind Then
    'アクティブ ウィンドウのスナップショットを取得する
    '以下の２方法どれでもOK(Win98SE/WinXP/Win95）
    'どの方法でも上記確認機種は同じ動作しますのでMSのサンプルの方法を使用
        Call keybd_event(VK_LMENU, &H56, _
                                KEYEVENTF_EXTENDEDKEY Or 0, 0)
        Call keybd_event(VK_SNAPSHOT, &H79, _
                                KEYEVENTF_EXTENDEDKEY Or 0, 0)
        Call keybd_event(VK_SNAPSHOT, &H79, _
                                KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
        Call keybd_event(VK_LMENU, &H56, _
                                KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
'　　　　==================== こちらでも同じようです ==================
'　　　　Call keybd_event(VK_LMENU, 0, _
　　　　　　　　　　　　　　　　KEYEVENTF_EXTENDEDKEY Or 0, 0)
'　　　　Call keybd_event(VK_SNAPSHOT, 0, _
　　　　　　　　　　　　　　　　KEYEVENTF_EXTENDEDKEY Or 0, 0)
'　　　　Call keybd_event(VK_SNAPSHOT, 0, _
　　　　　　　　　　　KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
'　　　　Call keybd_event(VK_LMENU, 0, _
　　　　　　　　　　　KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    ElseIf ActWind = False And OsVer < 5 Then
    '画面全体のスナップショットを取得する(Win98SE/Win95)
        Call keybd_event(VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY, 0)
        Call keybd_event(VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY Or _
                                                                          KEYEVENTF_KEYUP, 0)
    Else
    '画面全体のスナップショットを取得する(WinXP)
        Call keybd_event(VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0)
        Call keybd_event(VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY Or _
                                                                          KEYEVENTF_KEYUP, 0)
    End If
'
    sngSt = Timer                           ' Windows7 には、この遅延Loopが必要
    Do While Timer - sngSt < 0.5
       DoEvents
    Loop
    'クリップボード内にビットマップ形式のデータがあるか調べる
    If Clipboard.GetFormat(vbCFBitmap) Then
        'ファイル名を自動生成
        MyFileName = App.path & "\..\data\" & gcoxFlName$ & Format$(Now, "yymmddhhmmss") & ".BMP"
        '表示データーをビットマップ形式のデータで保存
        Set PicData = Clipboard.GetData
        Call SavePicture(PicData, MyFileName)
        If PrintOn Then
            '印刷する場合
            With Printer
                .ScaleMode = vbMillimeters
                .PaperSize = vbPRPSA4
                .Orientation = vbPRORLandscape
                .PaintPicture PicData, 10, 0
                .EndDoc
            End With
        End If
    Else
        MsgBox "保存出来ませんでした。"
    End If
End Sub
'
'
'
'Private Sub Command1_Click()
''アクティブウインドウのみをクリップボードにコピー
'    Call SaveWindowPic(True, False)     '印刷する場合は　True に設定
'End Sub
'
'Private Sub Command2_Click()
''スクリーン全体をクリップボードにコピー
'    Call SaveWindowPic(False, False)
'End Sub




'NQD Vbの場合
'＜NQD70_SC＞へ追加
'フラグの追加 冒頭の宣言部
'Dim iflgSCopy As Boolean   ' ScreenCopy フラグ
'
'＜コマンドボタンの追加＞
'Private Sub Command3_Click()
''''アクティブウインドウをクリップボードにコピー印刷する。　True に設定
'  If iflgSCopy = True Then
'          iflgSCopy = False          'ScreenCopy　受付解除
'          Command3.BackColor = CmndColoff(0)
'    Else
'          iflgSCopy = True      'ScreenCopy　受付
'          Command3.BackColor = CmndColon(1)    ' on 1=red
'  End If
'End Sub
'
'<NQD70_SCの本体への　call文追加＞
'>'/* データの保存　*/
'>      If lDtSaveFlg = True Then
'>        ResDtSave i_s, stime
'>        lDtSaveFlg = False
'>      End If
''
'' ScreenCopy iflgSCopy=True の場合、ScreenCopy
'    If iflgSCopy = True Then
'        Call SaveWindowPic(True, False)     'Active Windowの保存
'    End If
'    iflgSCopy = False          'ScreenCopy　受付解除
'    Command3.BackColor = CmndColoff(0)
''



