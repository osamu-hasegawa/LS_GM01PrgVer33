VERSION 5.00
Begin VB.Form PGM_Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "メニュー"
   ClientHeight    =   6408
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   8304
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6408
   ScaleWidth      =   8304
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command1 
      Caption         =   "空成形−排出"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Index           =   4
      Left            =   1260
      TabIndex        =   38
      Top             =   2680
      Visible         =   0   'False
      Width           =   2244
   End
   Begin VB.CommandButton Command1 
      Caption         =   "連続成形再開"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Index           =   1
      Left            =   1260
      TabIndex        =   37
      Top             =   2040
      Visible         =   0   'False
      Width           =   2244
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   492
      Left            =   6000
      TabIndex        =   36
      Top             =   4200
      Width           =   372
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   240
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   0
      Left            =   6360
      LinkTimeout     =   2
      TabIndex        =   33
      Text            =   "2"
      Top             =   4320
      Width           =   732
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'ﾌﾗｯﾄ
      Caption         =   "搬送開始"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   4320
      TabIndex        =   32
      Top             =   4200
      Width           =   1236
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
      Index           =   0
      Left            =   240
      TabIndex        =   29
      Top             =   5280
      Visible         =   0   'False
      Width           =   1236
   End
   Begin VB.CommandButton Command1 
      Caption         =   "カウンタリセット"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Index           =   8
      Left            =   1260
      TabIndex        =   24
      Top             =   3960
      Width           =   2244
   End
   Begin VB.CommandButton Command1 
      Caption         =   "G  原点出し実行"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   240
      TabIndex        =   22
      Top             =   5760
      Width           =   2244
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   324
      Top             =   216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "スケール変更"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Index           =   6
      Left            =   1260
      TabIndex        =   19
      Top             =   3360
      Width           =   2244
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
      Index           =   6
      Left            =   6444
      TabIndex        =   18
      Top             =   5760
      Width           =   1236
   End
   Begin VB.CommandButton Command2 
      Caption         =   "edit"
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
      Index           =   5
      Left            =   5148
      TabIndex        =   17
      Top             =   5760
      Width           =   1236
   End
   Begin VB.CommandButton Command2 
      Caption         =   "メモ帳"
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
      Left            =   3852
      TabIndex        =   16
      Top             =   5760
      Width           =   1236
   End
   Begin VB.CommandButton Command2 
      Caption         =   "読出し"
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
      Left            =   2556
      TabIndex        =   15
      Top             =   5760
      Width           =   1236
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I O チェック"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   5
      Left            =   5640
      TabIndex        =   6
      Top             =   3600
      Width           =   1524
   End
   Begin VB.CommandButton Command1 
      Caption         =   "データ出力"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1644
   End
   Begin VB.CommandButton Command1 
      Caption         =   "連続成形"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Index           =   2
      Left            =   1260
      TabIndex        =   4
      Top             =   1440
      Width           =   2244
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1回成形"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LS-33_Ver190501e"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   252
      Index           =   13
      Left            =   2880
      TabIndex        =   35
      Top             =   600
      Width           =   2892
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "回"
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
      Index           =   12
      Left            =   7200
      TabIndex        =   34
      Top             =   4320
      Width           =   276
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
      Index           =   7
      Left            =   5628
      TabIndex        =   31
      Top             =   3240
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "速度設定電圧"
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
      Index           =   8
      Left            =   4080
      TabIndex        =   30
      Top             =   3240
      Width           =   1524
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "回"
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
      Left            =   7320
      TabIndex        =   28
      Top             =   1596
      Width           =   276
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
      Index           =   6
      Left            =   5652
      TabIndex        =   27
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "ショット数Ｔ"
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
      Index           =   10
      Left            =   3960
      TabIndex        =   26
      Top             =   1560
      Width           =   1548
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   5
      Left            =   1680
      TabIndex        =   25
      Top             =   5280
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   4
      Left            =   3885
      TabIndex        =   23
      Top             =   5280
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "時間"
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
      Index           =   9
      Left            =   4824
      TabIndex        =   21
      Top             =   2892
      Width           =   516
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
      Index           =   3
      Left            =   5652
      TabIndex        =   20
      Top             =   2856
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
      Index           =   2
      Left            =   2850
      TabIndex        =   14
      Top             =   4725
      Width           =   4560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "制御ファイル名："
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
      Index           =   7
      Left            =   825
      TabIndex        =   13
      Top             =   4725
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Ｋｇ"
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
      Index           =   6
      Left            =   7272
      TabIndex        =   12
      Top             =   2496
      Width           =   516
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "ｍｍ"
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
      Index           =   5
      Left            =   7272
      TabIndex        =   11
      Top             =   1992
      Width           =   516
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
      Index           =   1
      Left            =   5652
      TabIndex        =   10
      Top             =   2460
      Width           =   1500
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
      Index           =   0
      Left            =   5652
      TabIndex        =   9
      Top             =   1992
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "圧力"
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
      Index           =   4
      Left            =   4824
      TabIndex        =   8
      Top             =   2496
      Width           =   516
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Ｚ位置"
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
      Index           =   3
      Left            =   4644
      TabIndex        =   7
      Top             =   2028
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "モニタ"
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
      Left            =   6000
      TabIndex        =   2
      Top             =   1080
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "成  形"
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
      Index           =   1
      Left            =   1875
      TabIndex        =   1
      Top             =   1050
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Precision Glass Mold System"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   396
      Index           =   0
      Left            =   1548
      TabIndex        =   0
      Top             =   144
      Width           =   5508
   End
End
Attribute VB_Name = "PGM_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    PGM_Menu
'
'            update: 2002.8.10 s.f roz(0),roz(1)を突当成形のﾊﾟﾗﾒｰﾀへ'
'            update: 2002.10.16 KYOCERA ﾒﾆｭｰ画面起動時の原点信号出力ON→OFF
'　　　　　　　　　　　　　　　　　　　　　"原点"完了後にOrgON追加
'            update: 2002.10.17 KYOCERA 原点復帰後に初回原点復帰完了ﾌﾗｸﾞgOrgStartFlgをON
'                                       原点信号をﾀｲﾏｰで監視
'                                       原点でないと自動成形ﾓｰﾄﾞ移行不可
'            update: 2002.10.18 KYOCERA 原点表示の修正 If gOrgStartFlg = False Then...End If追加
'            update: 2002.10.25 s.f. Ver．表示修正
'            update: 2002.10.26 s.f. 「真空到達」無効へ
'            update: 2003. 8.26 s.f. * 指定圧力＋２００Ｋｇ以上で非常停止 *
'            update: 2003. 9.11 s.f. LS21_TC　成形終了時の非常停止エラー対策
'            update: 2003. 9.12 s.f. genten()　原点出し後　HiSpeedを指定値に戻す。
'
'            update: 2003.12.15 s.f. LS-32立上げに伴う変更　MplDef.bas　のみ　新規　2003.11.04付け
'　　　　　　　　　　　　　　　　　　これに伴い　PGM_MenuのVERﾅﾝﾊﾞｰを　LS-32　へ変更
'
'            update: 2004. 3. 8 s.f. LS21_SC 変更　成形軸制御モード　’７’追加　（上軸衝突判定付）
'                                    RecEmgDTsave 非常停止メッセージの保存
'            update: 2004. 3.12 s.f.  速度指令電圧　表示
'            update: 2004. 3.20 s.f.  LS31へ移植　MplDef.basのみ　旧Ver　2002.1.13付けへ戻す。
'
'            update: 2004.3.20  s.f. MYEdit.frm　の　SetData(),GetData()　を変更（3/8変更のバグ修正　'edit'の読み込み書き出しエラー）
'　　　　　　　　　　　　　　　　　　　軸制御ｺﾏﾝﾄﾞ　7追加： 現在有効コマンド 0,1,2,3,7,8,9
'
'            update: 2004. 3.30 s.f   非常停止ﾒｯｾｰｼﾞバグ修正
'            update: 2004. 4.23 s.f   timeupで非常停止
'            update: 2004. 4.24 s.f.  LS21_TC内のカウンタ、ﾀｸﾄﾀｲﾑ、表示　改造
'
'            update: 2004.4.25  s.f   Myedit　の  VScroll1(j).min = 210 * lK1     "200"を"210"へ変更
'            update: 2004.5. 5  s.f   温度係数、肉厚補正ルーチン　追加  PGM_KTD,My_lib,MYEDIT, LS21_SC, LS21_TC
'            update: 2004.5.12  s.f   PGM_KTD　"ｵｰﾊﾞｰﾌﾛｰ"対策　　wTm0!,wTm1!  global化,  LS21_SCと　LS21_TC から　dim削除
'            update: 2004.5.17  s.f   'S'ｺﾏﾝﾄﾞ　バグ対策
'            update: 2004.5.18  s.f   バグ対策 & T係数表示
'            update: 2004.6. 5  s.f   「Vエディット」非表示色変更
'            update: 2004.8.17  s.f   ｵｰﾊﾞｰﾌﾛｰ"対策  p(ist0)をppへ  ”：”複数の行を無くす
'            update: 2004.8.27 - 10.30 s.f   T係数関数変更、0.01=1℃ 「ＤＣ　０」コマンド　成形前に型在否チェックセンサーのチェック機能追加
'            update: 2004.10.30 s.f   成形プロセスグラフ表示　温度表示色　緑色へ変更
'            update: 2004.11.2 s.f     T係数関数変更　元へ戻す。
'            update: 2004.12.20 s.f    LS21_TC  DCコマンド　　バグ修正
'            update: 2005. 5.25 s.f    Version No表示追加
'            update: 2005. 7.18 s.f    加圧時間　平均値表示,1回成形後の冷却追加
'            update: 2005. 7.25 s.f    加圧時間制御のデバッグ
'            update: 2005. 9.27 s.f   保温停止モード　追加
'            update: 2005. 9.28 s.f   T係数　表示色変更
'            update: 2005. 9.28a s.f  上記デバッグ  型がない時は　保温停止　実施しない
'            update: 2005.11. 4 s.f  LS21_SC　表示変更。速度制御電圧表示削除。T係数、Z３補正表示部変更,加圧時間制御バグ修正
'            update: 2005.11. 6 s.f   オーバーフロー対策 idc65536,idc256,ddc05, my_lib 書替え　long,double指定へ
'                                      Mpldef 変更　C870contini
'            update: 2005.11.22 s.f   Melec C-870 counter動作バグ修正　コンペアカウンタ値セット時　符号反転　　setcm1
'                                     オーバーフローエラー対策　idc16777216、idc8388607　追加
'            update: 2005.11.23 s.f   11/22 変更のバグ修正　成形軸制御　「C870sts　resetするまで　読み飛ばす」を　復活
'　　　　　　　　　　　　　　　　　　画面下表示　シンプル化　（スピード低下防止の為）
'            update: 2005.11.26 s.f   すべての　function　に　型宣言をつける　　　overflow対策
'                                     すべての　sub　の引数に　型宣言をつける
'                                     sdata,
'            update: 2005.12.17 s.f  LS21-SC,  LS21-TC 変更 、　最近頻発の timeup 対策
'                                    Do-Loop 外の　DoEvent削除 OverFlow 対策 s.f.
'                                    コマンドの　evtime　取り込みを　コマンド開始時へ変更
'　　　　　　　　　　　　　　　　　　DCコマンド　LAコマンド　再チェック修正
'　　　　　　　　　　　　　　　　　　連続前コマンド　evtime　と　fintime　表記入れ替え
'
'            update: 2006. 3. 3 s.f  edit 使用時　do　loopから抜ける
'            update: 2006. 4.14 s.f  on error goto を入れる
'            update: 2006. 4.15 s.f  error 表示、搬送回数スクロール指定
'            update: 2006. 5. 9 s.f  O.F.error 表示　軸制御　end3　追加,  tstime=0#
'            update: 2006. 5.14 s.f 　r_pres()の　DoEvents 　 forの外へ移動　s.f  ものすごく効く
'　　　　　　　　　　　　　　　　　  すべて抜くと　LS_TC　プログラム暴走する（LS_SCは　OK)’
'            update: 2006. 5.15 s.f  5分間保温停止　追加
'            update: 2006. 5.18 s.f 　r_pres()の　DoEvents 　削除、　”J"、”S"に　Doevents　追加
'                                     myEdit へ　LA、DC　追加
'            update: 2006. 5.19 s.f 　My_edit内から　メモ帳　呼び出し、　myeditの　DC　削除
'            update: 2006. 5.23 s.f 　cal_pid 変更 overFlow 対策
'            update: 2006. 5.26 s.f 　AdRead ppos ツイカ
'            update: 2006. 7.12 s.f   My_lib  r_z!()  w1,w2,w3 long → integer  (overflow 対策) これが真因か？
'            update: 2006. 7.12 s.f  加圧時間自動調整　’有効’へ
'            update: 2006. 8. 2 s.f  「1回成形」冷却時間カウントダウン　バグ修正
'            update: 2006.12.21 s.f  「1回成形」冷却時間カウントダウン　バグ修正
'       Ver.3.33R_061221 2006.12.21 s.f  LS-33改　対応　　VacuumON、VacuumOFF　を廃止、SeikeiON,SeikeiOFF新設　DO3　割り当て変更
'       Ver.3.33R_070827 2007.08.27 s.f  非常停止時の　処置追加
'       Ver.3.33R_070927 2007.09.27 s.f  Z補正　指定したｾｸﾞﾒﾝﾄNo.へ　できるようにする
'       Ver.3.33R_071113 2007.11.13 s.f  「強制ソーク」復活、　「1回成形」enable=Falseへ
'       Ver.3.33R_071119 2007.11.19 s.f  加圧時間制御　バグ修正（edit時、データ継承）、平均値AND最新値で　更新判定へ
'　　　　　　　　　　　　　　　　　　　　加圧時間制御　ON時は、　T係数データの　ファイルからの読み込みしない
'       Ver.3.33R_071120 2007.11.20 s.f  バグ修正、　空成形-排出　追加、　連続成形再開　追加
'       Ver.3.33R_071120 2007.11.21 s.f  加圧制御　平均値計算　今回の加圧時間　重み2.0へ
'                                        -a : 型順　表示　調整数追加
'       Ver.3.33R_071126 2007.11.26 s.f  型順　表示バグ修正
'       Ver.3.33R_071127 2007.11.27 s.f  型順　表示バグ修正
'       Ver.3.33R_071127a 2007.11.30 s.f  SaikaiFlg, ISokuFlag バグ対策：　FrmMenuFlg=Falseの場所入れ替え
'       Ver.3.33R_071203 2007.12.03 s.f  非常停止　メッセージ追加、変更
'       Ver.3.33R_071203a 2007.12.06 s.f  型No.ポインター　初期値　1　→　0
'       Ver.3.33R_071218 2007.12.18 s.f  「非常停止（Ｎｏ．）」　→　「ＰＣ→非常停止（Ｎｏ．１）」 ほか　非常停止表示ルーチンバグ修正
'       Ver.3.33R_071221 2007.12.21 s.f  速度制御値の表示　削除 Label7(0),(1)
'       Ver.3.33R_071224 2007.12.24 s.f  加圧制御　判断基準　今回の加圧時間を考慮へ
'       Ver.3.33R_080114 2008. 1.14 s.f  強制ソークﾌﾗｸﾞ　バグ修正
'       Ver.3.33R_080218 2008. 2.18 s.f  軸制御７　F*0.7　→　F*1.0　で　判定へ変更
'       Ver.3.33R_080221 2008. 2.21 s.f  軸制御1,3,7 PCでのZ行過ぎチェック　１秒に１回へ変更。
'                                       　Z3補正の　No.変更　editでは　禁止とする。（エディターで変更可）
'       Ver.3.33R_080223 2008. 2.23 s.f  上記変更のバグ修正。
'       Ver.3.33R_080304 2008. 3. 4 s.f  FbiDA,FbiAD拡張
'       Ver.3.33R_080608A 2008. 6. 8 s.f  LS１２　改造中版
'       Ver.3.33R_080714A 2008. 7.14 s.f  LS１２　改造中 その２　予備加熱３　電源追加
'       Ver.3.33R_080715A 2008. 7.14 s.f  "gemgmsg + " 追加
'       Ver.3.33R_080817A 2008. 8.17 s.f  型数　変更版へ katano、　型名称の　ディスク格納
'       Ver.3.33R_090921R 2009. 9.21 s.f  成形データへ　成形プロセスデータｓａｖｅを追加（recdtsave999)
'       Ver.3.33-12R_091226R 2009.12.26 s.f  エディタでNptepad使用のバグ修正。型順表示　成形室ピンク色
'       Ver.3.33-12R_091229R 2009.12.29 s.f  型順表示　成形室 薄朱色へ
'       Ver.3.33-12R_091230R 2009.12.30 s.f  型順表示　冷却　水色へ
'       Ver.3.33-12R-100304R 2010. 3. 4 s.f  初回ポインターカウントアップバグ対策　　成形有効無効判断から　i_s=0のときpassを削除。
'       Ver.3.33-12R-100310R 2010. 3.10 s.f  Timer値異常時　timeup処理の　skip実施。
'       Ver.3.33-12R-100412R 2010. 4.12 s.f  timeup処理の　skip ifに　ic(ist0)<10を追加。ﾎﾟｲﾝﾀｰｶｳﾝﾄｱｯﾌﾟのバグ修正のバグ修正、「成形の有効性判定」を　初回は別枠にする。
'       Ver.3.33-12R-100416R 2010. 4.16 s.f  成形回数指定の追加。 seikeiKaisu, s_kaisu
'       Ver.3.33-12R-111228R 2011.12.28 s.f  保温停止　復活
'       Ver.3.33-12R-120104R 2012.01.04 s.f  LAから進まないバグ修正、hotei.formへ　Dim fintime! 追加（なぜか抜けていた）
'       Ver.3.33-12R-120105  2012.01.05 s.f  LAから進まないバグ本当に修正。完成版
'       Ver.3.33-12R-120108  2012.01.08 s.f  「5分停止」効かないバグ修正
'       Ver.3.33-12R-120422  2012.04.22 s.f  screen　copy追加
'       Ver.3.33-12R-120430  2012.04.30 s.f  screen　copy　無効成形除外判断追加
'       Ver.3.33-12R-120513  2012.05.13 s.f  z3肉厚補正　入力制限変更　元:　+/-0.05　　新：　+0.05/-0.50
'       Ver.3.33-12R-120522  2012.05.22 s.f  PC→非常停止へ   gemgmsg = gemgmsg + "cal_pid" + Format(pa, "0.0")   '2010.5.19 s.f.追加 12R-120122R版対応のため」
'       Ver.3.33-12R-120522a 2012.07.01 s.f  cal_pid 異常判定　「＋500kgf、-100kgf」→「+500kgf、-200kgf」へ
'       Ver.3.33-12R-190428  2019.04.28 s.f  成形ショット数積算追加、Graphデータ取得新設、
'                                            30分枠を６０分へ拡張,KeikaTime追加
'       Ver.3.33-12R-190428a 2019.04.28 s.f  6ST/4ST共用化  katamax=4or6 をPGM_KTD.basで必ず指定すること！
'   　　Ver.3.33-12R-190428b 2019.05.01 s.f  MyEditのGraph　Picture1内のLine1の配列数を　21から30へ変更　Line1(0 to 30）
'                                             LS33-110804　共通化のため
'   　　Ver.33-190501 2019.05.01 s.f  12R190428b と　33R-110804　の共通化　ST数　7,6,4　st
'   　　Ver.33-190501a 2019.05.05 s.f  bug取り　datasaveボタン見えない→見える（visible=True）
'   　　Ver.33-190501b 2019.05.07 s.f  bug取り　表示位置バグ修正
'   　　Ver.33-190501c 2019.05.11 s.f  bug取り　表示位置バグ修正, katamax数により、型順表示枠を消す。MyEditの表示も消す
'                                      iknの成形室位置判断　if文追加, 型順表示位置（プログラム内の位置）高速ループ内から　1回の成形の開始時へ移動
'                                      coxデータのＨＤへの書き出し（毎回）
'   　　Ver.33-190501d 2019.05.20 s.f  bug取り　表示位置バグ修正 型順書き換え場所を”XR"へ戻す。成形室に型があるときのみショット数カウントアップ（idcflg1=0）
'                                       katahyjの書き換えcoxファイルread後へ場所変更
'   　　Ver.33-190501e 2019.05.25 s.f  成形データ表示の文字間へスペース２個へ ", " → ",  "
'                                      待ち時間入力制限　1000”→2400”　へ
'
' ******************************************************************************
Option Explicit
'
Dim di_d2%         '/* DIO_P　2ﾎﾟｰﾄ　ﾊﾞｯﾌｧ */
Dim OrgFlg%         '原点出し
Dim MemoFlg%        'メモ帳
Dim NextView%
Dim TrnsMax%        '搬送回数
Dim TrnsCnt%        '搬送カウンタ
Dim lTrnsFLg%       '搬送中フラグ
Dim lK1%            '回数カウンタ

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0  '１回成形
  NextView = 1                          '１回成形（オンリー）
  FrmMenuFlg = False                    'メニューから抜けるときfalse
Case 1  '連続成形再開                 ' 20071119 tsuika
  NextView = 2                          '連続成形再開（シングル） ' 20071119 tsuika
  lSokuFlg = True
  Saikaiflg = True                      ' 再開フラグ = true
  FrmMenuFlg = False                    'メニューから抜けるときfalse連続成形
Case 2  '連続成形
  NextView = 2                         '連続成形（シングル）
  FrmMenuFlg = False                    'メニューから抜けるときfalse連続成形
Case 3  'データ出力
  NextView = 3                          'データ出力
  FrmMenuFlg = False                    'メニューから抜けるときfalse
Case 4  '空成形-排出
  NextView = 2
  Karauchiflg = True
  FrmMenuFlg = False                    'メニューから抜けるときfalse連続成形
Case 5  'I O チェック
  NextView = 4                           '
  FrmMenuFlg = False                    'メニューから抜けるときfalse
  'Unload Me
  'IOChk.Show 1
  'adMain.Show
  'Sampling.Show
  'OutBox.Show
  'MplVbSmp.Show
Case 6  'スケール変更
  NextView = 5                           '
  FrmMenuFlg = False                    'メニューから抜けるときfalse
Case 7  '原点出し実行
  OrgFlg = True       '原点出し
  'genten
Case 8  'カウンタリセット
  InitDat(11) = 0                 '成形カウンタトウタル
  InitDtSave
End Select
End Sub

Private Sub Command2_Click(Index As Integer)
'
  'FrmMenuFlg = False                    'メニューから抜けるときfalse
  '
  Select Case Index
  Case 0    '真空到達
    gVumFlg = 1                       '真空到達=1
  Case 1    '搬送開始
    If lTrnsFLg = True Then
      lTrnsFLg = False                  '搬送中フラグ
      Command2(1).Caption = "搬送開始"
    Else
      Command2(1).Caption = "搬送中止"
      TrnsMax = Val(Text1(0).Text)      '搬送回数
      lTrnsFLg = True                   '搬送中フラグ
      PltPrns TrnsMax
    End If
  Case 3  '読み出し
    FrmMenuFlg = False                    'メニューから抜けるときfalse
    NextView = 6                           '
    'coxFlLoad
    'Label2(2) = gcoxFlName
    'cfileSave
  Case 4  'メモ帳
    FrmMenuFlg = False                    'メニューから抜けるときfalse
    NextView = 7                           '
    MemoFlg = True      'メモ帳
    'ExecMemo gcoxFldir, gcoxFlName
  Case 5  'edit
    FrmMenuFlg = False                    'メニューから抜けるときfalse
    NextView = 8                           '
    'Unload Me
    'MYEdit.Show 1
  Case 6  '終了
    FrmMenuFlg = False                    'メニューから抜けるときfalse
    InitDtSave
    BoardClose
    End
  End Select
End Sub

Private Sub SetData()

  'Label2(2) = gcoxFlName             '制御ファイル名
  
End Sub

Private Sub Form_Load()
  lSokuFlg = False        '強制ソークタイム   通常時は　OFF
  katCflag = False      ' プログラム開始時は、必ず加圧制御OFF
  Karauchiflg = False      ' プログラム開始時は、一旦false
  Saikaiflg = False         'プログラム開始時は、一旦false
'
  DispCenter Me
  versionNo = Label1(13)            '　VersionNo　表示用
  PGM_Menu.Caption = PGM_Menu.Caption + "     " + versionNo
  ViewFlg = 1                       '画面番号
  FrmMenuFlg = True                   'メニューから抜けるときfalse
  Timer1.Enabled = False
  Me.Show
  Label2(5).Caption = ""            '原点表示
  SetData
  SetVScroll1
  DispText1 2, True       'kaisuu set
  T_keisuCont(2) = 0                ' T係数　ﾎﾟｲﾝﾀｰのbackupｸﾘﾔ
  T_keisuCont(3) = 0                ' 型個数のbackupのｸﾘﾔ
   Timer1.Enabled = True
  Command1(0).Enabled = False       '2002.10.17 KYOCERA
  Command1(1).Enabled = False
  Command1(2).Enabled = False
  Command1(4).Enabled = False
End Sub
'-------------------------------------------------------------
'Private Sub ExecMemo(DDir$, flNm$)      ' 2006.5.19 PGM_KTD へ移動
'Dim ExecFl$, fl$
'Dim r!
'  fl = DDir$ & flNm
'  ExecFl = "C:\WINDOWS\NOTEPAD.EXE " & fl
''-------- メモ帳でflを開く
'  r = Shell(ExecFl, 1)
'  AppActivate r, True     'メモ帳が閉じるまで待つ
'End Sub

Private Sub genten()
Dim hspd As Long
'--------------
  Label2(4).Caption = "原点復帰中"
  Label2(5).Caption = ""
  C870Genten
'/* カウンタにゼロを書き込む */
  Ready_Wait
  C870CntPreSet 0   'ＣＯＵＮＴＥＲ ＰＲＥＳＥＴ ＣＯＭＭＡＮＤ
'/* 手動用　速度へ戻す */
  hspd = gHiSpeed * gRev2Disp / 60              '03.9.12変更
  C870HSPDSet hspd                              '03.9.12変更
  
'  C870HSPDSet 36256    '/* 36256 pps  3mm/sec 　旧　03.9.12変更
  Label2(4).Caption = ""
  gOrgFlg = True                       '原点復帰完了=TRUE
  OrgON                 '2002.10.16 KYOCERA
  gOrgStartFlg = True   '2002.10.17 KYOCERA
End Sub

Private Sub prcom(buf$, im%)
Dim nm$, comm$, fp$
Dim j%, fnum%
Dim dr$, fl$
  dr = App.path & "\..\cont\"
  fl = "prcom.dat"
  If im = 1 Then
    comm = "0"
    fnum = FreeFile
    Open dr & fl For Input As #fnum
      Line Input #fnum, comm
      Line Input #fnum, comm
    Close #fnum
  Else
    fnum = FreeFile
    Open dr & fl For Output As #fnum
      Write #fnum, comm
      Write #fnum, comm
    Close #fnum
  End If
End Sub

Private Sub ginit()
'/* タイトルの表示　*/
End Sub

Private Sub disp_t(ttime$)
  Label2(3).Caption = ttime
End Sub

Private Sub cal_pid(m_sa!, m_p!, m_lim!)
'  float  m_sa,     /* 設定圧力 */
'         m_p,      /* 設定Ｐ値 */
'         m_lim;    /* 設定リミット値 */
Dim i%, nout1%
Dim pa!, per!
Dim ch%, v!
  pa = r_pres()     '/* 圧力 */

  If pa > m_sa + 200# Then '/* 指定圧力＋２００Ｋｇ以上で非常停止 */
    hijyou
    Exit Sub
  End If

'/* ＰＩＤ演算 */

  per = 5 * (m_sa - pa) * Abs(m_sa - pa) / (m_p * m_p)
  If per > m_lim Then per = m_lim
  If per < (-1 * m_lim) Then per = -1 * m_lim
  nout1 = &H800 - Int(4.095 * per / 4#)
  'nout1 = Int(-40.95 * per) + &H800
  ch = 1
  'DaOut1 ch, nout1
  '
  v = per / 5
  'v = per / 5
  Label2(7).Caption = Format(v, "0.00")
  ch = 1
  DaVoltOut ch, v

End Sub

Private Sub qd62_Main()
Dim c$, mc0$, mc1$
Dim cname$, DName$, ttime$, chaz$, chap$, stime$
Dim i%, j%, imo%, ic%, c0%, ndata%
Dim ie02%, ie03%, ie04%
Dim ie%, ie0%, ie1%, ie2%, ie3%, ie4%, ie5%
Dim z!, apre!
Dim roz(0 To 2)                          '突当成形ﾊﾟﾗﾒｰﾀ　幅、時間
Dim fp$
'
Dim ch%, nTime!, g_sts%
Dim hspd As Long
Dim lspd As Long
Dim FlgAuto%
'Dim sdt1$, sdt2$, sdt3$          '  2004.3.30  追加  s.f  2006.4.14 global へ
'
  cname = "cont\\          "
  DName = "data\\          "
  ie02 = 0: ie03 = 0: ie04 = 0
  ie = 0: ie0 = 0: ie1 = 0: ie2 = 0: ie3 = 0: ie4 = 0: ie5 = 0
  z = 0: apre = 0
'/* ＡＴＣ温度リセット */
  ch = 2
  DaVoltOut ch, 0   '常温設定
  ch = 3
  DaVoltOut ch, 0   '常温設定
'/* コントロールファイル名の読み出し */
  cfileLoad
  Label2(2).Caption = gcoxFlName
'/* ロボットデータのワンボードへの転送 */
  rozFileLoad
'/***********     ﾒﾚｯｸ　C-853ボード初期設定　　　*************/
  'DioAllReset
  C870SpecInit    '/* SPEC INITIALIZE CMD OUT */
  C870CntInit     '/* カウンタボードの初期設定 */
  C870AccRate     '/* 加減速ﾚｰﾄｾｯﾄｺﾏﾝﾄﾞ */
  C870DelayTime   '/* ディレータイム設定 */
  ServoON         '/* サーボｏｎ */
  '--------------- 速度の設定
  hspd = gHiSpeed * gRev2Disp / 60
  C870HSPDSet hspd
  lspd = gLwSpeed * gRev2Disp / 60
  C870LSPDSet lspd
  rstcm1                      '  C870 compare register reset
'/***********     ﾒﾚｯｸ　C-853ボード初期設定　終了  *************/
OrgExec:
  SeikeiOFF          '成形OFF　待機中
  TrnsReqON         '搬送依頼信号ＯＮ
'/* リセットスイッチ入力待ち */
'    Label2(4).Caption = "異常リセット信号待ち"
'    While SystemReadyChk() = 0
      'FrmEmg.Show
'      DoEvents
'    Wend
'/* サーボモータの原点出し */
  CtlDisp
'  genten
'  Ready_Wait
  OrgFlg = False       '原点出し
  OrgOFF               '----------- 原点LED          2002.10.16 KYOCERA
'/* グラフィック画面の初期化 */
'/* ファンクションコード表示 */
'/* メニューの表示　*/
'/* メニューの選択　*/
  ic = 2: c0 = 0: mc1 = 0: imo = 0

  Do
    If FrmMenuFlg = False Then
      Exit Do        'メニューから抜けるときfalse
    End If
    If OrgFlg = True Then Exit Do             '原点出し
    If SystemReadyChk() = 0 Then Exit Do      'システムレディがoffならシステムレディ待ち
    '
    If ArmChk <> 0 Then               'アラームメッセージ
      frmerr_sign.Show 1
    End If
'/* マニコン入力処理 */
  z = r_z()
  If imo = 3 Then cal_pid gM_sa, gM_p, gM_lim
' FlgAuto = AutoChk()        '自動状態ﾁｪｯｸ？ (<>0 自動)
  FlgAuto = 0                '強制的に自動状態 にする　自動=0
  If FlgAuto = 0 Then          '------- 自動の時SW-BOX2は無効
    ch = 1: mc0 = BitRd(ch) And &HF     'mc0=inp(DIO_P+1);
  Else
    mc0 = 0
  End If
  '
  If (mc0 And &H6) = &H6 And z > pv_ch And imo <> 3 Then
      C870SlowStop    'outp(AX_COM,0xfe); /* 停止 */
      CtlVelo         'outp(DIO_P+3,0x05);/* 速度ﾓｰﾄﾞ */
      imo = 3
      mc1 = mc0
   End If
'
  If mc0 <> mc1 Then
      mc1 = mc0
      Select Case mc0
      Case &H6                        '上方向に動く
        g_sts = GentenCmdChk          '搬送シリンダの原点を確認
        If g_sts = 1 Then
          'di_d2 = di_d2 & &HBF          '/* 原点LED　OFF */
          gOrgFlg = False                '原点復帰完了=TRUE
          OrgOFF    'ch = 1: outp ch, di_d2        'outp(DIO_P+1,di_d2);
          Ready_Wait                    'while((inp(AX_STS)&1)!=0);
          C870Command &H12              'outp(AX_COM,0x12);
          imo = 1
        End If
      Case &H5                         '下方向に動く
        gOrgFlg = False                '原点復帰完了=TRUE
        OrgOFF   'ch = 1: outp ch, di_d2        'outp(DIO_P+1,di_d2);
        Ready_Wait                    'while((inp(AX_STS)&1)!=0);
        C870Command &H13              'outp(AX_COM,0x13);
        imo = 1
      Case &HC
        pv_ch = r_z()
        rozFileSave
      Case Else     'default:
        If imo = 3 Then
          imo = 0
          CtlDisp                   ' /* 位置ﾓｰﾄﾞ */
          ch = 1: DaVoltOut ch, 0   '速度指令電圧０
        End If
        If imo = 1 Then
          imo = 0
          C870SlowStop              ' /* 停止 */
        End If
      End Select
    End If
'/* 時計　圧力　Ｚ値 の表示 */
    ttime = Time$       '_strtime(ttime);

  If Mid(ttime, 7, 1) <> stime Then

      '/* 速度をゼロ */
    ch = 1: DaVoltOut ch, 0
  '/* １秒に１回時計表示 */
    If Int(nTime) <> Int(Timer) Then
      nTime = Timer
      Label2(3).Caption = ttime   'disp_t(ttime);
      'txtcolor(3);

  '/* Ｚ位置表示 */
      Label2(0).Caption = Format(z, "0.000")
  '/* 圧力表示 */
      apre = r_pres()   '/* 圧力読み取り */
      Label2(1).Caption = Format(apre, "0.000")
    End If
  '
  'ショット数Ｔ
    Label2(6).Caption = Format(InitDat(11), "0")
  '
  If gOrgStartFlg = False Then  '2002.10.18 KYOCERA
    If gOrgFlg = True Then '原点復帰完了=TRUE
      Label2(5).Caption = "原点"
    Else
      Label2(5).Caption = ""
    End If
  End If
    '-------------- ピラニ計読み
'    LS21S_Monitor    '2006.12.21 削除 s.f
  End If
  '-------------- ピラニ計読み
  '    LS21S_Monitor
  '/* エラー表示 */
  '------------------ BITS を読む
  '2002.01.15削除→ArmChkとEmgChkに変更
'/* キーボード入力 */
     DoEvents
  Loop
  '
  'TrnsReqOFF    '搬送依頼信号ＯＦＦ
  
  If MemoFlg = True Then             'FKeyメモ帳の処理
    MemoFlg = False
    FrmMenuFlg = True
    ExecMemo gcoxFldir, gcoxFlName
    GoTo OrgExec:
  End If

  If OrgFlg = True Then              '原点出し
    genten
    GoTo OrgExec:
  End If
  If SystemReadyChk() = 0 Then       'システムレディがoffならシステムレディ待ち
    RecEmgDtSave sdt3, sdt1, sdt2    '非常停止メッセージの保存  2004.3.8
    FrmMenuFlg = False
    Unload Me
    ReadyFrm.Show
  End If
  If ArmChk <> 0 Then               'アラームメッセージ
    frmerr_sign.Show 1
  End If
  '---------------------------- 画面が変わると時の処理
  If FrmMenuFlg = False Then            'メニューから抜けるときfalse
    FrmMenuFlg = True                   'メニューから抜けるときfalse
    Select Case NextView
    Case 1  '成形（オンリー）
      Unload Me
      LS21_TC.Show
    Case 2  '連続成形画面  成形（シングル）
      Unload Me
      LS21_SC.Show
    Case 3  'データ出力
      Unload Me
      LS21_ResGph.Show
    Case 4  'I O チェック
      Unload Me
      IOChk.Show
    Case 5  'スケール変更
      Unload Me
      LS21_GphScale.Show
    Case 6  '読み出し
      coxFlLoad
      Label2(2) = gcoxFlName
      cfileSave
      GoTo OrgExec:
    Case 7  'メモ帳
      ExecMemo gcoxFldir, gcoxFlName
      GoTo OrgExec:
    Case 8  'edit
      Unload Me
      MYEdit.Show
    End Select
  End If
End Sub


Private Sub Timer1_Timer()
  Timer1.Enabled = False
  qd62_Main
End Sub

Private Sub PltPrns(n%)
Dim i%, sts%, stsEmg%
'--------- パレット循環
  Timer1.Enabled = False
  i = n
  'Text1(0).Text = Format(TrnsMax - (n - i), "0")
  For i = 1 To n
    '
    PCTrnsReq     ' パレット1順指令
    Text1(0).Text = Format(i, "0")
    WaitSec 1
    sts = 0
    Do
      sts = PCTrnsChk()   'PCから搬送中=1
      stsEmg = SystemReadyChk()  '非常停止
      If sts = 0 Or stsEmg = 0 Or lTrnsFLg = False Then Exit Do
      DoEvents
    Loop
    '
    If stsEmg = 0 Or lTrnsFLg = False Then Exit For
  
  Next i
  Text1(0).Text = Format(n, "0")
  lTrnsFLg = False                  '搬送中フラグ
  Command2(1).Caption = "搬送開始"
  Timer1.Enabled = True
End Sub
'2002.10.17 KYOCERA
Private Sub Timer2_Timer()
  
  If gOrgStartFlg = True Then
    If r_z > 0.1 Then
      OrgOFF
      Label2(5).Caption = ""
      Command1(0).Enabled = False
      Command1(1).Enabled = False
      Command1(2).Enabled = False
      Command1(4).Enabled = False
    Else
      OrgON
      Label2(5).Caption = "原点"
      Command1(0).Enabled = True
      Command1(1).Enabled = True
      Command1(2).Enabled = True
      Command1(4).Enabled = True
    End If
  End If
      
End Sub

Private Sub DispText1(dt!, flg%)   '  回数
  If flg = False Then
    VScroll1.Visible = False
    Text1(0).Visible = False
  Else
    VScroll1.Visible = True
    VScroll1.Value = dt * lK1
    Text1(0).Visible = True
    Text1(0).Text = Format(dt, "###")
  End If
End Sub
Private Sub SetVScroll1()               ' VSScrollの量ｓｅｔ
    lK1 = 1
    VScroll1.min = 50 * lK1
    VScroll1.max = 0 * lK1
    VScroll1.LargeChange = 1 * lK1
    VScroll1.SmallChange = 1 * lK1
End Sub
Private Sub VScroll1_Change()
Dim dt!
  dt = VScroll1.Value / lK1
  DispText1 dt, True       '回数
End Sub

