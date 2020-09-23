VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Magic Quest RPG Game"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox nobj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   7680
      Picture         =   "frmGame.frx":08CA
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   96
      Top             =   2280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nobj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   1
      Left            =   8160
      Picture         =   "frmGame.frx":0CFC
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   95
      Top             =   3960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox nobj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   2
      Left            =   5640
      Picture         =   "frmGame.frx":2B9E
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   94
      Top             =   3240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox nobj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   3
      Left            =   6840
      Picture         =   "frmGame.frx":4A40
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   93
      Top             =   3240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox nobj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   6720
      Picture         =   "frmGame.frx":68E2
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   92
      Top             =   4440
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   16
      Left            =   6120
      Picture         =   "frmGame.frx":6D14
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   91
      Top             =   5280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   15
      Left            =   5760
      Picture         =   "frmGame.frx":7146
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   90
      Top             =   6000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   14
      Left            =   5400
      Picture         =   "frmGame.frx":7578
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   89
      Top             =   6000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   13
      Left            =   5760
      Picture         =   "frmGame.frx":79AA
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   88
      Top             =   5640
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   12
      Left            =   5400
      Picture         =   "frmGame.frx":7DDC
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   87
      Top             =   5640
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   11
      Left            =   5760
      Picture         =   "frmGame.frx":820E
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   86
      Top             =   5280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   10
      Left            =   5400
      Picture         =   "frmGame.frx":8640
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   85
      Top             =   5280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   9
      Left            =   5040
      Picture         =   "frmGame.frx":8A72
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   84
      Top             =   6000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   8
      Left            =   5040
      Picture         =   "frmGame.frx":8EA4
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   83
      Top             =   5640
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   7
      Left            =   5040
      Picture         =   "frmGame.frx":92D6
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   82
      Top             =   5280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   6
      Left            =   4320
      Picture         =   "frmGame.frx":9708
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   81
      Top             =   6000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   5
      Left            =   6120
      Picture         =   "frmGame.frx":9B3A
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   80
      Top             =   5640
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   4680
      Picture         =   "frmGame.frx":9F6C
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   79
      Top             =   6000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   3
      Left            =   4680
      Picture         =   "frmGame.frx":A39E
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   78
      Top             =   5640
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   4440
      Picture         =   "frmGame.frx":A7D0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   77
      Top             =   5280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   4320
      Picture         =   "frmGame.frx":AFAA
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   76
      Top             =   5640
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox nTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   4680
      Picture         =   "frmGame.frx":B3DC
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   75
      Top             =   6360
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox obj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   6000
      Picture         =   "frmGame.frx":B80E
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   74
      Top             =   4440
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox objm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   6600
      Picture         =   "frmGame.frx":BC40
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   73
      Top             =   4440
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   16
      Left            =   3240
      Picture         =   "frmGame.frx":C072
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   72
      Top             =   4320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   15
      Left            =   2880
      Picture         =   "frmGame.frx":C4A4
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   64
      Top             =   5040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox bsLouis 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   3480
      Picture         =   "frmGame.frx":C8D6
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   48
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrFPS 
      Interval        =   1000
      Left            =   840
      Top             =   3240
   End
   Begin VB.PictureBox objm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   3
      Left            =   6600
      Picture         =   "frmGame.frx":D660
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   46
      Top             =   3240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox obj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   3
      Left            =   6000
      Picture         =   "frmGame.frx":F502
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   45
      Top             =   3240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox obj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   2
      Left            =   4800
      Picture         =   "frmGame.frx":113A4
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   42
      Top             =   3240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox objm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   2
      Left            =   5520
      Picture         =   "frmGame.frx":13246
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   41
      Top             =   3240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   14
      Left            =   2520
      Picture         =   "frmGame.frx":150E8
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   39
      Top             =   5040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   13
      Left            =   2880
      Picture         =   "frmGame.frx":1551A
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   35
      Top             =   4680
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   12
      Left            =   2520
      Picture         =   "frmGame.frx":1594C
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   34
      Top             =   4680
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   11
      Left            =   2880
      Picture         =   "frmGame.frx":15D7E
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   10
      Left            =   2520
      Picture         =   "frmGame.frx":161B0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   32
      Top             =   4320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   9
      Left            =   2160
      Picture         =   "frmGame.frx":165E2
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   28
      Top             =   5040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   8
      Left            =   2160
      Picture         =   "frmGame.frx":16A14
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   27
      Top             =   4680
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   7
      Left            =   2160
      Picture         =   "frmGame.frx":16E46
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   26
      Top             =   4320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox buffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   7440
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox objm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   1
      Left            =   7920
      Picture         =   "frmGame.frx":17278
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   23
      Top             =   3960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox objm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   8040
      Picture         =   "frmGame.frx":1911A
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox obj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   1
      Left            =   7560
      Picture         =   "frmGame.frx":1954C
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   21
      Top             =   3840
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox obj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   7800
      Picture         =   "frmGame.frx":1B3EE
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   6
      Left            =   1440
      Picture         =   "frmGame.frx":1B820
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox enemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   120
      Picture         =   "frmGame.frx":1BC52
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox enemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   0
      Left            =   120
      Picture         =   "frmGame.frx":1C8F4
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox charm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   4200
      Picture         =   "frmGame.frx":1D4D6
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox chars 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   4320
      Picture         =   "frmGame.frx":1F018
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   5
      Left            =   3240
      Picture         =   "frmGame.frx":20B5A
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   1800
      Picture         =   "frmGame.frx":20F8C
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   3
      Left            =   1800
      Picture         =   "frmGame.frx":213BE
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   1560
      Picture         =   "frmGame.frx":217F0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   1440
      Picture         =   "frmGame.frx":21FCA
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   1800
      Picture         =   "frmGame.frx":223FC
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer tmrM 
      Interval        =   100
      Left            =   120
      Top             =   3000
   End
   Begin VB.PictureBox stats 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   4560
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   25
      Top             =   0
      Width           =   2175
      Begin VB.Timer tmrTime 
         Interval        =   5000
         Left            =   240
         Top             =   2160
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "Menu"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   2520
         Width           =   1575
      End
   End
   Begin VB.PictureBox main 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   6735
      TabIndex        =   49
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   255
         Left            =   600
         TabIndex        =   63
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Game"
         Height          =   255
         Left            =   600
         TabIndex        =   60
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton cmdRoll 
         Caption         =   "Roll Stats"
         Height          =   255
         Left            =   3600
         TabIndex        =   58
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtDefense 
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "3"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtAttack 
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "5"
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "New Game"
         Height          =   255
         Left            =   600
         TabIndex        =   51
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit Game"
         Height          =   255
         Left            =   600
         TabIndex        =   50
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Defense:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   55
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Magic Quest"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   840
         TabIndex        =   54
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "KevCom(R) 2002"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1680
         TabIndex        =   53
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Attack:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   52
         Top             =   1680
         Width           =   735
      End
   End
   Begin VB.PictureBox market 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   4575
      TabIndex        =   5
      Top             =   -120
      Width           =   4575
      Begin VB.CommandButton cmdExitMarket 
         Caption         =   "Exit Market"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Timer tmrMoney 
         Left            =   1800
         Top             =   1200
      End
      Begin VB.ListBox lstMCost 
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ListBox lstICost 
         Height          =   255
         Left            =   1320
         TabIndex        =   30
         Top             =   1920
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdSell 
         Caption         =   "Sell"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ListBox lstStuff 
         Height          =   1230
         Left            =   1920
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuy 
         Caption         =   "Buy"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ListBox lstBuy 
         Height          =   1425
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCost 
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   44
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cost:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   43
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblMoney 
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   36
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Money:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "What do you want to buy?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1440
         TabIndex        =   40
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.PictureBox menu 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdExitG 
         Caption         =   "Exit Game"
         Height          =   255
         Left            =   1680
         TabIndex        =   62
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton cmdLoad2 
         Caption         =   "Load Game"
         Height          =   255
         Left            =   1680
         TabIndex        =   61
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Game"
         Height          =   255
         Left            =   1680
         TabIndex        =   59
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdDrop 
         Caption         =   "Drop Item"
         Height          =   255
         Left            =   480
         TabIndex        =   47
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdUseItem 
         Caption         =   "Use Item"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ListBox lstItems 
         Height          =   1815
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.PictureBox battle 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   4575
      TabIndex        =   66
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdAttack 
         Caption         =   "Attack!"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   2400
         Width           =   975
      End
      Begin VB.Timer tmrAttack 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   1560
         Top             =   1680
      End
      Begin VB.Timer tmrAttack2 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   2040
         Top             =   1680
      End
      Begin VB.Timer tmrAIAttack 
         Interval        =   1000
         Left            =   1080
         Top             =   1680
      End
      Begin VB.Timer tmrMeters 
         Interval        =   10
         Left            =   600
         Top             =   1680
      End
      Begin VB.PictureBox meter1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   360
         ScaleHeight     =   5
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   70
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox meter2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   2640
         ScaleHeight     =   5
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   69
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run!"
         Height          =   255
         Left            =   2160
         TabIndex        =   68
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdMagic 
         Caption         =   "Magic!"
         Height          =   255
         Left            =   1200
         TabIndex        =   67
         Top             =   2400
         Width           =   855
      End
      Begin VB.Timer tmrMagic 
         Interval        =   100
         Left            =   2520
         Top             =   1680
      End
      Begin VB.Image imgPlayer 
         Height          =   495
         Left            =   360
         Picture         =   "frmGame.frx":2282E
         Top             =   840
         Width           =   450
      End
      Begin VB.Image imgEnemy 
         Height          =   465
         Left            =   2640
         Picture         =   "frmGame.frx":2344C
         Top             =   840
         Width           =   465
      End
      Begin VB.Image imgAttack 
         Height          =   150
         Left            =   720
         Picture         =   "frmGame.frx":2402E
         Top             =   1080
         Width           =   900
      End
      Begin VB.Image imgAttack2 
         Height          =   150
         Left            =   1680
         Picture         =   "frmGame.frx":24778
         Top             =   1080
         Width           =   900
      End
      Begin VB.Image imgMagic 
         Height          =   480
         Left            =   2640
         Picture         =   "frmGame.frx":24EC2
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.Label lblM 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Welcome to Magic Quest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   1320
      TabIndex        =   37
      Top             =   3960
      Width           =   135
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAttack_Click()
imgAttack.Visible = True
tmrAttack.Enabled = True
End Sub

Private Sub cmdBuy_Click()
If lstBuy.ListIndex < 0 Then Exit Sub

Dim Cost, I

Cost = lstMCost.List(lstBuy.ListIndex)

If Cost < p.Pack.money Then

    p.Pack.money = p.Pack.money - Cost
    
    If AddPackItem(lstBuy.List(lstBuy.ListIndex), lstMCost.List(lstBuy.ListIndex)) = False Then
    p.Pack.money = p.Pack.money + Cost
    End If

    lstStuff.Clear

    For I = 1 To 20
        If p.Pack.Item(I).Act = True Then
            lstStuff.AddItem p.Pack.Item(I).Caption
        End If
    Next I

    frmGame.lblMoney.Caption = p.Pack.money & "$"

Else

    MsgBox ("You dont have enough money to buy item!"), , "MagicQuest"

End If
End Sub

Private Sub cmdDrop_Click()
If lstItems.ListIndex < 0 Then Exit Sub

If MsgBox("Are you sure you want to drop the item " & lstItems.List(lstItems.ListIndex) & "?", vbYesNo, "Drop Item") = vbNo Then Exit Sub

p.Pack.Item(lstItems.ListIndex + 1).Act = False
lstItems.RemoveItem lstItems.ListIndex
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdExitG_Click()
If MsgBox("Do you want to save your game before you exit?", vbYesNo, "Exit") = vbYes Then cmdSave_Click
End
End Sub

Private Sub cmdExitMarket_Click()
        main.Visible = False
        board.Visible = True
        battle.Visible = False
        menu.Visible = False
        market.Visible = False
End Sub

Private Sub cmdLoad_Click()
Dim I As Long, n1, n2, n3, n4, n5

stats.Visible = True
lblM.Caption = "Loading..."
Msg.Time = 9999999

Open App.path + "\save.dat" For Input As #1

Input #1, n1, n2, n3, n4, n5

If n5 = "NO SAVE FILE" Then
Close #1
MsgBox ("There is no saved game file..."), vbCritical, "Game Save Error"
Exit Sub
End If

main.Visible = n1
menu.Visible = n2
board.Visible = n3
market.Visible = n4
battle.Visible = n5

Input #1, p.X, p.Y, p.DestX, p.DestY, p.dir, p.Walking
Input #1, p.Pack.money, p.Pack.mDust, p.SpellCut, p.SpellFreeze
Input #1, p.B.Attack, p.B.Defense, p.B.Exp, p.B.Health, p.B.Level, p.B.MaxHealth, p.B.TExp

For I = 1 To 20
Input #1, p.Pack.Item(I).Act, p.Pack.Item(I).Caption, p.Pack.Item(I).Value
DoEvents
Next I

For I = 1 To 20
Input #1, o(I).Act, o(I).AI_Type, o(I).DestX, o(I).DestY, o(I).dir, o(I).Name, o(I).Person, o(I).plID, n1, o(I).SpeekI, o(I).Step, o(I).StepC, o(I).T, o(I).Walking, o(I).X, o(I).Y
o(I).Speek() = Split(n1, "|")
DoEvents
Next I

Input #1, Louis, Enem.Attack, Enem.Defense, Enem.HP, Enem.MaxHP, Enem.Name

Input #1, M.Mapheight, M.Mapwidth, M.MarketPath, M.MaxAttack, M.MaxDefense, M.MaxHP, M.Midi, M.Name
buffer.Width = M.Mapwidth * TileWidth
buffer.Height = M.Mapheight * TileHeight

For I = 0 To M.Mapheight * M.Mapwidth
Input #1, T(I).Ani, T(I).DoorDir, T(I).DoorPath, T(I).DoorWay, T(I).isMarket, T(I).plID, T(I).SignText, T(I).Type, T(I).Walkable, T(I).X, T(I).Y
DoEvents
Next I

Close #1

lblM.Caption = "Loaded!"
Msg.Time = 20
cmdMenu_Click
cmdMenu_Click
End Sub

Private Sub cmdLoad2_Click()
cmdLoad_Click
End Sub

Private Sub cmdMagic_Click()
tmrMagic.Enabled = True
imgMagic.Visible = True
End Sub

Private Sub cmdMenu_Click()
Select Case cmdMenu.Caption
    Case "Menu"
        LoadMenu
        cmdMenu.Caption = "Back to Game"
    Case "Back to Game"
        main.Visible = False
        board.Visible = True
        battle.Visible = False
        menu.Visible = False
        market.Visible = False
        cmdMenu.Caption = "Menu"
End Select
End Sub

Private Sub cmdMenu_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then KeyCode = 0
End Sub

Private Sub cmdOptions_Click()
Load frmOptions
frmOptions.Visible = True
End Sub

Private Sub cmdRoll_Click()
txtAttack.text = Int(Rnd * 10) + 1
txtDefense.text = 10 - Val(txtAttack.text)
End Sub

Private Sub cmdRun_Click()
If Enem.MaxHP > Int(Rnd * 5) + 10 Then Exit Sub

main.Visible = False
board.Visible = True
battle.Visible = False
menu.Visible = False
market.Visible = False
End Sub

Private Sub cmdSave_Click()
Dim I As Long

lblM.Caption = "Saving..."
MessageTime = 9999999

Open App.path + "\save.dat" For Output As #1

Write #1, main.Visible, menu.Visible, board.Visible, market.Visible, battle.Visible

Write #1, p.X, p.Y, p.DestX, p.DestY, p.dir, p.Walking
Write #1, p.Pack.money, p.Pack.mDust, p.SpellCut, p.SpellFreeze
Write #1, p.B.Attack, p.B.Defense, p.B.Exp, p.B.Health, p.B.Level, p.B.MaxHealth, p.B.TExp

For I = 1 To 20
Write #1, p.Pack.Item(I).Act, p.Pack.Item(I).Caption, p.Pack.Item(I).Value
Next I

For I = 1 To 20
Write #1, o(I).Act, o(I).AI_Type, o(I).DestX, o(I).DestY, o(I).dir, o(I).Name, o(I).Person, o(I).plID, Join(o(I).Speek, "|"), o(I).SpeekI, o(I).Step, o(I).StepC, o(I).T, o(I).Walking, o(I).X, o(I).Y
DoEvents
Next I

Write #1, Louis, Enem.Attack, Enem.Defense, Enem.HP, Enem.MaxHP, Enem.Name

Write #1, M.Mapheight, M.Mapwidth, M.MarketPath, M.MaxAttack, M.MaxDefense, M.MaxHP, M.Midi, M.Name
buffer.Width = M.Mapwidth * TileWidth
buffer.Height = M.Mapheight * TileHeight

For I = 0 To M.Mapheight * M.Mapwidth
Write #1, T(I).Ani, T(I).DoorDir, T(I).DoorPath, T(I).DoorWay, T(I).isMarket, T(I).plID, T(I).SignText, T(I).Type, T(I).Walkable, T(I).X, T(I).Y
DoEvents
Next I

Close #1

lblM.Caption = "Saved!"
MessageTime = 20
End Sub

Private Sub cmdSell_Click()
If lstStuff.ListIndex < 0 Then Exit Sub

Dim I As Long

I = lstStuff.ListIndex

If MsgBox("Are you sure want to sell " & lstStuff.List(I) & " for " & lstICost.List(I), vbYesNo, "Sell") = vbNo Then Exit Sub

p.Pack.money = p.Pack.money + Val(lstICost.List(I))
p.Pack.Item(I + 1).Act = False

lstStuff.Clear

For I = 1 To 20
    If p.Pack.Item(I).Act = True Then
        lstStuff.AddItem p.Pack.Item(I).Caption
    End If
Next I

frmGame.lblMoney.Caption = p.Pack.money & "$"
End Sub

Private Sub cmdStart_Click()
NewGame
main.Visible = False
board.Visible = True
battle.Visible = False
menu.Visible = False
market.Visible = False
stats.Visible = True
End Sub

Private Sub cmdUseItem_Click()
If lstItems.ListIndex < 0 Then Exit Sub
UseItem UCase(lstItems.List(lstItems.ListIndex)), lstItems.ListIndex
lstItems.RemoveItem lstItems.ListIndex
End Sub

Function BufZero(num) As Long
BufZero = num
If num < 0 Then BufZero = 0
End Function

Private Sub Form_Load()
Dim C, I As Long, per As Long, place, text

tmrTime_Timer

gSpeed = Val(GetSetting("MAGICRPG", "GEN", "SPEED", 1100))

Randomize

TFPS = 30

stats.Visible = False
main.Visible = True
board.Visible = False
battle.Visible = False
menu.Visible = False
market.Visible = False

Me.Visible = True

Do

    If TFPS And GetTickCount > 300 Then
        C = 0
        FPS = FPS + 1
        
        If board.Visible = True Then
            
            p.SpeachI = p.SpeachI - 1
            
            buffer.Cls

            MovePlayer
            MoveObjs

            For I = 0 To M.Mapheight * M.Mapwidth

                If isValue(T(I).Type, 2) = True Then
    
                    If T(I).AniC > 10 Then
                        T(I).AniC = 0
                        T(I).Ani = T(I).Ani + 1: If T(I).Ani > 1 Then T(I).Ani = 0
                    Else
                        T(I).AniC = T(I).AniC + 1
                    End If
    
                End If
            
            Select Case Night
            Case False
            E.DrawObj buffer.hdc, T(I).X, T(I).Y, TileWidth, TileHeight, Tile(T(I).Type).hdc, T(I).Ani * 18, 0, vbSrcCopy
            Case True
            E.DrawObj buffer.hdc, T(I).X, T(I).Y, TileWidth, TileHeight, nTile(T(I).Type).hdc, T(I).Ani * 18, 0, vbSrcCopy
            End Select
            
    Next I

    For I = 1 To 20
       If o(I).Act = True Then
            E.DrawObj buffer.hdc, o(I).X, o(I).Y, TileWidth, TileHeight, objm(o(I).T).hdc, o(I).Step * 18, o(I).dir * 18, Mask
            
            Select Case Night
            Case False
            E.DrawObj buffer.hdc, o(I).X, o(I).Y, TileWidth, TileHeight, obj(o(I).T).hdc, o(I).Step * 18, o(I).dir * 18, Sprite
            Case True
            E.DrawObj buffer.hdc, o(I).X, o(I).Y, TileWidth, TileHeight, nobj(o(I).T).hdc, o(I).Step * 18, o(I).dir * 18, Sprite
            End Select
            
        End If
    Next I

    E.DrawObj buffer.hdc, p.X, p.Y, 18, 18, charm.hdc, p.Step * 18, p.dir * 18, Mask
    E.DrawObj buffer.hdc, p.X, p.Y, 18, 18, chars.hdc, p.Step * 18, p.dir * 18, Sprite

    board.Cls
    E.DrawObj board.hdc, 0, 0, board.ScaleWidth, board.ScaleHeight, buffer.hdc, (p.X + 9) - (board.ScaleWidth \ 2), (p.Y + 9) - (board.ScaleHeight \ 2), vbSrcCopy
End If

If main.Visible = False Then

    stats.Cls

    For place = 1 To 10

        Select Case place
            Case 1
                per = p.B.Health \ p.B.MaxHealth
                per = per * 100
    
                 text = "Health: " & p.B.Health & " Out of " & p.B.MaxHealth
    
            Case 2: text = "Money: " & p.Pack.money & "$"
             Case 3: text = "Magic: " & p.Pack.mDust
            Case 4: text = ""
            Case 5: text = "Stats"
            Case 6: text = "Attack: " & p.B.Attack
            Case 7: text = "Defense: " & p.B.Defense
            Case 8: text = "Exp: " & p.B.Exp & " Out of " & p.B.TExp
            Case 9: text = "Level: " & p.B.Level
            Case 10: text = ""
            
        End Select
    
        stats.CurrentX = stats.ScaleWidth \ 2 - stats.TextWidth(text) \ 2
        stats.CurrentY = stats.TextHeight("|") * place + 10
        stats.Print text
        
    Next place
        
        If Msg.Time > 0 Then
        board.ForeColor = vbWhite
        board.CurrentX = board.ScaleWidth \ 2 - board.TextWidth(Msg.message) \ 2
        board.CurrentY = board.ScaleHeight - board.TextHeight("|") - 3
        board.Print Msg.message
        End If
        
End If

SkipLoop:

Else

    C = C + 1
    
End If

DoEvents

Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstBuy_Click()
lblCost.Caption = Val(lstMCost.List(lstBuy.ListIndex)) & "$"
End Sub

Private Sub tmrAIAttack_Timer()
If battle.Visible = False Then Exit Sub
imgAttack2.Visible = True
tmrAttack2.Enabled = True
End Sub

Private Sub tmrAttack_Timer()
If battle.Visible = False Then Exit Sub
imgAttack.Visible = False
tmrAttack.Enabled = False

Enem.HP = Enem.HP - IIf((p.B.Attack - Enem.Defense) < 0, 0, (p.B.Attack - Enem.Defense))

If Enem.HP <= 0 Then

        If Enem.Name = "Louis" Then
        AddPackItem "Special Candy", 2000
        Louis = True
        p.Pack.money = p.Pack.money + 500
        o(1).DestX = o(1).X - 18: o(1).dir = nDir.Right
        o(1).X = o(1).DestX
        o(1).Speek() = Split("We'll youve beaten me|Bet if we had a rematch I'd win!", "|")
        End If
        
    p.B.Exp = p.B.Exp + (Enem.Attack + Enem.Defense)

    main.Visible = False
    board.Visible = True
    battle.Visible = False
    menu.Visible = False
    market.Visible = False

    If Int(Rnd * 5) = Int(Rnd * 4) Then
    
        Dim money
        money = Int(Rnd * 20) + 10
        p.Pack.money = p.Pack.money + money
            
                    lblM.Caption = money & "$ was found on the ground..."
                    MessageTime = 10 + Len(lblM.Caption)
    End If

    If p.B.Exp >= p.B.TExp Then
        p.B.Exp = 0
        p.B.Level = p.B.Level + 1
        p.B.Attack = p.B.Attack + 2
        p.B.Defense = p.B.Defense + 1
        p.B.Health = p.B.Health + 5
        p.B.MaxHealth = p.B.MaxHealth + 5
        p.B.TExp = p.B.TExp + (Int(Rnd * 10) + 1)
    End If
End If
End Sub

Private Sub tmrAttack2_Timer()
If battle.Visible = False Then Exit Sub
imgAttack2.Visible = False
tmrAttack2.Enabled = False

If Int(Rnd * 3) = 2 Then Exit Sub

p.B.Health = p.B.Health - IIf((Enem.Attack - p.B.Defense) < 0, 0, (Enem.Attack - p.B.Defense))

If p.B.Health <= 0 Then
    p.B.Health = 50
        
    main.Visible = False
    board.Visible = True
    battle.Visible = False
    menu.Visible = False
    market.Visible = False

    LoadMap "house.dat", 1, 2
End If
End Sub

Private Sub tmrFPS_Timer()
Me.Caption = "Magic Quest RPG Game - " & FPS & " FPS"
TFPS = FPS
FPS = 0
End Sub

Private Sub tmrM_Timer()
Msg.Time = Msg.Time - 1
End Sub

Private Sub tmrMagic_Timer()
tmrMagic.Enabled = False
imgMagic.Visible = False

If battle.Visible = True And p.Pack.mDust > 0 Then

    p.Pack.mDust = p.Pack.mDust - 2: If p.Pack.mDust < 0 Then p.Pack.mDust = 0

    Enem.HP = Enem.HP - IIf(((p.B.Attack * 2) - Enem.Defense) < 0, 0, (p.B.Attack - Enem.Defense))

    If Int(Rnd * 5) = Int(Rnd * 4) Then
        Dim money
        money = Int(Rnd * 20) + 10
        DoMsg money & "$ was found on the ground..."
    End If

    If Enem.HP <= 0 Then
        p.B.Exp = p.B.Exp + (Enem.Attack + Enem.Defense) + (Enem.MaxHP / 3)
        main.Visible = False
        board.Visible = True
        battle.Visible = False
        menu.Visible = False
        market.Visible = False
        
        If Enem.Name = "Louis" Then
        AddPackItem "Special Candy", 2000
        Louis = True
        p.Pack.money = p.Pack.money + 500
        o(1).DestX = o(1).X - 18: o(1).dir = nDir.Right
        o(1).X = o(1).DestX
        o(1).Speek() = Split("We'll youve beaten me|Bet if we had a rematch I'd win!|You should see my brother in Noko City, he is stronger than me!", "|")
        End If
        
        If p.B.Exp >= p.B.TExp Then
            p.B.Exp = 0
            p.B.Level = p.B.Level + 1
            p.B.Attack = p.B.Attack + 2
            p.B.Defense = p.B.Defense + 1
            p.B.Health = p.B.Health + 5
            p.B.MaxHealth = p.B.MaxHealth + 5
            p.B.TExp = p.B.TExp + (Int(Rnd * 10) + 1)
        End If
    End If
End If
End Sub

Private Sub tmrMeters_Timer()
If battle.Visible = False Then Exit Sub

Dim per

meter1.Cls
meter2.Cls

per = 100 \ 75
meter1.Line (0, 0)-(meter1.ScaleWidth * per, meter1.ScaleHeight), vbGreen, BF
per = Enem.HP \ Enem.MaxHP
meter2.Line (0, 0)-(meter2.ScaleWidth * per, meter2.ScaleHeight), vbGreen, BF
End Sub

Private Sub tmrMoney_Timer()
lblMoney.Caption = p.Pack.money
End Sub

Private Sub tmrTime_Timer()
Night = isNight
End Sub
