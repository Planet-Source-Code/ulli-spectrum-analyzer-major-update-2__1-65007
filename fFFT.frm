VERSION 5.00
Begin VB.Form fFFT 
   BackColor       =   &H00406070&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Spectrum Analyzer"
   ClientHeight    =   11025
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11925
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "fFFT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11025
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   890
      Left            =   11415
      Top             =   4365
   End
   Begin VB.Frame fr 
      BackColor       =   &H00406070&
      Caption         =   "Waterfall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   4245
      Index           =   5
      Left            =   90
      TabIndex        =   23
      Top             =   4545
      Width           =   11730
      Begin VB.PictureBox picWaterfall 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'Kein
         DrawStyle       =   2  'Punkt
         FontTransparent =   0   'False
         Height          =   3660
         Left            =   255
         MousePointer    =   2  'Kreuz
         ScaleHeight     =   244
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   749
         TabIndex        =   0
         Top             =   375
         Width           =   11235
      End
      Begin VB.Shape shp 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   2
         Height          =   3705
         Index           =   1
         Left            =   240
         Top             =   360
         Width           =   11280
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00406070&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2025
      Index           =   4
      Left            =   90
      TabIndex        =   21
      Top             =   8880
      Width           =   11715
      Begin VB.CommandButton btMagn 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Magnifier"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10755
         Style           =   1  'Grafisch
         TabIndex        =   17
         Top             =   885
         Width           =   855
      End
      Begin VB.CommandButton btStop 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10770
         Style           =   1  'Grafisch
         TabIndex        =   18
         Top             =   1410
         Width           =   855
      End
      Begin VB.CommandButton btStart 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10755
         Style           =   1  'Grafisch
         TabIndex        =   16
         Top             =   375
         Width           =   855
      End
      Begin VB.Frame fr 
         BackColor       =   &H00406070&
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   1530
         Index           =   6
         Left            =   4710
         TabIndex        =   31
         Top             =   270
         Width           =   2100
         Begin VB.CheckBox ckTimeWindow 
            BackColor       =   &H00406070&
            Caption         =   "Time Window"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Left            =   285
            TabIndex        =   35
            Top             =   1230
            Width           =   1290
         End
         Begin VB.OptionButton optWeighted 
            BackColor       =   &H00406070&
            Caption         =   "Weighted Average"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   240
            Left            =   270
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   915
            Width           =   1650
         End
         Begin VB.OptionButton optSavGol 
            BackColor       =   &H00406070&
            Caption         =   "Savitzki-Golay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   240
            Left            =   270
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   615
            Width           =   1320
         End
         Begin VB.OptionButton optFilterNone 
            BackColor       =   &H00406070&
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   240
            Left            =   270
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   315
            Width           =   720
         End
      End
      Begin VB.Frame fr 
         BackColor       =   &H00406070&
         Caption         =   "Spectrum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1530
         Index           =   2
         Left            =   3180
         TabIndex        =   28
         Top             =   270
         Width           =   1410
         Begin VB.VScrollBar scrShift 
            Height          =   990
            LargeChange     =   20
            Left            =   945
            Max             =   -50
            Min             =   50
            SmallChange     =   5
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
         End
         Begin VB.VScrollBar scrGain 
            Height          =   990
            LargeChange     =   2
            Left            =   255
            Max             =   1
            Min             =   10
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   270
            Value           =   1
            Width           =   240
         End
         Begin VB.Label lb 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Shift"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   3
            Left            =   900
            TabIndex        =   30
            Top             =   1290
            Width           =   330
         End
         Begin VB.Label lb 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Gain"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   29
            Top             =   1290
            Width           =   330
         End
      End
      Begin VB.Frame fr 
         BackColor       =   &H00406070&
         Caption         =   "Waterfall"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   1530
         Index           =   1
         Left            =   6930
         TabIndex        =   25
         Top             =   270
         Width           =   3675
         Begin VB.VScrollBar scrSpeed 
            Height          =   990
            LargeChange     =   40
            Left            =   1995
            Max             =   200
            SmallChange     =   5
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   270
            Value           =   100
            Width           =   240
         End
         Begin VB.CheckBox ckHum 
            BackColor       =   &H00406070&
            Caption         =   "Hum Sup"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Left            =   2505
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1035
            Width           =   975
         End
         Begin VB.VScrollBar scrHue 
            Height          =   990
            LargeChange     =   45
            Left            =   255
            Max             =   0
            Min             =   212
            SmallChange     =   9
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
         End
         Begin VB.CheckBox ckFreeze 
            BackColor       =   &H00406070&
            Caption         =   "Freeze"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Left            =   2490
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   645
            Width           =   795
         End
         Begin VB.CheckBox ckTicks 
            BackColor       =   &H00406070&
            Caption         =   "Time Ticks"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Left            =   2490
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   270
            Width           =   1110
         End
         Begin VB.VScrollBar scrContrast 
            Height          =   990
            Left            =   1410
            Max             =   1
            Min             =   14
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   270
            Value           =   7
            Width           =   240
         End
         Begin VB.VScrollBar scrBright 
            Height          =   990
            LargeChange     =   48
            Left            =   825
            Max             =   0
            Min             =   200
            SmallChange     =   12
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
         End
         Begin VB.Label lb 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Speed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   5
            Left            =   1845
            TabIndex        =   34
            Top             =   1290
            Width           =   495
         End
         Begin VB.Label lb 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   4
            Left            =   195
            TabIndex        =   33
            Top             =   1290
            Width           =   360
         End
         Begin VB.Label lb 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Contr"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   27
            Top             =   1290
            Width           =   405
         End
         Begin VB.Label lb 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Bright"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   0
            Left            =   735
            TabIndex        =   26
            Top             =   1290
            Width           =   405
         End
      End
      Begin VB.Frame fr 
         BackColor       =   &H00406070&
         Caption         =   "Samples"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1530
         Index           =   0
         Left            =   165
         TabIndex        =   22
         Top             =   270
         Width           =   2895
         Begin VB.VScrollBar scrReso 
            Height          =   990
            LargeChange     =   2
            Left            =   2235
            Max             =   0
            Min             =   5
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
         End
         Begin VB.OptionButton optSamples 
            BackColor       =   &H00406070&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   5
            Left            =   1260
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1005
            Width           =   885
         End
         Begin VB.OptionButton optSamples 
            BackColor       =   &H00406070&
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   4
            Left            =   1260
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   675
            Width           =   885
         End
         Begin VB.OptionButton optSamples 
            BackColor       =   &H00406070&
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   0
            Left            =   285
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   330
            Width           =   885
         End
         Begin VB.OptionButton optSamples 
            BackColor       =   &H00406070&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   1
            Left            =   285
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   675
            Width           =   885
         End
         Begin VB.OptionButton optSamples 
            BackColor       =   &H00406070&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   2
            Left            =   285
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1005
            Width           =   885
         End
         Begin VB.OptionButton optSamples 
            BackColor       =   &H00406070&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   3
            Left            =   1260
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   330
            Width           =   885
         End
         Begin VB.Label lbReso 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "x 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Left            =   2535
            TabIndex        =   45
            Top             =   675
            Width           =   210
         End
         Begin VB.Label lbFFT 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "FFT = 12345 bins"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Left            =   300
            TabIndex        =   44
            Top             =   1290
            Width           =   1245
         End
         Begin VB.Label lb 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Resolution"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Index           =   6
            Left            =   1980
            TabIndex        =   43
            Top             =   1290
            Width           =   750
         End
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00406070&
      Caption         =   "Spectrum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4395
      Index           =   3
      Left            =   90
      TabIndex        =   19
      Top             =   60
      Width           =   11730
      Begin VB.Frame fr 
         BorderStyle     =   0  'Kein
         Height          =   150
         Index           =   8
         Left            =   195
         TabIndex        =   37
         Top             =   4125
         Width           =   10620
         Begin VB.HScrollBar scrRange 
            Height          =   135
            Left            =   -240
            TabIndex        =   38
            Top             =   15
            Width           =   11100
         End
      End
      Begin VB.PictureBox picSpectrum 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   240
         MousePointer    =   2  'Kreuz
         ScaleHeight     =   233
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   749
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   300
         Width           =   11235
         Begin VB.Label lbBuffering 
            Alignment       =   2  'Zentriert
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   5550
            TabIndex        =   36
            Top             =   675
            Visible         =   0   'False
            Width           =   105
         End
      End
      Begin VB.Label lbrange 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0E0FF&
         Height          =   165
         Left            =   10905
         TabIndex        =   41
         Top             =   4110
         Width           =   30
      End
      Begin VB.Shape shp 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   2
         Height          =   3540
         Index           =   0
         Left            =   225
         Top             =   285
         Width           =   11280
      End
      Begin VB.Label lbFreq 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   735
         TabIndex        =   20
         Top             =   3855
         Visible         =   0   'False
         Width           =   60
      End
   End
   Begin VB.Menu mnuSelInp 
      Caption         =   "Select Input"
      Begin VB.Menu mnuSelDev 
         Caption         =   "Recording Device..."
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "Run"
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
         Shortcut        =   {F5}
         Tag             =   "{F5}"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
         Shortcut        =   {F8}
         Tag             =   "{F8}"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "fFFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   _____
'== TO DO ========================================================================================================
'   ¯¯¯¯¯
'This is the shortcut character for (System Administration | Sound and Audio Devices | Audio | Recording | Volume)
'and YOU MAY HAVE TO CHANGE that for the appropriate character in your language; in German it happens to be a U
'    ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Const ShortcutU = "U"
'
'=================================================================================================================
'
'24 apr 2006
'added about box
'added range
'added resolution
'fixed bug with restart after stop
'added buffering timer (with low sample rates buffering may take quite long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'19 apr 2006
'added magnifier
'added overlap
'added a time window funktion
'streamlined some code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'15 apr 2006 UMG
'
'changed color conversion from RGB to HLS (hue, luminance, saturation)
'changed tick generation
'added mixer input line selection and menu
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'remove after testing

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private AtStart             As Currency
Private AtEnd               As Currency
Private CPUSpeed            As Currency

Private cFourier            As clsFourier

Private NumSamples          As Long     'number of samples we're gonna use
Private Resolution          As Long
Private NumTimesRes         As Long
Private RangeMin            As Long
Private RangeMax            As Long
Private Hgt                 As Long     'height for bitblt
Private Wid                 As Long     'width for bitblt
Private LastPixelNum        As Long     'last pixel
Private Delay               As Long
Private TickCount           As Long
Private WaitTime            As Long
Private fHs1                As Long     'hum sup
Private fHs2                As Long     'hum sup
Private WtfColor            As Long     '0 red, 42 yellow, 85 green, 122 cyan, 170 blue, 212 magenta

Private Break               As Boolean
Private Finito              As Boolean
Private Freeze              As Boolean
Private WantsTicks          As Boolean
Private TickGate            As Boolean
Private MagnLoaded          As Boolean
Private PWVisible           As Boolean

Private Ovrlap              As Double   'for overlapping
Private HumSupp             As Double

Private tmp                 As Double
Private Max                 As Double

Private Brightness          As Double
Private Saturation          As Double
Private Const TwoFiveFive   As Double = 255
Private FactA               As Double
Private FactB               As Double

Private Gain                As Double
Private Shift               As Double
Private Contrast            As Double
Private FreqMin             As Double
Private FreqMax             As Double

Private SmfWeights          As Variant 'smoothing weights
Private Type SmoothingFilter
    SmfPtr          As Long
    SmfPrevValues() As Double
End Type
Private SmfFilters()        As SmoothingFilter

Private Sub btMagn_Click()

    MagnLoaded = Not MagnLoaded
    If MagnLoaded Then
        Load fMagnifier
      Else 'MAGNLOADED = FALSE/0
        Unload fMagnifier
    End If
    picWaterfall.SetFocus

End Sub

Private Sub btStart_Click()

    SendKeys mnuStart.Tag 'this has the funny side effect that it toggles
    picWaterfall.SetFocus 'the numlock light on the keyboard, but only when in IDE and not always

End Sub

Private Sub btStop_Click()

    SendKeys mnuStop.Tag 'this does the same

End Sub

Private Sub ckFreeze_Click()

    Freeze = (ckFreeze = vbChecked)
    If Freeze Then
        If WantsTicks Then
            With picWaterfall
                .DrawStyle = vbSolid
                picWaterfall.Line (.ScaleWidth, 2)-(0, 2)
                .CurrentY = .CurrentY + 1
                picWaterfall.Print TimeValue(Now)
                .DrawStyle = vbDot
            End With 'picWaterfall
        End If
        TickCount = 0
    End If
    picWaterfall.SetFocus

End Sub

Private Sub ckHum_Click()

    If ckHum = vbChecked Then
        fHs2 = 48.7 * Resolution
        fHs1 = fHs2 / (1 + 0.07 / Resolution)
        fHs2 = fHs2 * (1 + 0.07 / Resolution)
      Else 'NOT CKHUM...
        fHs1 = -1
        fHs2 = -1
    End If
    picWaterfall.SetFocus

End Sub

Private Sub ckTicks_Click()

    WantsTicks = (ckTicks = vbChecked)
    picWaterfall.SetFocus

End Sub

Private Sub ckTimeWindow_Click()

    cFourier.WithTimeWindow = (ckTimeWindow = vbChecked)
    picWaterfall.SetFocus

End Sub

Private Function ConvertToColor(Value As Double) As Long

    Select Case Value
      Case Is > TwoFiveFive
        Value = TwoFiveFive
      Case Is < 0
        Value = 0
    End Select
    Saturation = TwoFiveFive - Value
    If Saturation = 0 Then
        ConvertToColor = RGB(TwoFiveFive, TwoFiveFive, TwoFiveFive)
      Else 'NOT SATURATION...
        If Value <= TwoFiveFive / 2 Then
            FactA = Value * (TwoFiveFive + Saturation) / TwoFiveFive
          Else 'NOT Value...
            FactA = Value + Saturation - Value * Saturation / TwoFiveFive
        End If
        FactB = Value + Value - FactA
        ConvertToColor = RGB(HUEtoRGB(WtfColor + TwoFiveFive / 3, FactA, FactB), HUEtoRGB(WtfColor, FactA, FactB), HUEtoRGB(WtfColor - TwoFiveFive / 3, FactA, FactB))
    End If

End Function

Private Sub EvalParam()

  Dim w As Long

    NumTimesRes = NumSamples * Resolution
    lbFFT = "FFT: " & NumTimesRes & " Bins"
    w = NumTimesRes / 2 - 1
    With scrRange
        If w < Wid Then
            .Enabled = False
            RangeMin = 0
            RangeMax = w
          Else 'NOT W...
            .Max = w - Wid - 1
            .Enabled = True
        End If
        scrRange_Change
    End With 'SCRRANGE
    ckHum_Click

End Sub

Private Sub FFT()

  Dim i         As Long
  Dim j         As Long
  Dim yz        As Single
  Dim TickTime  As String

    Max = 1
    Do
        QueryPerformanceCounter AtStart
        j = 0
        If SoundBufferIsReady And Not Freeze Then
            With cFourier
                'fill in FFT samples
                For i = PtrOverlap To PtrOverlap + NumTimesRes - 1
                    j = j + 1
                    .RealIn(j) = SoundGetSample(i, NumTimesRes)
                Next i
                If PWVisible Then 'hide the wait notice
                    lbBuffering.Visible = False
                    PWVisible = False
                    btMagn.Enabled = True
                End If
                DoEvents
                With picSpectrum
                    .ScaleHeight = Max * 1.05
                    yz = Max * 1.025 - Shift * Max
                    Max = 0.1
                    .Cls
                    .CurrentX = -1
                    .CurrentY = yz
                    j = 0
                    For i = RangeMin To RangeMax
                        tmp = Smoothed(cFourier.ComplexOut(i + 1), j)
                        If tmp > Max Then 'find biggest out-value so that we can scale the picbox
                            Max = tmp
                        End If
                        'draw the resulting spectrum
                        picSpectrum.Line -(j, yz - tmp * Gain)
                        'draw waterfall
                        Select Case i
                          Case fHs1 To fHs2
                            tmp = tmp / HumSupp
                        End Select
                        picWaterfall.PSet (j, 1), ConvertToColor(tmp * Contrast + Brightness)
                        j = j + 1
                    Next i
                End With 'picSpectrum 'CFOURIER
                'time ticks
                With picWaterfall
                    BitBlt .hDC, 0, 1, Wid, Hgt, .hDC, 0, 0, vbSrcCopy
                    TickTime = CStr(TimeValue(Now))
                    TickCount = TickCount + 1
                    TickGate = (TickGate Or (Right$(TickTime, 1) = "9")) And TickCount > 49
                    If WantsTicks And TickGate Then
                        If Right$(TickTime, 1) = "0" Then
                            picWaterfall.Line (.ScaleWidth, 2)-(0, 2)
                            .CurrentY = .CurrentY + 1
                            picWaterfall.Print TickTime
                            TickGate = False
                            TickCount = 0
                        End If
                    End If
                End With 'picWaterfall
            End With 'CFOURIER
        End If
        DoEvents
        If Delay Then
            Sleep Delay
            DoEvents
        End If
        If MagnLoaded Then
            fMagnifier.Repaint
            DoEvents
        End If
        If j Then 'we had a cycle so we calculate the overlap stride...
            QueryPerformanceCounter AtEnd
            PtrOverlap = PtrOverlap + Ovrlap * (AtEnd - AtStart) / CPUSpeed
        End If
    Loop Until Break
    If Not Finito Then
        WaitTime = Resolution * 2
        tmrWait.Enabled = True
        tmrWait_Timer
        lbBuffering.Visible = True
        PWVisible = True
    End If

End Sub

Private Sub Form_Load()

  Dim i As Long

    If InIDE Then
        MsgBox "Please compile me; FFT is 15 times faster when compiled.", , "Fourier Transformation"
    End If
    Show
    Resolution = 1 'these two are here to prevent division by zero
    LastPixelNum = 1
    DoEvents
    Set cFourier = New clsFourier
    QueryPerformanceFrequency CPUSpeed

    For i = 0 To 5
        optSamples(i).Caption = 2 ^ (i + 10)
    Next i
    With picWaterfall
        Wid = .ScaleWidth
        LastPixelNum = Wid - 1
        Hgt = .ScaleWidth
        .SetFocus
    End With 'picWaterfall
    ReDim SmfFilters(0 To Wid)
    For i = 0 To Wid
        ReDim SmfFilters(i).SmfPrevValues(0 To 6) 'currenty max num of filter weights is 7
    Next i
    optSamples(1) = True
    optWeighted = True
    ckTimeWindow = vbChecked
    scrHue = 120
    scrGain_Change
    scrShift_Change
    scrContrast_Change
    scrBright_Change
    scrSpeed_Change
    scrReso_Change
    ckTicks = vbChecked
    If SoundCheckDevice = False Then
        Unload Me
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lbFreq = vbNullString

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SoundStopRecording
    tmrWait.Enabled = False
    Finito = True
    Break = True

End Sub

Private Sub fr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    lbFreq = vbNullString

End Sub

Private Function HUEtoRGB(ByVal Hue As Long, Fa As Double, Fb As Double) As Long

    Select Case Hue
      Case Is < 0
        Hue = Hue + TwoFiveFive
      Case Is > TwoFiveFive
        Hue = Hue - TwoFiveFive
    End Select
    Select Case Hue
      Case Is < TwoFiveFive / 6
        HUEtoRGB = Fb + 6 * (Fa - Fb) * Hue / TwoFiveFive
      Case Is < TwoFiveFive / 2
        HUEtoRGB = Fa
      Case Is < TwoFiveFive * 2 / 3
        HUEtoRGB = Fb + 6 * (Fa - Fb) * (TwoFiveFive * 2 / 3 - Hue) / TwoFiveFive
      Case Else
        HUEtoRGB = Fb
    End Select

End Function

Private Function InIDE(Optional c As Boolean = False) As Boolean

  Static b As Boolean

    b = c
    If b = False Then
        Debug.Assert InIDE(True)
    End If
    InIDE = b

End Function

Private Sub mnuAbout_Click()

    Load frmAbout
    With frmAbout
        .Theme = 14
        .AppIcon(vbBlack) = Icon
        .Title(vbGreen) = App.ProductName
        .Version(&HC0C0A0) = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        .Copyright(&HC0C0A0) = App.LegalCopyright
        .Otherstuff1(vbYellow) = "Low Frequency Spectrum Analyzer"
        .Otherstuff2(vbYellow) = "Use microphone or line input to display spectrum and waterfall of frequencies from 1Hz to 16384Hz with a resolution of 0.03Hz"
        .Show vbModal, Me
    End With 'FRMABOUT
    Set frmAbout = Nothing

End Sub

Private Sub mnuExit_Click()

    Unload Me

End Sub

Private Sub mnuSelDev_Click()

  'this opens the recording device selection dialog window

    If Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,2", vbNormalFocus) Then
        SendKeys "%" & ShortcutU & "%{F4}", True
    End If
    picWaterfall.SetFocus

End Sub

Private Sub mnuStart_Click()

    picWaterfall.SetFocus
    mnuAbout.Enabled = False
    ckFreeze = vbUnchecked
    Freeze = False
    picWaterfall.Cls
    WaitTime = Resolution * 2
    lbBuffering.Visible = True
    tmrWait_Timer
    tmrWait.Enabled = True
    PWVisible = True
    Finito = False
    Break = False
    DoEvents
    Do
        cFourier.NumberOfSamples = NumTimesRes
        Break = False
        If SoundStartRecording(NumTimesRes, NumSamples) Then
            Ovrlap = NumSamples * 1.05  '1.05 because we don't wanna be late for breakfast
            lbFreq.Visible = True
            FFT
          Else 'NOT SOUNDSTARTRECORDING(NUMTIMESRES,...
            lbBuffering.Visible = False
            Finito = True
        End If
    Loop Until Finito
    If MagnLoaded Then
        MagnLoaded = False
        Unload fMagnifier
    End If

End Sub

Private Sub mnuStop_Click()

    mnuAbout.Enabled = True
    picWaterfall.SetFocus
    lbBuffering.Visible = False
    tmrWait.Enabled = False
    Finito = True
    Break = True
    btMagn.Enabled = False

End Sub

Private Sub optFilterNone_Click()

    SmfWeights = Array(0, 1)
    picWaterfall.SetFocus

End Sub

Private Sub optSamples_Click(Index As Integer)

    NumSamples = 2 ^ (Index + 10)
    With scrReso
        .Enabled = False
        .Value = 0
        .Min = 5 - Index
        .Enabled = True
    End With 'SCRRESO
    EvalParam
    Break = True
    picWaterfall.SetFocus

End Sub

Private Sub optSavGol_Click()

    SmfWeights = Array(0.086, -0.143, -0.086, 0.257, 0.886) 'savitzky golay 2-4-0
    picWaterfall.SetFocus

End Sub

Private Sub optWeighted_Click()

    SmfWeights = Array(0.02, 0.03, 0.05, 0.1, 0.2, 0.3, 0.3) 'weighted running average
    picWaterfall.SetFocus

End Sub

Private Sub picSpectrum_GotFocus()

    picWaterfall.SetFocus

End Sub

Private Sub picSpectrum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lbBuffering.Visible Then
        lbFreq = vbNullString
      Else 'LBBUFFERING.VISIBLE = FALSE/0
        With picSpectrum
            lbFreq = Round(FreqMin + (FreqMax - FreqMin) * X / LastPixelNum)
            lbFreq.Left = .Left + ScaleX(X, vbPixels, ScaleMode) - lbFreq.Width / 2
        End With 'picSpectrum
    End If

End Sub

Private Sub picWaterfall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picSpectrum_MouseMove Button, Shift, X, Y

End Sub

Private Sub scrBright_Change()

    Brightness = scrBright
    picWaterfall.SetFocus

End Sub

Private Sub scrBright_Scroll()

    scrBright_Change

End Sub

Private Sub scrContrast_Change()

    Contrast = 1.5 ^ scrContrast
    HumSupp = Contrast / 3 + 0.667
    picWaterfall.SetFocus

End Sub

Private Sub scrContrast_Scroll()

    scrContrast_Change

End Sub

Private Sub scrGain_Change()

    Gain = scrGain
    picWaterfall.SetFocus

End Sub

Private Sub scrGain_Scroll()

    scrGain_Change

End Sub

Private Sub scrHue_Change()

    WtfColor = scrHue
    picWaterfall.ForeColor = ConvertToColor(160)
    picSpectrum.ForeColor = ConvertToColor(128)
    picWaterfall.SetFocus

End Sub

Private Sub scrHue_Scroll()

    scrHue_Change

End Sub

Private Sub scrRange_Change()

    If scrRange.Enabled Then
        RangeMin = scrRange
        RangeMax = RangeMin + Wid
        FreqMax = Int(RangeMax / Resolution) + 1
      Else 'SCRRANGE.ENABLED = FALSE/0
        FreqMax = Wid + 1
    End If
    FreqMin = Int(RangeMin / Resolution) + 1
    lbrange = FreqMin & " - " & FreqMax
    picWaterfall.SetFocus

End Sub

Private Sub scrRange_Scroll()

    scrRange_Change

End Sub

Private Sub scrReso_Change()

    Resolution = 2 ^ scrReso
    lbReso = "x " & Resolution
    EvalParam
    Break = True
    picWaterfall.SetFocus

End Sub

Private Sub scrReso_Scroll()

    scrReso_Change

End Sub

Private Sub scrShift_Change()

    Shift = scrShift / 110
    picWaterfall.SetFocus

End Sub

Private Sub scrShift_Scroll()

    scrShift_Change

End Sub

Private Sub scrSpeed_Change()

    Delay = scrSpeed ^ 1.2
    picWaterfall.SetFocus

End Sub

Private Sub scrSpeed_Scroll()

    scrSpeed_Change

End Sub

Private Function Smoothed(Value As Double, ByVal FilterNum As Long) As Double

  Dim i As Long
  Dim j As Long
  Dim k As Long

  'smoothing filters (each point has it's own filter)

    k = UBound(SmfWeights)
    Smoothed = Value * SmfWeights(k)
    With SmfFilters(FilterNum)
        j = .SmfPtr - 1
        For i = k - 1 To 0 Step -1
            j = (j + 1) Mod k
            Smoothed = Smoothed + .SmfPrevValues(j) * SmfWeights(i)
        Next i
        .SmfPtr = j 'now j points to the oldest recent value
        .SmfPrevValues(j) = Value
    End With 'SMFFILTERS(FILTERNUM)

End Function

Private Sub tmrWait_Timer()

    lbBuffering = "Buffering - Please wait " & WaitTime & " Secs"
    WaitTime = WaitTime - 1
    If WaitTime = 0 Then
        tmrWait.Enabled = False
    End If

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Apr-24 11:50)  Decl: 90  Code: 599  Total: 689 Lines
':) CommentOnly: 35 (5,1%)  Commented: 40 (5,8%)  Empty: 156 (22,6%)  Max Logic Depth: 7
