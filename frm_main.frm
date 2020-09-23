VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Win 8 Metro"
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8880
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10425
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox pic_tmp_icoperso 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   10140
      ScaleHeight     =   600
      ScaleWidth      =   885
      TabIndex        =   97
      Top             =   645
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox Text_name 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2925
      TabIndex        =   96
      Text            =   "Text1"
      Top             =   1515
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "utiliser mais invisible"
      Height          =   4920
      Left            =   16275
      TabIndex        =   63
      Top             =   7860
      Visible         =   0   'False
      Width           =   4170
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1290
         Left            =   240
         TabIndex        =   80
         Top             =   3480
         Width           =   3540
         Begin VB.CheckBox Check1 
            Caption         =   "Alignement"
            Height          =   195
            Left            =   3195
            TabIndex        =   87
            ToolTipText     =   "Advance Alignement/Taille"
            Top             =   195
            Width           =   210
         End
         Begin VB.TextBox txt_copier 
            BackColor       =   &H00D8E9EC&
            DragMode        =   1  'Automatic
            Height          =   285
            Left            =   2235
            Locked          =   -1  'True
            TabIndex        =   84
            Top             =   150
            Width           =   660
         End
         Begin VB.PictureBox pic_coller 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00D8E9EC&
            DragMode        =   1  'Automatic
            Height          =   300
            Left            =   2250
            Picture         =   "frm_main.frx":0CCA
            ScaleHeight     =   240
            ScaleWidth      =   255
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   540
            Width           =   315
         End
         Begin VB.ComboBox Combo_dim 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   150
            Width           =   1380
         End
         Begin VB.ComboBox Combo_dim_destination 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   540
            Width           =   1380
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Picker le    : "
            Height          =   195
            Left            =   150
            TabIndex        =   86
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Coller sur :"
            Height          =   285
            Left            =   150
            TabIndex        =   85
            Top             =   570
            Width           =   945
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1290
         Left            =   195
         TabIndex        =   73
         Top             =   2085
         Width           =   3645
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            Height          =   765
            Left            =   1665
            ScaleHeight     =   765
            ScaleWidth      =   930
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   315
            Width           =   930
            Begin VB.CommandButton cmd_heightminus 
               Caption         =   "H-"
               Height          =   315
               Left            =   495
               TabIndex        =   78
               Top             =   435
               Width           =   420
            End
            Begin VB.CommandButton cmd_widthminus 
               Caption         =   "W-"
               Height          =   315
               Left            =   480
               TabIndex        =   77
               Top             =   45
               Width           =   420
            End
            Begin VB.CommandButton cmd_height 
               Caption         =   "H+"
               Height          =   315
               Left            =   15
               TabIndex        =   76
               Top             =   435
               Width           =   420
            End
            Begin VB.CommandButton cmd_width 
               Caption         =   "W+"
               Height          =   315
               Left            =   15
               TabIndex        =   75
               Top             =   45
               Width           =   420
            End
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Touche de Redimentionnement"
            Height          =   585
            Left            =   285
            TabIndex        =   79
            Top             =   375
            Width           =   1440
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1560
         Left            =   165
         TabIndex        =   64
         Top             =   435
         Width           =   3705
         Begin VB.CommandButton cmd_cree_button 
            Height          =   330
            Left            =   2955
            Picture         =   "frm_main.frx":1080
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Nouveau Programme"
            Top             =   270
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            Height          =   930
            Left            =   1560
            ScaleHeight     =   930
            ScaleWidth      =   1035
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   360
            Width           =   1035
            Begin VB.CommandButton cmd_deplacright 
               BackColor       =   &H00D8E9EC&
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Marlett"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   500
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   690
               TabIndex        =   70
               Top             =   300
               Width           =   330
            End
            Begin VB.CommandButton cmd_deplacbottom 
               BackColor       =   &H00D8E9EC&
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "Marlett"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   500
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   345
               TabIndex        =   69
               Top             =   630
               Width           =   345
            End
            Begin VB.CommandButton cmd_deplactop 
               BackColor       =   &H00D8E9EC&
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "Marlett"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   500
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   345
               TabIndex        =   68
               Top             =   0
               Width           =   345
            End
            Begin VB.CommandButton cmd_deplacleft 
               BackColor       =   &H00D8E9EC&
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Marlett"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   500
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   0
               TabIndex        =   67
               Top             =   315
               Width           =   360
            End
         End
         Begin VB.CommandButton Command1 
            Caption         =   "annuler suppressions"
            Height          =   225
            Left            =   2775
            TabIndex        =   65
            Top             =   1095
            Visible         =   0   'False
            Width           =   195
         End
      End
   End
   Begin VB.Timer timer_unload_pic_prog 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8085
      Top             =   285
   End
   Begin VB.PictureBox pic_close_sur_mousemove 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   6375
      Picture         =   "frm_main.frx":160A
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   61
      ToolTipText     =   "Supprimer ce Programme"
      Top             =   495
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   5685
      Top             =   1230
   End
   Begin VB.PictureBox pic_system 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      ForeColor       =   &H00808080&
      Height          =   720
      Index           =   5
      Left            =   4365
      Picture         =   "frm_main.frx":1748
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   5970
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic_desktop 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   0
      ScaleHeight     =   3600
      ScaleWidth      =   5925
      TabIndex        =   41
      Top             =   2235
      Visible         =   0   'False
      Width           =   5955
      Begin VB.PictureBox pic_close_desktopico 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1605
         Picture         =   "frm_main.frx":2B8A
         ScaleHeight     =   300
         ScaleWidth      =   720
         TabIndex        =   43
         Top             =   75
         Width           =   720
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   1515
         Left            =   15
         TabIndex        =   42
         Top             =   15
         Width           =   1785
         ExtentX         =   3149
         ExtentY         =   2672
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   -1  'True
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5205
      Top             =   1680
   End
   Begin VB.PictureBox pic_sel 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6420
      Picture         =   "frm_main.frx":370C
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6195
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pic_min_close 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   9840
      Picture         =   "frm_main.frx":3BC2
      ScaleHeight     =   210
      ScaleWidth      =   420
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   195
      Width           =   420
      Begin VB.Label label_close 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   150
         Left            =   225
         TabIndex        =   32
         ToolTipText     =   "Quitter"
         Top             =   30
         Width           =   150
      End
      Begin VB.Label label_min 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   150
         Left            =   30
         TabIndex        =   31
         ToolTipText     =   "Réduire"
         Top             =   30
         Width           =   150
      End
   End
   Begin VB.PictureBox pic_folder 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   7380
      Picture         =   "frm_main.frx":409C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6405
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic_shortcut 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   7800
      Picture         =   "frm_main.frx":4626
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox ongletsel 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   8520
      Picture         =   "frm_main.frx":4A68
      ScaleHeight     =   315
      ScaleWidth      =   810
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox ongletnormal 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   9405
      Picture         =   "frm_main.frx":581E
      ScaleHeight     =   270
      ScaleWidth      =   810
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox pic_frame_config 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   6750
      Picture         =   "frm_main.frx":63E8
      ScaleHeight     =   5055
      ScaleWidth      =   3105
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   3105
      Begin VB.PictureBox pic_onglet 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   2
         Left            =   1680
         ScaleHeight     =   300
         ScaleWidth      =   750
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   390
         Width           =   750
         Begin VB.Label Label_onglet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Advanced"
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   20
            Top             =   45
            Width           =   720
         End
      End
      Begin VB.PictureBox pic_onglet 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   1
         Left            =   885
         ScaleHeight     =   300
         ScaleWidth      =   750
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   390
         Width           =   750
         Begin VB.Label Label_onglet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edition"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   19
            Top             =   45
            Width           =   480
         End
      End
      Begin VB.PictureBox pic_onglet 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   0
         Left            =   90
         ScaleHeight     =   300
         ScaleWidth      =   750
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   390
         Width           =   750
         Begin VB.Label Label_onglet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   18
            Top             =   45
            Width           =   375
         End
      End
      Begin VB.CheckBox Check_editable 
         Caption         =   "Mode Edition"
         Height          =   345
         Left            =   2475
         TabIndex        =   5
         Top             =   285
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Frame Frame_config 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Paramètre"
         ForeColor       =   &H00800000&
         Height          =   4230
         Left            =   105
         TabIndex        =   21
         Top             =   750
         Visible         =   0   'False
         Width           =   2910
         Begin VB.PictureBox pic_onglet 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Height          =   300
            Index           =   4
            Left            =   900
            ScaleHeight     =   300
            ScaleWidth      =   750
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   15
            Width           =   750
            Begin VB.Label Label_onglet 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Aspect"
               Height          =   195
               Index           =   4
               Left            =   180
               TabIndex        =   103
               Top             =   45
               Width           =   495
            End
         End
         Begin VB.PictureBox pic_onglet 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Height          =   300
            Index           =   3
            Left            =   120
            ScaleHeight     =   300
            ScaleWidth      =   750
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   15
            Width           =   750
            Begin VB.Label Label_onglet 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Param"
               Height          =   195
               Index           =   3
               Left            =   180
               TabIndex        =   101
               Top             =   45
               Width           =   450
            End
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   225
            Left            =   1785
            TabIndex        =   99
            Top             =   3300
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            _Version        =   393216
            Max             =   255
            SelStart        =   150
            TickStyle       =   3
            Value           =   150
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            Height          =   570
            Left            =   855
            ScaleHeight     =   570
            ScaleWidth      =   1905
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1620
            Width           =   1905
            Begin VB.CommandButton cmd_recharger_param 
               Caption         =   "Restore"
               Height          =   510
               Left            =   990
               TabIndex        =   47
               ToolTipText     =   "Restaurer la dernière configuration"
               Top             =   30
               Width           =   900
            End
            Begin VB.CommandButton cmd_save_config 
               Caption         =   "Save"
               Height          =   510
               Left            =   30
               TabIndex        =   46
               ToolTipText     =   "Sauvegarder la configuration actuelle"
               Top             =   30
               Width           =   900
            End
         End
         Begin VB.PictureBox pic_system 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H00808080&
            Height          =   720
            Index           =   4
            Left            =   1830
            Picture         =   "frm_main.frx":99CA
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   28
            TabStop         =   0   'False
            Tag             =   "special;rundll32.exe shell32.dll,Control_RunDLL"
            ToolTipText     =   "Panneau de Configuration"
            Top             =   2490
            Width           =   720
         End
         Begin VB.PictureBox pic_system 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H00808080&
            Height          =   720
            Index           =   3
            Left            =   1020
            Picture         =   "frm_main.frx":A894
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   27
            TabStop         =   0   'False
            Tag             =   "special;::{645FF040-5081-101B-9F08-00AA002F954E}"
            ToolTipText     =   "Corbeille"
            Top             =   3270
            Width           =   720
         End
         Begin VB.PictureBox pic_system 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H00808080&
            Height          =   720
            Index           =   2
            Left            =   210
            Picture         =   "frm_main.frx":B75E
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   26
            TabStop         =   0   'False
            Tag             =   "special;::{208D2C60-3AEA-1069-A2D7-08002B30309D}"
            ToolTipText     =   "Connexions réseau"
            Top             =   3270
            Width           =   720
         End
         Begin VB.PictureBox pic_system 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H00808080&
            Height          =   720
            Index           =   1
            Left            =   1020
            Picture         =   "frm_main.frx":C628
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   25
            TabStop         =   0   'False
            Tag             =   "special;::{450D8FBA-AD25-11D0-98A8-0800361B1103}"
            ToolTipText     =   "Mes documents"
            Top             =   2490
            Width           =   720
         End
         Begin VB.PictureBox pic_system 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H00808080&
            Height          =   720
            Index           =   0
            Left            =   210
            Picture         =   "frm_main.frx":D4F2
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   24
            TabStop         =   0   'False
            Tag             =   "special;::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
            ToolTipText     =   "Poste de Travail"
            Top             =   2490
            Width           =   720
         End
         Begin VB.Frame Frame_aspect 
            BackColor       =   &H00D8E9EC&
            Height          =   1290
            Left            =   105
            TabIndex        =   105
            Top             =   285
            Width           =   2715
            Begin VB.CheckBox Check_roundcorner 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Round corners"
               Height          =   270
               Left            =   645
               TabIndex        =   116
               Top             =   990
               Width           =   1545
            End
            Begin VB.CheckBox Check_centrer 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Vertical"
               Height          =   240
               Left            =   645
               TabIndex        =   113
               Top             =   210
               Width           =   900
            End
            Begin VB.CheckBox Check_ico_text_horizontal 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Horizontal"
               Height          =   225
               Left            =   1590
               TabIndex        =   112
               Top             =   225
               Width           =   1035
            End
            Begin VB.CheckBox Check_cadre 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Frame"
               Height          =   225
               Left            =   645
               TabIndex        =   111
               ToolTipText     =   "Afficher un Cadre autour des Programmes"
               Top             =   510
               Width           =   840
            End
            Begin VB.CheckBox simule_3d 
               BackColor       =   &H00D8E9EC&
               Caption         =   "3D Buton"
               Height          =   225
               Left            =   1590
               TabIndex        =   110
               ToolTipText     =   "Simule 3D button en MouseMove"
               Top             =   525
               Width           =   1035
            End
            Begin VB.CheckBox check_change_backcolor_mousemove 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Change Color"
               Height          =   255
               Left            =   645
               TabIndex        =   109
               ToolTipText     =   "Changer Couleur qd souris au dessus d'un Programme"
               Top             =   750
               Width           =   1470
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Style : "
               Height          =   195
               Left            =   135
               TabIndex        =   114
               Top             =   210
               Width           =   510
            End
         End
         Begin VB.Frame Frame_param 
            BackColor       =   &H00D8E9EC&
            Height          =   1290
            Left            =   105
            TabIndex        =   104
            Top             =   285
            Width           =   2715
            Begin VB.CheckBox Check_hide_desktop_icon 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Hide Desktop Icones"
               Height          =   435
               Left            =   105
               TabIndex        =   108
               Top             =   120
               Width           =   2370
            End
            Begin VB.CheckBox check_indice_folder 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Folder Indicator"
               Height          =   285
               Left            =   105
               TabIndex        =   107
               Top             =   540
               Width           =   1755
            End
            Begin VB.CheckBox Check_caption_statusbar 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Show path in StatusBar"
               Height          =   360
               Left            =   105
               TabIndex        =   106
               Top             =   855
               Width           =   2550
            End
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Dragdrop icone to empty space)"
            Height          =   195
            Left            =   225
            TabIndex        =   35
            Top             =   3990
            Width           =   2370
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Insert system icone : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   150
            TabIndex        =   34
            Top             =   2190
            Width           =   2565
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Config : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   150
            TabIndex        =   29
            Top             =   1770
            Width           =   660
         End
      End
      Begin VB.Frame Frame_couleur 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Couleur picker"
         ForeColor       =   &H00800000&
         Height          =   4230
         Left            =   105
         TabIndex        =   6
         Top             =   750
         Width           =   2910
         Begin VB.CheckBox Check_gradient 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Soften the colors with a gradient"
            Height          =   405
            Left            =   165
            TabIndex        =   115
            Top             =   3675
            Width           =   2475
         End
         Begin VB.ComboBox Combo_color 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   240
            Width           =   840
         End
         Begin VB.PictureBox Shape_pick_color 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   150
            ScaleHeight     =   210
            ScaleWidth      =   210
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   285
            Width           =   240
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   2910
            Left            =   165
            Picture         =   "frm_main.frx":E3BC
            ScaleHeight     =   2910
            ScaleWidth      =   2460
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   645
            Width           =   2460
            Begin VB.Image imgMarker 
               Enabled         =   0   'False
               Height          =   165
               Left            =   1395
               Picture         =   "frm_main.frx":258D6
               Top             =   540
               Width           =   165
            End
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Apply Color  to :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1155
            TabIndex        =   36
            Top             =   210
            Width           =   900
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            Height          =   480
            Left            =   465
            MouseIcon       =   "frm_main.frx":25922
            Picture         =   "frm_main.frx":25A74
            ToolTipText     =   "déplacer pour choisir une couleur"
            Top             =   180
            Width           =   480
         End
         Begin VB.Image imgcircle 
            Height          =   480
            Left            =   2265
            Picture         =   "frm_main.frx":25BC6
            Top             =   495
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Frame Frame_mode_edition 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Mode Edition"
         ForeColor       =   &H00800000&
         Height          =   4230
         Left            =   105
         TabIndex        =   9
         Top             =   750
         Visible         =   0   'False
         Width           =   2910
         Begin VB.CheckBox Check2 
            Caption         =   "Magnetisme"
            Height          =   195
            Left            =   2640
            TabIndex        =   95
            ToolTipText     =   "Magnetisme"
            Top             =   1380
            Width           =   225
         End
         Begin VB.CheckBox check_move_libre 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Align using the Grid"
            Height          =   255
            Left            =   195
            TabIndex        =   62
            Top             =   525
            Value           =   1  'Checked
            Width           =   1770
         End
         Begin VB.Frame Frame_magnetisme 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Magnetism "
            ForeColor       =   &H00800000&
            Height          =   2430
            Left            =   60
            TabIndex        =   90
            Top             =   1530
            Visible         =   0   'False
            Width           =   2790
            Begin VB.ComboBox combo_valeur_espace 
               Height          =   315
               ItemData        =   "frm_main.frx":25D18
               Left            =   1440
               List            =   "frm_main.frx":25D1A
               Style           =   2  'Dropdown List
               TabIndex        =   93
               Top             =   795
               Width           =   1215
            End
            Begin VB.CommandButton Command8 
               BackColor       =   &H00FFFFFF&
               Height          =   585
               Left            =   120
               Picture         =   "frm_main.frx":25D1C
               Style           =   1  'Graphical
               TabIndex        =   92
               Top             =   300
               Width           =   1200
            End
            Begin VB.CommandButton Command9 
               BackColor       =   &H00FFFFFF&
               Height          =   570
               Left            =   120
               Picture         =   "frm_main.frx":2873A
               Style           =   1  'Graphical
               TabIndex        =   91
               Top             =   1005
               Width           =   1200
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "with : "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   1455
               TabIndex        =   94
               Top             =   480
               Width           =   450
            End
         End
         Begin VB.Frame Frame_alignement 
            BackColor       =   &H00D8E9EC&
            Height          =   1515
            Left            =   60
            TabIndex        =   49
            Top             =   2175
            Width           =   2790
            Begin VB.ComboBox combo_uniformiser_taille 
               Height          =   315
               Left            =   1515
               Style           =   2  'Dropdown List
               TabIndex        =   89
               Top             =   1065
               Width           =   1155
            End
            Begin VB.PictureBox pic_uniforme_les_deux 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   2340
               Picture         =   "frm_main.frx":2A7C0
               ScaleHeight     =   240
               ScaleWidth      =   270
               TabIndex        =   57
               ToolTipText     =   "Les Deux"
               Top             =   480
               Width           =   270
            End
            Begin VB.PictureBox pic_uniforme_hauteur 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1950
               Picture         =   "frm_main.frx":2AB82
               ScaleHeight     =   240
               ScaleWidth      =   255
               TabIndex        =   56
               ToolTipText     =   "Hauteur"
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox pic_uniforme_largeur 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   1575
               Picture         =   "frm_main.frx":2AF04
               ScaleHeight     =   225
               ScaleWidth      =   255
               TabIndex        =   55
               ToolTipText     =   "Largeur"
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox pic_alignement_gauche 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   390
               Picture         =   "frm_main.frx":2B252
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   54
               ToolTipText     =   "Gauche"
               Top             =   360
               Width           =   225
            End
            Begin VB.PictureBox pic_alignement_droit 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   945
               Picture         =   "frm_main.frx":2B564
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   53
               ToolTipText     =   "Droite"
               Top             =   360
               Width           =   225
            End
            Begin VB.PictureBox pic_alignement_bas 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   690
               Picture         =   "frm_main.frx":2B876
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   52
               ToolTipText     =   "Bas"
               Top             =   570
               Width           =   225
            End
            Begin VB.PictureBox pic_alignement_haut 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   690
               Picture         =   "frm_main.frx":2BB88
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   51
               ToolTipText     =   "Haut"
               Top             =   225
               Width           =   225
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "using "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   1560
               TabIndex        =   88
               Top             =   795
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Uniform Size"
               Height          =   195
               Index           =   2
               Left            =   1485
               TabIndex        =   58
               Top             =   150
               Width           =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Align"
               Height          =   195
               Index           =   1
               Left            =   105
               TabIndex        =   50
               Top             =   150
               Width           =   345
            End
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Move : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   0
            Left            =   195
            TabIndex        =   71
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label_desell_all 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Deselect All"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1095
            TabIndex        =   40
            Top             =   3975
            Width           =   825
         End
         Begin VB.Label Label_sel_all 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "select All"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   195
            TabIndex        =   39
            Top             =   3975
            Width           =   630
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "(Mouse or Keys : Shift+Arrows to resize)"
            Height          =   435
            Left            =   150
            TabIndex        =   12
            Top             =   1575
            Width           =   2535
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "____________________________"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   1290
            Width           =   2520
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "(For more precision, you can use les touches CTRL+Arrow keys)"
            Height          =   585
            Left            =   120
            TabIndex        =   10
            Top             =   825
            Width           =   2655
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00BDA591&
         X1              =   3375
         X2              =   30
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00B4562C&
         BorderWidth     =   2
         Height          =   255
         Left            =   15
         Top             =   15
         Width           =   210
      End
      Begin VB.Label label0 
         Caption         =   " "
         Height          =   195
         Left            =   2820
         TabIndex        =   4
         ToolTipText     =   "Fermer"
         Top             =   90
         Width           =   195
      End
   End
   Begin VB.PictureBox pic_barre 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8F3F0&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   15735
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   $"frm_main.frx":2BE9A
      Top             =   6870
      Width           =   15765
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   30
         Picture         =   "frm_main.frx":2BF39
         ScaleHeight     =   360
         ScaleWidth      =   4170
         TabIndex        =   98
         ToolTipText     =   "Rechercher  ... (Bureau & Menu démarrer)"
         Top             =   30
         Width           =   4170
         Begin VB.TextBox txt_search 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   75
            TabIndex        =   117
            Top             =   45
            Width           =   3690
         End
      End
      Begin VB.PictureBox pic_show_desktop 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6645
         Picture         =   "frm_main.frx":30DDB
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   120
         Width           =   240
      End
      Begin VB.PictureBox pic_resize 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   6855
         MousePointer    =   8  'Size NW SE
         Picture         =   "frm_main.frx":310BD
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   165
         Width           =   165
      End
      Begin VB.Label label_caption 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8F3F0&
         Caption         =   "F1 = Astuce"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4260
         TabIndex        =   33
         Top             =   75
         Width           =   1005
      End
   End
   Begin VB.PictureBox pic_prog 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E3BD00&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1020
      Index           =   0
      Left            =   360
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   1020
      ScaleWidth      =   1185
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   1185
   End
   Begin VB.Shape SelectBox 
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   4605
      Top             =   330
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label_horloge2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   6375
      OLEDropMode     =   1  'Manual
      TabIndex        =   60
      Top             =   8070
      Width           =   1185
   End
   Begin VB.Label Label_horloge1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1155
      Left            =   6270
      OLEDropMode     =   1  'Manual
      TabIndex        =   59
      Top             =   7125
      Width           =   2445
   End
   Begin VB.Shape Shape_form 
      Height          =   540
      Left            =   15
      Top             =   15
      Width           =   510
   End
   Begin VB.Shape Shape_sel2 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   2775
      Top             =   495
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape_sel 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   1935
      Shape           =   2  'Oval
      Top             =   630
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Public Declare Function SendMessage Lib _
  "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long

'activer la prise en charge style thème xp
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long


Dim mheight As Integer

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTBOTTOMRIGHT = 17
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

'determine le vrai width et height de screen (espace non reservé)
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, _
    ByVal uParam As Long, _
    lpvParam As Any, _
    ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Dim ResizeMe As Boolean
Dim ctl As CControlSizer

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public indexencours As Integer
''Public indexencours_conteneur As Integer

Dim oldindex  As Integer
Dim old_backcolor As Long
Dim info_corbeille As String
Dim simule3d As Boolean

Private Type POINTAPI
  X As Long
  Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private pT As POINTAPI
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long


'Public Const SHGFI_ICON = &H100 'pour avoir l'icone avec le petit indicateur raccourcis (shortcut)
Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000  ' System icon index
Private Const SHGFI_LARGEICON = &H0        ' Large icon
Private Const SHGFI_SMALLICON = &H1        ' Small icon
Private Const ILD_TRANSPARENT = &H1        ' Display transparent
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
   Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
   Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
   (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal X&, ByVal Y&, ByVal flags&) As Long

Private ShInfo As SHFILEINFO


Private Const DT_CALCRECT = &H400
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_EXPANDTABS = &H40
Private Const DT_NOPREFIX = &H800
Private Const DT_WORDBREAK = &H10
Private Const DT_TABSTOP = &H80
Private Const DT_WORD_ELLIPSIS = &H40000
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
   (ByVal hdc As Long, _
    ByVal lpString As String, _
    ByVal nCount As Long, _
    lpRect As RECT, _
    ByVal uFormat As Long) As Long
Private Declare Function GetWindowRect Lib "user32" _
  (ByVal hwnd As Long, _
   lpRect As RECT) As Long


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'desktop ico show/hide
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long



'draw raccourci icon
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Const DI_NORMAL = &H3

Dim xlocation As Integer
Dim ylocation As Integer
Dim dragdropico As Boolean

'pour open file dialog (au lieu de l'ocx commondialog)
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type





'traitement recycle bin
''Private Const S_OK = &H0
''Private Type ULARGE_INTEGER
'' LowPart As Long
'' HighPart As Long
''End Type
''Private Type SHQUERYRBINFO
'' cbSize As Long
'' i64Size As ULARGE_INTEGER
'' i64NumItems As ULARGE_INTEGER
''End Type
''Private Declare Function SHQueryRecycleBin Lib "shell32.dll" _
''Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, _
''pSHQueryRBInfo As SHQUERYRBINFO) As Long
''
''Private Declare Function EmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Enum rbEmptyBinEnum
  SHERB_NOCONFIRMATION = &H1
  SHERB_NOPROGRESSUI = &H2
  SHERB_NOSOUND = &H4
End Enum
Private Declare Function EmptyRecycleBin Lib "shell32" Alias "SHEmptyRecycleBinA" _
    (Optional ByVal hwnd As Long = 0, _
      Optional ByVal pszRootPath As String = vbNullString, _
      Optional ByVal dwFlags As rbEmptyBinEnum = 0) _
    As Long

Private Type SHQUERYRBINFO
  cbSize As Long
  i64Size As Currency
  i64NumItems As Currency
End Type
Private Declare Function SHQueryRecycleBin Lib "shell32" Alias "SHQueryRecycleBinA" _
    (ByVal pszRootPath As String, _
      pSHQueryRBInfo As SHQUERYRBINFO) _
    As Long
Dim corbeille_vide As Boolean
Dim c_size As Variant
Dim num_items As Variant



Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long


'Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long


'pour la transparence

Private Const LWA_COLORKEY = 1
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Dim top_pic_barre As Integer
Dim barre_visible As Boolean
Public mparent As String
Dim deplacment_effectue As Boolean


'=========== la Selection par souris =============
'Private Type SelectGroup
'    blnOnClipBrd As Boolean
'    ActiveGrip As Integer
'    ctl As Control
'    dX1 As Single
'    dY1 As Single
'    dX2 As Single
'    dY2 As Single
'End Type
'Private SelectedCtl() As SelectGroup
'Private NumInGrp As Integer

Dim blnDragStarted As Boolean
Dim blnMouseIsDown As Boolean
Dim StartX As Single
Dim StartY As Single



Sub alignement(Pos As String)


Dim valeur As Integer
Dim aucune_selection As Boolean

    aucune_selection = True
    'décté la valeur la plus adequat
    For i = 0 To pic_prog.UBound
        If pic_prog(i).WhatsThisHelpID = 1 Then
            aucune_selection = False: Exit For
        End If
    Next

''    If aucune_selection = True Then
''        For i = 0 To pic_conteneur.uBound
''            If pic_conteneur(i).WhatsThisHelpID = 1 Then
''                aucune_selection = False: Exit For
''            End If
''        Next
''    End If
    
    If aucune_selection = True Then Exit Sub
    
    'une valeur par defaut
    If Pos = "gauche" Then
        valeur = Me.Width
    ElseIf Pos = "droit" Then
        valeur = 0
    ElseIf Pos = "haut" Then
        valeur = Me.Height
    ElseIf Pos = "bas" Then
        valeur = 0
    End If

    'détecté la valeur la plus adequat
    For i = 0 To pic_prog.UBound
        If pic_prog(i).WhatsThisHelpID = 1 Then 'Or i = index Then 'déjà sel
            If Pos = "gauche" Then
                If pic_prog(i).Left < valeur Then valeur = pic_prog(i).Left
            ElseIf Pos = "droit" Then
                If pic_prog(i).Left + pic_prog(i).Width > valeur Then valeur = pic_prog(i).Left + pic_prog(i).Width
            ElseIf Pos = "haut" Then
                If pic_prog(i).Top < valeur Then valeur = pic_prog(i).Top
            ElseIf Pos = "bas" Then
                If pic_prog(i).Top + pic_prog(i).Height > valeur Then valeur = pic_prog(i).Top + pic_prog(i).Height
            End If
        End If
    Next i

''    'détecté la valeur la plus adequat conteneur
''    For i = 0 To pic_conteneur.uBound
''        If pic_conteneur(i).WhatsThisHelpID = 1 Then 'Or i = index Then 'déjà sel
''            If pos = "gauche" Then
''                If pic_conteneur(i).Left < valeur Then valeur = pic_conteneur(i).Left
''            ElseIf pos = "droit" Then
''                If pic_conteneur(i).Left + pic_conteneur(i).Width > valeur Then valeur = pic_conteneur(i).Left + pic_conteneur(i).Width
''            ElseIf pos = "haut" Then
''                If pic_conteneur(i).Top < valeur Then valeur = pic_conteneur(i).Top
''            ElseIf pos = "bas" Then
''                If pic_conteneur(i).Top + pic_conteneur(i).Height > valeur Then valeur = pic_conteneur(i).Top + pic_conteneur(i).Height
''            End If
''        End If
''    Next i
    
    'appliquer sur tous les objets sélectionnés
    For i = 0 To pic_prog.UBound
        If pic_prog(i).WhatsThisHelpID = 1 Then 'Or i = index Then 'déjà sel
            If Pos = "gauche" Then
                pic_prog(i).Left = valeur
            ElseIf Pos = "droit" Then
                pic_prog(i).Left = valeur - pic_prog(i).Width
            ElseIf Pos = "haut" Then
                pic_prog(i).Top = valeur
            ElseIf Pos = "bas" Then
                pic_prog(i).Top = valeur - pic_prog(i).Height
            End If
        End If
    Next i

''    'appliquer sur tous les objets sélectionnés
''    For i = 0 To pic_conteneur.uBound
''        If pic_conteneur(i).WhatsThisHelpID = 1 Then 'Or i = index Then 'déjà sel
''            If pos = "gauche" Then
''                pic_conteneur(i).Left = valeur
''            ElseIf pos = "droit" Then
''                pic_conteneur(i).Left = valeur - pic_conteneur(i).Width
''            ElseIf pos = "haut" Then
''                pic_conteneur(i).Top = valeur
''            ElseIf pos = "bas" Then
''                pic_conteneur(i).Top = valeur - pic_conteneur(i).Height
''            End If
''        End If
''    Next i
    
'    'indexencours = index
'    'pour avoir les même dimension
'    pic_prog(indexencours).ZOrder 0
'    ctl.AttachControl pic_prog(indexencours)
''If indexencours <> -1 And indexencours_conteneur <> -1 Then

''Elseif
If indexencours <> -1 Then
    pic_prog(indexencours).ZOrder 0
    ctl.AttachControl pic_prog(indexencours)
''Else
''    pic_conteneur(indexencours_conteneur).ZOrder 0
''    ctl.AttachControl pic_conteneur(indexencours_conteneur)
End If

End Sub

Sub alignement_height_ou_width(espace As Integer, sens As String)

Dim petit As Integer
Dim dimension_destination As Integer


    If sens = "top+height" Then
        petit = Me.Height
        For i = 0 To pic_prog.UBound
            If pic_prog(i).WhatsThisHelpID = 1 Or i = indexencours Then 'déjà sel
                If petit > pic_prog(i).Top Then petit = pic_prog(i).Top: dimension_destination = pic_prog(i).Top + pic_prog(i).Height + espace '15
            End If
        Next i
        'If Combo_dim = "Left+Width" Then txt_copier = pic_prog(index).Left + pic_prog(index).Width + 15
        'appliquer sur tous les objets sélectionnés
        For i = 0 To pic_prog.UBound
            If pic_prog(i).WhatsThisHelpID = 1 Or i = indexencours Then 'déjà sel
                If pic_prog(i).Top <> petit Then pic_prog(i).Top = dimension_destination
            End If
        Next i
        If indexencours <> -1 Then ctl.AttachControl pic_prog(indexencours)
    Else 'left+width
        petit = Me.Width
        For i = 0 To pic_prog.UBound
            If pic_prog(i).WhatsThisHelpID = 1 Or i = indexencours Then 'déjà sel
                If petit > pic_prog(i).Left Then petit = pic_prog(i).Left: dimension_destination = pic_prog(i).Left + pic_prog(i).Width + espace '15
            End If
        Next i
        'appliquer sur tous les objets sélectionnés
        For i = 0 To pic_prog.UBound
            If pic_prog(i).WhatsThisHelpID = 1 Or i = indexencours Then 'déjà sel
                If pic_prog(i).Left <> petit Then pic_prog(i).Left = dimension_destination
            End If
        Next i
        If indexencours <> -1 Then ctl.AttachControl pic_prog(indexencours)
    End If

End Sub

Sub animation_barre(show_hide As Integer)

Timer2.Enabled = False
'Me.Enabled = False

'1 = show
'0 = hide

Dim i As Integer
    
'Shape_form.Visible = False : Label_horloge1.Visible = False: Label_horloge2.Visible = False  'pour éviter le flicker

If show_hide = 1 Then 'show
    
    pic_barre.ZOrder 0: pic_min_close.ZOrder 0
    pic_min_close.Top = -285
    pic_barre.Top = Me.Height + 60
    pic_min_close.Visible = True
    pic_barre.Visible = True
    
    For i = 1 To 290 Step 10
        pic_min_close.Top = pic_min_close.Top + 10 ' 15 / i '(285 - i)
        pic_barre.Top = pic_barre.Top - 10 '(285 * i / 100)
        'Me.Refresh
        Sleep 10
        DoEvents
        'pic_barre.Refresh: pic_min_close.Refresh

    Next i
    
    pic_min_close.Top = 15
    pic_barre.Top = Height - pic_barre.Height

    barre_visible = True

Else

    
    For i = 1 To 290 Step 10
        pic_min_close.Top = pic_min_close.Top - 10 ' 15 / i '(285 - i)
        pic_barre.Top = pic_barre.Top + 10 '(285 * i / 100)
        'Me.Refresh
        Sleep 10
        DoEvents
        'pic_barre.Refresh: pic_min_close.Refresh

    Next i
    
    pic_min_close.Visible = False
    pic_barre.Visible = False
    
    barre_visible = False

End If

'Shape_form.Visible = True: Label_horloge1.Visible = True: Label_horloge2.Visible = True
Me.Refresh
Timer2.Enabled = True
'Me.Enabled = True

End Sub

Sub horloge()


If Label_horloge1 <> Format$(Now, "hh:mm") Then Label_horloge1 = Format$(Now, "hh:mm")

If Label_horloge2 <> Format$(Now, "dd mmm yyyy") Then Label_horloge2 = Format$(Now, "dd mmm yyyy")

Label_horloge2.Move Label_horloge1.Left - ((Label_horloge2.Width - Label_horloge1.Width) / 2), Label_horloge1.Top + Label_horloge1.Height - 100 '+ 5






End Sub

Sub key_lancer_prog(touche As String)

Dim i As Integer
Dim Pos As Integer
Dim nom As String

For i = 0 To pic_prog.UBound
    
    If pic_prog(i).Visible Then
        nom = mcaption(i)
        Pos = InStr(1, nom, "&")
        If Pos <> 0 And Pos <> Len(nom) Then
            If LCase$(touche) = LCase$(Mid$(nom, Pos + 1, 1)) Then pic_prog_MouseDown i, 1, 0, 0, 0: Exit Sub
        End If
    End If
    
Next i

End Sub

Sub mon_special_ico_quelquesoit_la_taille(chemin As String, taille As Integer)




Dim hImage As Long
hImage = LoadImage(0, chemin, 1, taille, taille, 16) ', 16, LR_LOADFROMFILE) 'App.hInstance

If hImage <> 0 Then
    pic_tmp_icoperso.Picture = Nothing
    ImageList1.ListImages.Clear
    ImageList1.ListImages.Add 1, "p1", IconToPicture(hImage)
    Set pic_tmp_icoperso.Picture = ImageList1.ListImages(1).ExtractIcon
End If


End Sub

Sub move_frm_param(index As Integer)

Dim mleft As Integer
Dim mtop As Integer

'frm_param.Show

mleft = Me.Left + pic_prog(index).Left + pic_prog(index).Width + 15
If mleft + frm_param.Width > Me.Left + Me.Width Then mleft = Me.Left + Me.Width - frm_param.Width - 15

mtop = Me.Top + pic_prog(index).Top + 15
If mtop + frm_param.Height > Me.Top + Me.Height Then mtop = Me.Top + Me.Height - frm_param.Height - 15


frm_param.Move mleft, mtop


End Sub

Sub rond_pic(index As Integer, cadre As Boolean)

If cadre Then
        If Me.Picture <> 0 Then
            'pic_prog(index).Cls
            pic_prog(index).PSet (15, 15), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (0, 0), Me.Point(pic_prog(index).Left, pic_prog(index).Top) ' GetPixel(Me.hdc, pic_prog(index).Left, pic_prog(index).Top)
    
            pic_prog(index).PSet (pic_prog(index).Width - 30, 15), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (pic_prog(index).Width - 15, 0), Me.Point(pic_prog(index).Left + pic_prog(index).Width - 15, pic_prog(index).Top) 'vbWhite
    
            pic_prog(index).PSet (15, pic_prog(index).Height - 30), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (0, pic_prog(index).Height - 15), Me.Point(pic_prog(index).Left, pic_prog(index).Top + pic_prog(index).Height - 15) 'vbWhite
    
            pic_prog(index).PSet (pic_prog(index).Width - 30, pic_prog(index).Height - 30), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (pic_prog(index).Width - 15, pic_prog(index).Height - 15), Me.Point(pic_prog(index).Left + pic_prog(index).Width - 15, pic_prog(index).Top + pic_prog(index).Height - 15) 'vbWhite
        Else
            'pic_prog(index).Cls
            pic_prog(index).PSet (15, 15), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (0, 0), Me.BackColor
    
            pic_prog(index).PSet (pic_prog(index).Width - 30, 15), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (pic_prog(index).Width - 15, 0), Me.BackColor
    
            pic_prog(index).PSet (15, pic_prog(index).Height - 30), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (0, pic_prog(index).Height - 15), Me.Point(pic_prog(index).Left, pic_prog(index).Top + pic_prog(index).Height - 15) 'vbWhite
    
            pic_prog(index).PSet (pic_prog(index).Width - 30, pic_prog(index).Height - 30), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (pic_prog(index).Width - 15, pic_prog(index).Height - 15), Me.BackColor
        End If
Else
        If Me.Picture <> 0 Then
            'pic_prog(index).Cls
            'pic_prog(index).PSet (15, 15), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (0, 0), Me.Point(pic_prog(index).Left, pic_prog(index).Top) ' GetPixel(Me.hdc, pic_prog(index).Left, pic_prog(index).Top)
    
            'pic_prog(index).PSet (pic_prog(index).Width - 30, 15), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (pic_prog(index).Width - 15, 0), Me.Point(pic_prog(index).Left + pic_prog(index).Width - 15, pic_prog(index).Top) 'vbWhite
    
            'pic_prog(index).PSet (15, pic_prog(index).Height - 30), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (0, pic_prog(index).Height - 15), Me.Point(pic_prog(index).Left, pic_prog(index).Top + pic_prog(index).Height - 15) 'vbWhite
    
            'pic_prog(index).PSet (pic_prog(index).Width - 30, pic_prog(index).Height - 30), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (pic_prog(index).Width - 15, pic_prog(index).Height - 15), Me.Point(pic_prog(index).Left + pic_prog(index).Width - 15, pic_prog(index).Top + pic_prog(index).Height - 15) 'vbWhite
        Else
            'pic_prog(index).Cls
            'pic_prog(index).PSet (15, 15), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (0, 0), Me.BackColor
    
            'pic_prog(index).PSet (pic_prog(index).Width - 30, 15), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (pic_prog(index).Width - 15, 0), Me.BackColor
    
            'pic_prog(index).PSet (15, pic_prog(index).Height - 30), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (0, pic_prog(index).Height - 15), Me.Point(pic_prog(index).Left, pic_prog(index).Top + pic_prog(index).Height - 15) 'vbWhite
    
            'pic_prog(index).PSet (pic_prog(index).Width - 30, pic_prog(index).Height - 30), BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            pic_prog(index).PSet (pic_prog(index).Width - 15, pic_prog(index).Height - 15), Me.BackColor
        End If
End If

End Sub

Sub print_txt(X As Integer, Y As Integer, index As Integer, txt As String)

Dim rc As RECT
    
    rc.Left = X / Screen.TwipsPerPixelX
    rc.Top = Y / Screen.TwipsPerPixelY
    rc.Right = pic_prog(index).Width 'pic_prog(Index).CurrentX + pic_prog(Index).TextWidth(txt)
    rc.Bottom = pic_prog(index).Height 'pic_prog(Index).CurrentY + pic_prog(Index).TextHeight(txt)
    
    DrawText pic_prog(index).hdc, txt, Len(txt), rc, DT_SINGLELINE 'DT_CENTER


End Sub

Sub traitement_special_dossier(folder As String)

'pic_show_desktop.Enabled = False 'pour empêcher le click alors que la procédure est toujours en cours
 'Me.Enabled = False
    'shellfoldernav (&H0)
    myfoldernav folder
    pic_desktop.Width = Me.Width
    pic_desktop.Height = Me.Height ' - pic_barre.Height '(Me.Height / 2) '- pic_barre.Height
    pic_desktop.Top = 0 '(Me.Height / 2) - pic_barre.Height
    WebBrowser1.Width = pic_desktop.Width - 60
    WebBrowser1.Height = pic_desktop.Height - 60
    pic_close_desktopico.Move pic_desktop.Width - pic_close_desktopico.Width - 280, 15
    
    AnimateForm Me, pic_desktop, -1, -1, aload, 15, 5, 2, 25
    pic_desktop.Visible = True
    pic_desktop.ZOrder 0
    WebBrowser1.SetFocus
    WebBrowser1.ZOrder 0
    pic_close_desktopico.ZOrder 0
'Me.Enabled = True


End Sub

Sub uniformiser(Pos As String)


Dim h As Integer
Dim w As Integer
Dim aucune_selection As Boolean

    aucune_selection = True
    
    For i = 0 To pic_prog.UBound
        If pic_prog(i).WhatsThisHelpID = 1 Then
            aucune_selection = False: Exit For
        End If
    Next

    
    If aucune_selection = True Then Exit Sub
    
    If combo_uniformiser_taille = "Petit" Then
        w = Me.Width
        h = Me.Height
        'décté la valeur la plus adequat
        For i = 0 To pic_prog.UBound
            If pic_prog(i).WhatsThisHelpID = 1 Then 'Or i = index Then 'déjà sel
                If Pos = "largeur" Then
                    If pic_prog(i).Width < w Then w = pic_prog(i).Width
                ElseIf Pos = "hauteur" Then
                    If pic_prog(i).Height < h Then h = pic_prog(i).Height
                ElseIf Pos = "lesdeux" Then
                    If pic_prog(i).Width < w Then w = pic_prog(i).Width
                    If pic_prog(i).Height < h Then h = pic_prog(i).Height
                End If
            End If
        Next i
    Else 'au plus grand
        w = 0 'Me.Width
        h = 0 'Me.Height
        'décté la valeur la plus adequat
        For i = 0 To pic_prog.UBound
            If pic_prog(i).WhatsThisHelpID = 1 Then 'Or i = index Then 'déjà sel
                If Pos = "largeur" Then
                    If pic_prog(i).Width > w Then w = pic_prog(i).Width
                ElseIf Pos = "hauteur" Then
                    If pic_prog(i).Height > h Then h = pic_prog(i).Height
                ElseIf Pos = "lesdeux" Then
                    If pic_prog(i).Width > w Then w = pic_prog(i).Width
                    If pic_prog(i).Height > h Then h = pic_prog(i).Height
                End If
            End If
        Next i
    End If
    
    'appliquer sur tous les objets sélectionnés
    For i = 0 To pic_prog.UBound
        If pic_prog(i).WhatsThisHelpID = 1 Then 'Or i = index Then 'déjà sel
            If Pos = "largeur" Then
                pic_prog(i).Width = w
            ElseIf Pos = "hauteur" Then
                pic_prog(i).Height = h
            ElseIf Pos = "haut" Then
                pic_prog(i).Top = valeur
            ElseIf Pos = "lesdeux" Then
                pic_prog(i).Width = w
                pic_prog(i).Height = h
            End If
        End If
    Next i


''ElseIf
If indexencours <> -1 Then
    pic_prog(indexencours).ZOrder 0
    ctl.AttachControl pic_prog(indexencours)
''Else
''    pic_conteneur(indexencours_conteneur).ZOrder 0
''    ctl.AttachControl pic_conteneur(indexencours_conteneur)
End If

End Sub


Sub QueryRecycleBin(Optional ByVal RootPath As String = vbNullString, _
                            Optional ByRef Size As Variant, _
                            Optional ByRef NumItems As Variant)
Dim pQRBI As SHQUERYRBINFO
  
  'Make sure variants are initialized to
  ' Decimal subtype (largest numeric)
  Size = CDec(0)
  NumItems = CDec(0)
  
  With pQRBI
    .cbSize = Len(pQRBI)
    
    SHQueryRecycleBin RootPath, _
                      pQRBI
    
    'The Currency data type is the only pure
    ' 64-bits numeric data type VB has, but
    ' it inserts a comma at the fourth
    ' position from the right...
    'Let's correct this!
    Size = .i64Size * 10000
    NumItems = .i64NumItems * 10000
  End With
End Sub
Private Function BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long
'pour mon gradient
    Dim lCFrom As Long
    Dim lCTo As Long
    Dim lSrcR As Long
    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
   
    lCFrom = GetLngColor(oColorFrom)
    lCTo = GetLngColor(oColorTo)
    
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
   
    BlendColor = RGB( _
      ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
      ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
      ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))
      
End Function

Private Function GetLngColor(Color As Long) As Long
'pour mon gradient
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function
Sub animation()

'AnimateForm pic_desktop,
AnimateForm Me, pic_desktop, -1, -1, aload, 5, 15, 2, 25



End Sub
Public Sub no_transparence()

Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
Call SetLayeredWindowAttributes(Me.hwnd, 0, 255, LWA_ALPHA)

End Sub

Function shellfoldernav(id As Long) As Integer
  
  On Error GoTo err_seven
  
  Dim nFolder As SpecialShellFolderIDs
  Dim pidl As Long
  Dim cbpidl As Integer
  Dim abpidl() As Byte
  Dim avpidl As Variant
  Dim sPath As Long
  
  'Label1 = ""
  DoEvents
  MousePointer = vbHourglass
  
  nFolder = id '&H0 'Combo1.ItemData(Combo1.ListIndex)
  
  ' Get the pointer to the folder's item ID list from
  ' it's specified folder ID, returns 0 on success
  If SHGetSpecialFolderLocation(hwnd, nFolder, pidl) = NOERROR Then
    If pidl Then
      
      ' Get the folder's item ID list size
      cbpidl = GetPIDLSize(pidl)
      If cbpidl Then
        
        ' Reallocate the byte array and copy the folder's item ID list to the array.
        ReDim abpidl(cbpidl - 1)   ' array is zero based
        MoveMemory abpidl(0), ByVal pidl, cbpidl
        
        ' Load the pidl's byte aray into the variant, tada, a SAFEARRAY...
        avpidl = abpidl
        
        ' Navigate the browser to the folder's pidl...!!!
        WebBrowser1.Navigate2 avpidl
        'WebBrowser1.Visible = True
        
      End If   ' cbpidl
      
      ' Free the memory the shell allocated for the pidl
      Call CoTaskMemFree(pidl)
      
    End If   ' pidl
  End If   ' SHGetSpecialFolderLocation
'
'  ' Show what's happening with the folder...
'  If (pidl = 0) Then
'    'Label1 = "The folder does not exist on this system."
'
'  Else
'    Label1 = GetSpecialFolderPath(hWnd, nFolder)
'    If (Len(Label1) = 0) Then
'      If (nFolder <= CSIDL_TEMPLATES) Then
'        Label1 = "Folder has no path (is a vitual folder)."
'       Else
'        Label1 = "Is either a vitual folder, or Shell32.dll is < v4.71 " & _
'                       "and it doesn't know about the CSIDL value)"
'      End If
'    End If
'
'  End If
  
  MousePointer = vbDefault
  shellfoldernav = 0
  Exit Function
  
err_seven:
'la corbeille s'ouvre bien sous xp mais pas sous seven donc si erreur
'ouvrir explorer et pointer vers la corbeille
  
  'If corbeille Then
    'Shell "explorer.exe /n," & Trim$(Mid$(pic_prog(index).Tag, Len("special") + 2, Len(pic_prog(index).Tag))), vbNormalFocus
    shellfoldernav = 1   'donc error
  'End If
  
  MousePointer = vbDefault

End Function

Function myfoldernav(folder As String) As Integer
  
  'On Error GoTo err_seven
  
  Dim nFolder As SpecialShellFolderIDs
  Dim pidl As Long
  Dim cbpidl As Integer
  Dim abpidl() As Byte
  Dim avpidl As Variant
  Dim sPath As Long
  
  'Label1 = ""
  DoEvents
  MousePointer = vbHourglass
  
    WebBrowser1.Navigate2 folder 'avpidl
  
  
  
  MousePointer = vbDefault
  myfoldernav = 0
  Exit Function
  
err_seven:
'la corbeille s'ouvre bien sous xp mais pas sous seven donc si erreur
'ouvrir explorer et pointer vers la corbeille
  
  'If corbeille Then
    'Shell "explorer.exe /n," & Trim$(Mid$(pic_prog(index).Tag, Len("special") + 2, Len(pic_prog(index).Tag))), vbNormalFocus
    myfoldernav = 1   'donc error
  'End If
  
  MousePointer = vbDefault

End Function


Function empty_recycle()

'refresh info corbeille
info_corbeille = "": Timer1_Timer


If MsgBox("Etes-vous sûr de vouloir vider la Corbeille?", vbYesNo + vbInformation, "Vider la Corbeille") = vbNo Then Exit Function

'State the Variables
Dim Para As Long, nRet As Long

'Para = 7
nRet = EmptyRecycleBin(frmmain.hwnd, vbNullString, 7) 'Para)


End Function
Sub actualise_caption(index As Integer, Optional txt As String, Optional posy As Integer, Optional special As Boolean)



Dim icochemin As String
Dim taille As String * 3
    If txt = "" Then txt = mcaption(index)
    
    'pic_prog(index).CurrentX = 150: pic_prog(index).CurrentY = posy '650
    
    'If Check_centrer.Value = 1 Then pic_prog(index).CurrentX = (pic_prog(index).Width - TextWidth(txt)) / 2
    
    'If index = 1 Then Stop
    
    'cas spécial corbeille = pas de centrage
    If pic_prog(index).Tag = "special;::{645FF040-5081-101B-9F08-00AA002F954E}" Then
        pic_prog(index).CurrentY = posy: pic_prog(index).CurrentX = 150
        'pic_prog(index).Print txt
        print_txt pic_prog(index).CurrentX, pic_prog(index).CurrentY, index, txt
        Exit Sub
    End If
    
        'si pas icon system alors vérifie si icone perso
        'If Not special Then
            'si ico perso ajuster position caption
            icochemin = GetFromINI("config", "progico" & index, App.path & "\config.ini")
            If icochemin <> "" And File_Exist(icochemin) = True Then 'icone perso
                taille = GetFromINI("config", "progicotaille" & index, App.path & "\config.ini")
                If Trim$(taille) = "" Then taille = "128"
                pic_prog(index).CurrentY = pic_prog(index).Height - TextHeight(txt) - 45
                If Check_centrer.Value = 1 Then
                    pic_prog(index).CurrentX = (pic_prog(index).Width - TextWidth(txt)) / 2
                    'pic_prog(index).CurrentY = pic_prog(index).Height - TextHeight(txt) - 45
                    'If special Then
                    '    pic_prog(index).CurrentY = (pic_prog(index).Height + (128 * 15)) / 2  '+ TextHeight(txt)
                    'Else
                        ''pic_prog(index).CurrentY = (pic_prog(index).Height + (taille * 15)) / 2  '+ TextHeight(txt)
                    'End If
                ElseIf Check_ico_text_horizontal.Value = 1 Then
                    'pic_prog(index).CurrentX = (pic_prog(index).Width) / 4 + (32 * 15)
                    'pic_prog(index).CurrentX = (10 + 48) * Screen.TwipsPerPixelX
                    'If special Then
                    '    pic_prog(index).CurrentX = (10 + 128) * Screen.TwipsPerPixelX  '+ TextHeight(txt)
                    'Else
                        pic_prog(index).CurrentX = (13 + taille) * Screen.TwipsPerPixelX  '+ TextHeight(txt)
                    'End If
                    pic_prog(index).CurrentY = (pic_prog(index).Height - TextHeight(txt)) / 2
                Else
                    pic_prog(index).CurrentX = 150 ': pic_prog(index).CurrentY = posy '650
                    'pic_prog(index).CurrentY = (pic_prog(index).Height + (taille * 15)) / 2
                End If
                
                'pic_prog(index).Print txt
                print_txt pic_prog(index).CurrentX, pic_prog(index).CurrentY, index, txt
                Exit Sub

            End If
        'End If
'    If Check_ico_text_horizontal.Value = 1 Then
'        'pic_prog(index).CurrentX = (pic_prog(index).Width) / 4 + (32 * 15)
'        pic_prog(index).CurrentX = (12 + 32) * Screen.TwipsPerPixelX
'        pic_prog(index).CurrentY = (pic_prog(index).Height - TextHeight(txt)) / 2
'    End If
    
    
    If Check_centrer.Value = 1 Then
        pic_prog(index).CurrentX = (pic_prog(index).Width - TextWidth(txt)) / 2
        pic_prog(index).CurrentY = pic_prog(index).Height - TextHeight(txt) - 45 'posy '650
        'If special Then
        '    pic_prog(index).CurrentY = (pic_prog(index).Height + (48 * 15)) / 2  '+ TextHeight(txt)
        'Else
        '    pic_prog(index).CurrentY = (pic_prog(index).Height + (32 * 15)) / 2  '+ TextHeight(txt)
        'End If
    ElseIf Check_ico_text_horizontal.Value = 1 Then
        'pic_prog(index).CurrentX = (pic_prog(index).Width) / 4 + (32 * 15)
        'pic_prog(index).CurrentX = (10 + 48) * Screen.TwipsPerPixelX
        If special Then 'pour les ico system
            pic_prog(index).CurrentX = (10 + 48) * Screen.TwipsPerPixelX  '+ TextHeight(txt)
        Else
            pic_prog(index).CurrentX = (13 + 32) * Screen.TwipsPerPixelX  '+ TextHeight(txt)
        End If
        pic_prog(index).CurrentY = (pic_prog(index).Height - TextHeight(txt)) / 2
    Else
        'posy = pic_prog(index).Height - TextHeight(txt) - 45
        pic_prog(index).CurrentX = 150: pic_prog(index).CurrentY = pic_prog(index).Height - TextHeight(txt) - 45 'posy '650
    End If
    
    'pic_prog(Index).Print txt
    print_txt pic_prog(index).CurrentX, pic_prog(index).CurrentY, index, txt
    'pic_prog(Index).Refresh

End Sub

Public Function File_Exist(sFullPath As String, Optional recherche_folder As Boolean = False) As Boolean



    Dim wshShell As Object
    Dim wshLink As Object
    Set wshShell = CreateObject("Scripting.FileSystemObject") 'CreateObject("WScript.FileSystemObject")
    'Set wshLink = wshShell.FileExists(strPath)
    'File_folder_Exists
    File_Exist = wshShell.FileExists(sFullPath)
    
    'vérifie si c un dossier
    If recherche_folder Then
        'If File_Exist = False Then
            File_Exist = wshShell.FolderExists(sFullPath)
        'End If
    End If
    'GetTarget = wshLink.TargetPath
    Set wshLink = Nothing
    Set wshShell = Nothing
    
    
End Function
Private Function ExtractFileName(ByVal strPath As String) As String

Dim Pos As Integer

On Error GoTo err1
  ' StrReverse is only working in VB6
    Pos = InStr(1, strPath, "\")
    If Pos <> 0 Then
        strPath = StrReverse(strPath)
        strPath = Left(strPath, InStr(strPath, "\") - 1)
        ExtractFileName = StrReverse(strPath)
    Else
        ExtractFileName = Trim$(strPath)
    End If
    
Exit Function
err1:
  ExtractFileName = Trim$(strPath)
  Exit Function
  
End Function
Sub charger_onglet(index As Integer)

Dim i As Integer

If index = 1 Then 'mode edition
    Check_editable.Value = 1
    'marqué le prog indexencours comme selectionné
    If indexencours <> -1 Then
        pic_prog(indexencours).WhatsThisHelpID = 1
        actualise_ico_caption indexencours
        pic_prog(indexencours).ZOrder 0
        ctl.AttachControl pic_prog(indexencours)
    End If
    indexencours = -1 '': indexencours_conteneur = -1
Else
    Check_editable.Value = 0
End If

Frame_couleur.Visible = False
Frame_mode_edition.Visible = False
Frame_config.Visible = False


If index = 0 Then
    Frame_couleur.Visible = True
ElseIf index = 1 Then
    Frame_mode_edition.Visible = True
ElseIf index = 2 Then
    Frame_config.Visible = True
'    If Frame_param.Visible = True Then
'        charger_onglet 4
'    Else
'        charger_onglet 5
'    End If
'ElseIf index = 3 Then
'    Frame_config.Visible = True
'    Frame_param.Visible = True
'ElseIf index = 4 Then
'    Frame_config.Visible = True
'    Frame_aspect.Visible = True
End If

For i = 0 To 2 'pic_onglet.UBound
    
    If i = index Then
        pic_onglet(i).Picture = ongletsel.Picture
'        If i < 3 Then
            pic_onglet(i).Top = 360
'        Else
'            pic_onglet(i).Top = 0
'        End If
        pic_onglet(i).ZOrder 0
        Shape_sel.Visible = False
    Else
        pic_onglet(i).Picture = ongletnormal.Picture
'        If i < 3 Then
            pic_onglet(i).Top = 400
'        Else
'            pic_onglet(i).Top = 40
'        End If
        pic_barre.ZOrder 0
    End If
    
Next i

If index <> 1 Then 'eliminer les selection
    For i = 0 To pic_prog.UBound
        If pic_prog(i).WhatsThisHelpID = 1 Then pic_prog(i).WhatsThisHelpID = 0: actualise_ico_caption i
        'pic_prog(i).DragMode = 0
    Next i
End If

End Sub

Sub charger_onglet2(index As Integer)

Dim i As Integer

'If index = 1 Then 'mode edition
'    Check_editable.Value = 1
'    'marqué le prog indexencours comme selectionné
'    If indexencours <> -1 Then
'        pic_prog(indexencours).WhatsThisHelpID = 1
'        actualise_ico_caption indexencours
'        pic_prog(indexencours).ZOrder 0
'        ctl.AttachControl pic_prog(indexencours)
'    End If
'    indexencours = -1 '': indexencours_conteneur = -1
'Else
'    Check_editable.Value = 0
'End If

Frame_param.Visible = False
Frame_aspect.Visible = False

For i = 3 To pic_onglet.UBound
    
    If i = index Then
        pic_onglet(i).Picture = ongletsel.Picture
        pic_onglet(i).Top = 0
        pic_onglet(i).ZOrder 0
        Shape_sel.Visible = False
    Else
        pic_onglet(i).Picture = ongletnormal.Picture
        pic_onglet(i).Top = 40
        pic_barre.ZOrder 0
    End If
    
Next i

If index = 3 Then
    'Frame_config.Visible = True
    Frame_param.Visible = True
ElseIf index = 4 Then
    'Frame_config.Visible = True
    Frame_aspect.Visible = True
End If




End Sub


Function indexiconcorbeille() As Integer

Dim i As Integer

indexiconcorbeille = -1

For i = 0 To pic_prog.UBound
    'Debug.Print mcaption(i)
    If mcaption(i) = "::{645FF040-5081-101B-9F08-00AA002F954E}" Then indexiconcorbeille = i: Exit For
Next i


End Function

Function label_special(index As Integer) As String

    If pic_prog(index).Tag = "special;::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then label_special = "Poste de Travail"
    If pic_prog(index).Tag = "special;::{450D8FBA-AD25-11D0-98A8-0800361B1103}" Then label_special = "Mes documents"
    If pic_prog(index).Tag = "special;::{208D2C60-3AEA-1069-A2D7-08002B30309D}" Then label_special = "Connexions réseau"
    If pic_prog(index).Tag = "special;::{645FF040-5081-101B-9F08-00AA002F954E}" Then label_special = "Corbeille | Click Gauche = Ouvrir Corbeille | Click Droit = Vider Corbeille"
    If pic_prog(index).Tag = "special;rundll32.exe shell32.dll,Control_RunDLL" Then label_special = "Panneau de Configuration" '"Imprimantes et télécopieurs"


End Function

Sub loadwallpeper()

On Error GoTo err1
Dim f As String
    
    On Error Resume Next
    
    f = GetFromINI("config", "wallpaper", App.path & "\config.ini")

    If f <> "" Then
        If File_Exist(f) = True Then
            Me.Picture = LoadPicture(f)
        Else
            Me.Picture = Nothing
            Me.BackColor = GetFromINI("config", "formcolor", App.path & "\config.ini")
        End If
    Else
        Me.Picture = Nothing
        Me.BackColor = GetFromINI("config", "formcolor", App.path & "\config.ini")
    End If

Exit Sub

err1:
MsgBox "Erreur : " & err.Number & " | " & err.Description, vbCritical, App.FileDescription
Exit Sub

End Sub

Private Function replace_percent(txt As String) As String
    Dim Pos As Integer
    Dim chaine_avant As String
    Dim chaine_apres As String
    Dim chaine_percent As String
    Dim Text As String
    
'exemple
'%appdata%\microsoft\internet explorer\quick launch
'%HOMEDRIVE%\Documents and Settings\%username%
'%appdata%
'%allusersprofile%
'%USERPROFILE%
'%USERPROFILE%\Local Settings\Temporary Internet Files


'Environment variables
'MS DOS
'%APPEND%
'%BLASTER%
'%COMSPEC%
'%COPYCMD%
'%DIRCMD%
'%DOSSHELL%
'%MSDOSDATA%
'%NO_SEP%
'%PATH%
'%PROMPT%
'%TEMP%
'%TZ%
'additional variables in: MS WINDOWS
'%CMDLINE%
'%COMPUTERNAME%
'%TMP%
'%winbootdir%
'%windir%
'%WINPMT%
'additional variables in: MS WINDOWS NT
'%ALLUSERSPROFILE%
'%APPDATA%
'%CD%
'%CMDCMDLINE%
'%CMDEXTVERSION%
'%COMMONPROGRAMFILES%
'%DATE%
'%ERRORLEVEL%
'%HOMEDRIVE%
'%HOMEPATH%
'%HOMESHARE%
'%LOGONSERVER%
'%MORE%
'%NUMBER_OF_PROCESSORS%
'%OS%
'%OS2LIBPATH%
'%PATHEXT%
'%PROCESSOR_ARCHITECTURE%
'%PROCESSOR_IDENTIFER%
'%PROCESSOR_LEVEL%
'%PROCESSOR_REVISION%
'%PROGRAMFILES%
'%RANDOM%
'%SESSIONNAME%
'%SYSTEMDRIVE%
'%SYSTEMROOT%
'%TIME%
'%USERDOMAIN%
'%USERNAME%
'%USERPROFILE%

    
On Error GoTo err1
    
reappliquer:

    Text = txt
    Pos = InStr(1, Text, "%")
    If Pos <> 0 Then
        If Pos <> 1 Then
            chaine_avant = Mid$(Text, 1, Pos - 1)
        Else
            chaine_avant = ""
        End If
        
        chaine_percent = Mid$(Text, Pos + 1, InStr(Pos + 1, Text, "%") - Pos - 1)
        
        If Environ$(chaine_percent) <> "" Then
            Pos = InStr(Pos + 1, Text, "%") '- pos
            chaine_apres = Mid$(Text, Pos + 1, Len(Text))
            
            replace_percent = chaine_avant & Trim$(Environ$(chaine_percent)) & chaine_apres
            'Exit Function
            
            txt = replace_percent
            'vérifie si ya pas d'autre %xxx% ds la chaine
            Pos = InStr(1, txt, "%")
            If Pos <> 0 Then GoTo reappliquer
            
            
            
            'MsgBox Environ$(text)
        Else
            replace_percent = Text
        End If
        
    Else
        replace_percent = Text
    End If


Exit Function

err1:
    replace_percent = Text
End Function


Private Function GetTarget(strPath As String) As String
    'Gets target path from a shortcut file
    On Error GoTo Error_Loading
    Dim wshShell As Object
    Dim wshLink As Object
    Set wshShell = CreateObject("WScript.Shell")
    Set wshLink = wshShell.CreateShortcut(strPath)
    'récupérer le chemin du fichier source depuis le raccourcis
    GetTarget = wshLink.TargetPath
    Set wshLink = Nothing
    Set wshShell = Nothing
    
    'vérifie si le fichier source exist
    If File_folder_Exists(GetTarget) = False Then GetTarget = "" 'source raccourcis introuvable

    
    Exit Function

Error_Loading:
    If err.Number = -2147352567 Then 'Le chemin du raccourci doit finir par .lnk ou .url.
        'c pas un raccourcis donc envoyer le mm chemin
        GetTarget = strPath
    Else
        GetTarget = ""
    End If
    
    Exit Function
End Function

Private Function File_folder_Exists(sFullPath As String) As Boolean



    Dim wshShell As Object
    Dim wshLink As Object
    Set wshShell = CreateObject("Scripting.FileSystemObject") 'CreateObject("WScript.FileSystemObject")
    'Set wshLink = wshShell.FileExists(strPath)
    File_folder_Exists = wshShell.FileExists(sFullPath)
    
    'vérifie si c un dossier
    If File_folder_Exists = False Then
        File_folder_Exists = wshShell.FolderExists(sFullPath)
    End If
    'GetTarget = wshLink.TargetPath
    Set wshLink = Nothing
    Set wshShell = Nothing
    
    
End Function
Sub actualise_ico_caption(index As Integer, Optional txt As String)


    If pic_prog(index).Tag = "" Then Exit Sub
        
        If Mid$(pic_prog(index).Tag, 1, Len("special")) = "special" Then
            traitement_special index
            Exit Sub
        End If
        
        
'utilise le pic temporaire


        
        pic_prog(index).Cls
        'DrawGradient pic_prog(index).hDC, 0, 0, pic_prog(index).ScaleWidth / 15, pic_prog(index).ScaleHeight / 15, BlendColor(pic_prog(index).BackColor, vbBlack, 190), pic_prog(index).BackColor, False
        
        If Check_gradient.Value = 1 Then
            'If indexencours = index Then
            '    Call PaintGradient(pic_prog(index).hDC, 0, 0, pic_prog(index).ScaleWidth / 15, pic_prog(index).ScaleHeight / 15, BlendColor(pic_prog(index).BackColor, vbWhite, 100), pic_prog(index).BackColor, 264)
            'Else
                Call PaintGradient(pic_prog(index).hdc, 0, 0, pic_prog(index).ScaleWidth / 15, pic_prog(index).ScaleHeight / 15, BlendColor(pic_prog(index).BackColor, vbWhite, Slider1.Value), pic_prog(index).BackColor, 132)
            'End If
        End If
        'pic_prog(index).Picture = Nothing
        actualise_caption index, txt, 650
        
        DrawIcon mtag(index), index, 0, False
        
        pic_prog(index).Refresh
        
End Sub

Sub charge_param()

    'Me.Show
        
        
    On Error Resume Next
    Dim i As Integer, j As Integer
    Dim nbre As Integer
    Dim f As String
    Dim style As Integer
    Dim leconteneur As String
    'Dim leconteneur_control As PictureBox
    
    If File_Exist(App.path & "\config.ini") = False Then
        Call create_ini
    End If
   
    
    'Check_centrer.Value = GetFromINI("config", "centrer", App.path & "\config.ini")
    style = 0
    style = GetFromINI("config", "style", App.path & "\config.ini")
    If style = 1 Then
        Check_centrer.Value = 1
    ElseIf style = 2 Then
        Check_ico_text_horizontal.Value = 1
    Else
        Check_centrer.Value = 0: Check_ico_text_horizontal.Value = 0
    End If
    
    check_indice_folder.Value = GetFromINI("config", "indiceFolder", App.path & "\config.ini")
    Check_hide_desktop_icon.Value = GetFromINI("config", "desktopicohide", App.path & "\config.ini")
    Check_caption_statusbar.Value = GetFromINI("config", "showcaptionstatusbar", App.path & "\config.ini")
    
    Check_cadre.Value = GetFromINI("config", "affichecadre", App.path & "\config.ini")
    check_change_backcolor_mousemove.Value = GetFromINI("config", "changebackcolormousemove", App.path & "\config.ini")
    Check_gradient.Value = GetFromINI("config", "gradient", App.path & "\config.ini")
    Check_roundcorner.Value = GetFromINI("config", "roundcorner", App.path & "\config.ini")
    
    
    Call Check_hide_desktop_icon_Click

    Me.Left = GetFromINI("config", "formleft", App.path & "\config.ini")
    Me.Top = GetFromINI("config", "formtop", App.path & "\config.ini")
    Me.Width = GetFromINI("config", "formwidth", App.path & "\config.ini")
    Me.Height = GetFromINI("config", "formheight", App.path & "\config.ini")
    Me.BackColor = GetFromINI("config", "formcolor", App.path & "\config.ini")
    
    simule_3d.Value = GetFromINI("config", "simule3d", App.path & "\config.ini")
    Label_horloge1.ForeColor = GetFromINI("config", "horlogecolor", App.path & "\config.ini")
    Label_horloge2.ForeColor = Label_horloge1.ForeColor
    
'    Shape_sel.BorderColor = vbWhite
'    Shape_sel.BorderColor = GetFromINI("config", "cadrecolor", App.path & "\config.ini")
    
    'pic_bordure_color.BackColor = Shape_sel.BorderColor
    
    f = GetFromINI("config", "wallpaper", App.path & "\config.ini")
    If File_Exist(f) = True Then Me.Picture = LoadPicture(f)
    
    'décharger les controls existant
    If pic_prog.UBound > 0 Then
        For i = pic_prog.UBound To 1 Step -1
            Unload pic_prog(i)
        Next i
    End If
''    If pic_conteneur.uBound > 0 Then
''        For i = pic_conteneur.uBound To 1 Step -1
''            Unload pic_conteneur(i)
''        Next i
''    End If
    
    
    
     nbre = GetFromINI("config", "nbreprog", App.path & "\config.ini")

    
    For i = 0 To nbre
        If i <> 0 Then Load pic_prog(i)
        'pic_prog(i).Visible = True
        pic_prog(i).Tag = GetFromINI("config", "progtag" & i, App.path & "\config.ini")
        pic_prog(i).Left = GetFromINI("config", "progleft" & i, App.path & "\config.ini")
        pic_prog(i).Top = GetFromINI("config", "progtop" & i, App.path & "\config.ini")
        pic_prog(i).Width = GetFromINI("config", "progwidth" & i, App.path & "\config.ini")
        pic_prog(i).Height = GetFromINI("config", "progheight" & i, App.path & "\config.ini")
        pic_prog(i).BackColor = GetFromINI("config", "progcolor" & i, App.path & "\config.ini")
        pic_prog(i).ForeColor = 0
        pic_prog(i).ForeColor = GetFromINI("config", "progtextcolor" & i, App.path & "\config.ini")
        pic_prog(i).OLEDropMode = 1
        actualise_ico_caption i, mcaption(i)
    Next i
    
    For i = 0 To nbre
        pic_prog(i).Visible = True
    Next i







End Sub

Sub deplace_resize(index As Integer, what As Integer, valeur As String)

'what =
'0,left
'1,top
'2,width
'3,height
'pkoi valeur = string? parce que je peux pas passer le param -10 si c un integer (cela devient -1)

Dim i As Integer
Dim v As Integer

If indexencours = -1 Then Exit Sub ''And indexencours_conteneur = -1 Then Exit Sub 'aucun objet sélectionné


v = Val(valeur)



'appliquer sur tous les objets sélectionnés
For i = 0 To pic_prog.UBound
    If pic_prog(i).WhatsThisHelpID = 1 Or i = index Then 'déjà sel
        If what = 0 Then 'left
            pic_prog(i).Left = pic_prog(i).Left + v
        ElseIf what = 1 Then 'top
            pic_prog(i).Top = pic_prog(i).Top + v
        ElseIf what = 2 Then 'w
            pic_prog(i).Width = pic_prog(i).Width + v
        ElseIf what = 3 Then 'h
            pic_prog(i).Height = pic_prog(i).Height + v
        End If
    End If
Next i




'le cadre de selection doit suivre le pic_prog
'pic_prog(Index).ZOrder 0
'ctl.AttachControl pic_prog(Index)

''If indexencours <> -1 And indexencours_conteneur <> -1 Then

''ElseIf
If indexencours <> -1 Then
    pic_prog(indexencours).ZOrder 0
    ctl.AttachControl pic_prog(indexencours)
''Else
''    pic_conteneur(indexencours_conteneur).ZOrder 0
''    ctl.AttachControl pic_conteneur(indexencours_conteneur)
End If

End Sub

Sub dessine_icone(t As String, index As Integer)

    DrawIcon t, index, 0, False

    'pic_prog(index).PaintPicture pic16_intermediaire

End Sub

Sub DrawIcon(path As String, index As Integer, step_retour As Integer, isfodler As Boolean) ', Optional pointeur As Boolean)
Dim posx As Integer, posy As Integer
Dim taille As String * 3
Dim i As Integer
    Dim hImgLarge&
    Dim icochemin As String

''On Error Resume Next
'traitement spécial (poste de travail, corbeille ...

    pic_prog(index).DrawWidth = 1

    'si icone perso (seulement si c pas corbeille)
    icochemin = GetFromINI("config", "progico" & index, App.path & "\config.ini")
    If icochemin <> "" And File_Exist(icochemin) = True And pic_prog(index).Tag <> "special;::{645FF040-5081-101B-9F08-00AA002F954E}" Then 'icone perso
        taille = GetFromINI("config", "progicotaille" & index, App.path & "\config.ini")
        If Trim$(taille) = "" Then taille = "128"
      
      'pic_tmp_icoperso.Picture = LoadPicture(icochemin)
        mon_special_ico_quelquesoit_la_taille icochemin, Val(taille)
      'pic_tmp_icoperso.Picture = pic_tmp_icoperso.Image
      'pic_tmp_icoperso.Visible = True
      If Check_centrer.Value = 1 Then
          posx = ((pic_prog(index).Width / Screen.TwipsPerPixelX) - taille) / 2
          posy = ((pic_prog(index).Height / Screen.TwipsPerPixelY) - taille) / 2
      ElseIf Check_ico_text_horizontal.Value = 1 Then
          posx = 9
          posy = ((pic_prog(index).Height / Screen.TwipsPerPixelY) - taille) / 2
      Else
          posx = 5 '9
          posy = 5 '10
      End If
      'Call ImageList_Draw(hImgLarge&, ShInfo.iIcon, pic_prog(index).hdc, posx, posy, ILD_TRANSPARENT)
      Call PaintStandardPicture(pic_prog(index).hdc, pic_tmp_icoperso.Picture, posx, posy, Val(taille), Val(taille))
      If check_indice_folder.Value = 1 Then
        If File_or_folder(path) = "Folder" Then Call PaintStandardPicture(pic_prog(index).hdc, pic_folder.Picture, pic_prog(index).Width / Screen.TwipsPerPixelX - 20, 6, 16, 16)
      End If
      If pic_prog(index).WhatsThisHelpID = 1 Then 'selection
          'Call PaintStandardPicture(pic_prog(index).hdc, pic_sel.Picture, pic_prog(index).Width / Screen.TwipsPerPixelX - 20, pic_prog(index).Height / Screen.TwipsPerPixelY - 20, 16, 16)
            TransparentBlt pic_prog(index).hdc, pic_sel.hdc, (pic_prog(index).Width - pic_sel.Width) / Screen.TwipsPerPixelX, 1, pic_sel.Width / Screen.TwipsPerPixelX, pic_sel.Height / Screen.TwipsPerPixelY, vbCyan
            pic_prog(index).DrawWidth = 4
            pic_prog(index).Line (0, 0)-(pic_prog(index).Width - 10, pic_prog(index).Height - 10), 8880640, B
      End If
      'cadre
      'cadre mais pas de selection
      If Check_cadre.Value = 1 And pic_prog(index).WhatsThisHelpID = 0 Then pic_prog(index).Line (0, 0)-(pic_prog(index).Width - 10, pic_prog(index).Height - 10), BlendColor(pic_prog(index).BackColor, vbBlack, 150), B
  
      'rond  mais pas de selection
      If Check_roundcorner.Value = 1 And pic_prog(index).WhatsThisHelpID = 0 Then rond_pic index, Check_cadre.Value
    
      
      'spécial raccourcis (seulement un point au coin supérieur gauche
        'si raccourci donc ajouter une petite flèche pour destinguer les racc des autres fichiers
        If UCase(Right$(path, 3)) = "LNK" Then
            
            For i = 5 To 45 Step 5
                pic_prog(index).Circle ((70), (70)), i, BlendColor(pic_prog(index).BackColor, vbBlack, 150)
            Next i
        End If
      
    'End If

    Else 'récup ico depuis prog
    
    
        If Mid$(pic_prog(index).Tag, 1, Len("special")) = "special" Then
            If Check_centrer.Value = 1 Then
                posx = (pic_prog(index).Width - pic_system(0).Width) / 2
                posy = (pic_prog(index).Height - pic_system(0).Height) / 2
            ElseIf Check_ico_text_horizontal.Value = 1 Then
                posx = 60 '9
                posy = (pic_prog(index).Height - pic_system(0).Height) / 2
            Else
                posx = 45
                posy = 45
            End If
            
            If pic_prog(index).Tag = "special;::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then pic_prog(index).PaintPicture pic_system(0).Picture, posx, posy
            If pic_prog(index).Tag = "special;::{450D8FBA-AD25-11D0-98A8-0800361B1103}" Then pic_prog(index).PaintPicture pic_system(1).Picture, posx, posy
            If pic_prog(index).Tag = "special;::{208D2C60-3AEA-1069-A2D7-08002B30309D}" Then pic_prog(index).PaintPicture pic_system(2).Picture, posx, posy
            If pic_prog(index).Tag = "special;rundll32.exe shell32.dll,Control_RunDLL" Then pic_prog(index).PaintPicture pic_system(4).Picture, posx, posy
            'If pic_prog(index).Tag = "special;rundll32.exe shell32.dll,Control_RunDLL" Then Call PaintStandardPicture(pic_prog(index).hDC, pic_system(4).Picture, 45, 30, 32, 32)
            
            'cas spécial corbeille = pas de centrage
            'savoir si dessin corbeille vide ou pleine
            If pic_prog(index).Tag = "special;::{645FF040-5081-101B-9F08-00AA002F954E}" Then
                If corbeille_vide = True Then
                    pic_prog(index).PaintPicture pic_system(5).Picture, 45, 45
                Else
                    pic_prog(index).PaintPicture pic_system(3).Picture, 45, 45
                End If
            End If
            
            If pic_prog(index).WhatsThisHelpID = 1 Then 'selection
                'Call PaintStandardPicture(pic_prog(index).hdc, pic_sel.Picture, pic_prog(index).Width / Screen.TwipsPerPixelX - 20, pic_prog(index).Height / Screen.TwipsPerPixelY - 20, 16, 16)
                TransparentBlt pic_prog(index).hdc, pic_sel.hdc, (pic_prog(index).Width - pic_sel.Width) / Screen.TwipsPerPixelX, 1, pic_sel.Width / Screen.TwipsPerPixelX, pic_sel.Height / Screen.TwipsPerPixelY, vbCyan
                pic_prog(index).DrawWidth = 4
                pic_prog(index).Line (0, 0)-(pic_prog(index).Width - 10, pic_prog(index).Height - 10), 8880640, B
            End If
            
            
            'cadre
            If Check_cadre.Value = 1 And pic_prog(index).WhatsThisHelpID = 0 Then pic_prog(index).Line (0, 0)-(pic_prog(index).Width - 10, pic_prog(index).Height - 10), BlendColor(pic_prog(index).BackColor, vbBlack, 150), B
            'round
            If Check_roundcorner.Value = 1 And pic_prog(index).WhatsThisHelpID = 0 Then rond_pic index, Check_cadre.Value
            Exit Sub
        
        Else 'End If

    
            hImgLarge& = SHGetFileInfo(path, 0&, ShInfo, Len(ShInfo), _
            BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
            If Check_centrer.Value = 1 Then
                posx = (pic_prog(index).Width / Screen.TwipsPerPixelX - 32) / 2
                posy = (pic_prog(index).Height / Screen.TwipsPerPixelY - 32) / 2
            ElseIf Check_ico_text_horizontal.Value = 1 Then
                posx = 9
                posy = (pic_prog(index).Height / Screen.TwipsPerPixelY - 32) / 2
            Else
                posx = 9
                posy = 10
                'ImageList_Draw hImgLarge&, ShInfo.iIcon, pic_prog(index).hdc, 9, 10, ILD_TRANSPARENT
            End If
            ImageList_Draw hImgLarge&, ShInfo.iIcon, pic_prog(index).hdc, posx, posy, ILD_TRANSPARENT
            'si raccourci donc ajouter une petite flèche pour destinguer les racc des autres fichiers
            If UCase(Right$(path, 3)) = "LNK" Then
                If Check_centrer.Value = 1 Then
                    posx = (pic_prog(index).Width / Screen.TwipsPerPixelX - 32) / 2
                    posy = (pic_prog(index).Height / Screen.TwipsPerPixelY - 32) / 2 + 1
                ElseIf Check_ico_text_horizontal.Value = 1 Then
                    posx = 9
                    posy = (pic_prog(index).Height / Screen.TwipsPerPixelY - 32) / 2
                Else
                    posx = 10
                    posy = 12
                End If
                Call PaintStandardPicture(pic_prog(index).hdc, pic_shortcut.Picture, posx, posy, 32, 32)
            End If
            
            If check_indice_folder.Value = 1 Then
                If File_or_folder(path) = "Folder" Then Call PaintStandardPicture(pic_prog(index).hdc, pic_folder.Picture, pic_prog(index).Width / Screen.TwipsPerPixelX - 20, 6, 16, 16)
            End If
            
            If pic_prog(index).WhatsThisHelpID = 1 Then 'selection
                'Call PaintStandardPicture(pic_prog(index).hDC, pic_sel.Picture, pic_prog(index).Width / Screen.TwipsPerPixelX - 20, pic_prog(index).Height / Screen.TwipsPerPixelY - 20, 16, 16)
                TransparentBlt pic_prog(index).hdc, pic_sel.hdc, (pic_prog(index).Width - pic_sel.Width) / Screen.TwipsPerPixelX, 1, pic_sel.Width / Screen.TwipsPerPixelX, pic_sel.Height / Screen.TwipsPerPixelY, vbCyan
                pic_prog(index).DrawWidth = 4
                pic_prog(index).Line (0, 0)-(pic_prog(index).Width - 10, pic_prog(index).Height - 10), 8880640, B
                
            End If
            'cadre mais pas de selection
            If Check_cadre.Value = 1 And pic_prog(index).WhatsThisHelpID = 0 Then pic_prog(index).Line (0, 0)-(pic_prog(index).Width - 10, pic_prog(index).Height - 10), BlendColor(pic_prog(index).BackColor, vbBlack, 150), B
            'round mais pas de selection
            If Check_roundcorner.Value = 1 And pic_prog(index).WhatsThisHelpID = 0 Then rond_pic index, Check_cadre.Value
        End If
        
    End If
    
End Sub

Private Function File_or_folder(sFullPath As String) As String
    

    Dim wshShell As Object
    Dim wshLink As Object
    File_or_folder = ""
    
    Set wshShell = CreateObject("Scripting.FileSystemObject") 'CreateObject("WScript.FileSystemObject")
    'Set wshLink = wshShell.FileExists(strPath)
    File_or_folder = IIf(wshShell.FileExists(sFullPath), "File", "")
    
    'vérifie si c un dossier
    If File_or_folder = "" Then
        File_or_folder = IIf(wshShell.FolderExists(sFullPath), "Folder", "")
    End If
    'GetTarget = wshLink.TargetPath
    Set wshLink = Nothing
    Set wshShell = Nothing
    
    
End Function

Private Sub PaintStandardPicture(ByVal hDCDest As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    Optional ByVal xSrc As Long = 0, _
                                    Optional ByVal ySrc As Long = 0, _
                                    Optional ByVal hPal As Long = 0)
    
    If picSource Is Nothing Then Exit Sub
    
    Select Case picSource.Type
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            'Draw Icon onto DC
            DrawIconEx hDCDest, xDest, yDest, picSource.Handle, Width, Height, 0&, 0&, DI_NORMAL
    End Select
End Sub
Function mtag(index As Integer) As String

Dim Pos As Integer

Pos = InStr(1, pic_prog(index).Tag, ";")

If Pos <> 0 Then
    mtag = Mid$(pic_prog(index).Tag, 1, Pos - 1)
Else
    mtag = pic_prog(index).Tag
End If

mtag = replace_percent(mtag)

End Function

Function mcaption(index As Integer) As String

Dim Pos As Integer

Pos = InStr(1, pic_prog(index).Tag, ";")

If Pos <> 0 Then
    mcaption = Mid$(pic_prog(index).Tag, Pos + 1, Len(pic_prog(index).Tag))
    mcaption = Trim$(mcaption)
Else
    mcaption = ExtractFileName(pic_prog(index).Tag)
End If

End Function
Function nbre_total_prog() As Integer

Dim i As Integer
nbre_total_prog = 0
For i = 1 To pic_prog.UBound
    If pic_prog(i).Visible Then nbre_total_prog = nbre_total_prog + 1
Next i

End Function

Sub save_param()



    Dim j As Integer
        
    j = 0
    For i = 0 To pic_prog.UBound
        If pic_prog(i).Visible = True Then j = j + 1
    Next i
    Call WriteToINI("config", "nbreprog", j - 1 & "", App.path & "\config.ini")
    
    
    
    If Check_centrer.Value = 1 Then
        Call WriteToINI("config", "style", 1, App.path & "\config.ini")
    ElseIf Check_ico_text_horizontal.Value = 1 Then
        Call WriteToINI("config", "style", 2, App.path & "\config.ini")
    Else
        Call WriteToINI("config", "style", 0, App.path & "\config.ini")
    End If
    
    Call WriteToINI("config", "indiceFolder", check_indice_folder.Value, App.path & "\config.ini")
    Call WriteToINI("config", "desktopicohide", Check_hide_desktop_icon.Value, App.path & "\config.ini")
    Call WriteToINI("config", "affichecadre", Check_cadre.Value, App.path & "\config.ini")
    'Check_cadre.Value = GetFromINI("config", "affichecadre", App.path & "\config.ini")
    Call WriteToINI("config", "showcaptionstatusbar", Check_caption_statusbar.Value, App.path & "\config.ini")
    
    Call WriteToINI("config", "formleft", Me.Left, App.path & "\config.ini")
    Call WriteToINI("config", "formtop", Me.Top, App.path & "\config.ini")
    Call WriteToINI("config", "formwidth", Me.Width, App.path & "\config.ini")
    Call WriteToINI("config", "formheight", Me.Height, App.path & "\config.ini")
    Call WriteToINI("config", "formcolor", Me.BackColor, App.path & "\config.ini")
    
    Call WriteToINI("config", "simule3d", simule_3d.Value, App.path & "\config.ini")
    Call WriteToINI("config", "horlogecolor", Label_horloge1.ForeColor, App.path & "\config.ini")
    
    Call WriteToINI("config", "changebackcolormousemove", check_change_backcolor_mousemove.Value, App.path & "\config.ini")
    Call WriteToINI("config", "gradient", Check_gradient.Value, App.path & "\config.ini")
    Call WriteToINI("config", "roundcorner", Check_roundcorner.Value, App.path & "\config.ini")
    
    
    'Call WriteToINI("config", "cadrecolor", Shape_sel.BorderColor, App.path & "\config.ini")
    
    
    j = 0
    For i = 0 To pic_prog.UBound
        If pic_prog(i).Visible = True Then
            'If i = 29 Or i = 24 Then Stop
            'Debug.Print pic_prog(i).Tag
            
            Call WriteToINI("config", "progleft" & j, pic_prog(i).Left, App.path & "\config.ini")
            Call WriteToINI("config", "progtop" & j, pic_prog(i).Top, App.path & "\config.ini")
            Call WriteToINI("config", "progwidth" & j, pic_prog(i).Width, App.path & "\config.ini")
            Call WriteToINI("config", "progheight" & j, pic_prog(i).Height, App.path & "\config.ini")
            Call WriteToINI("config", "progcolor" & j, pic_prog(i).BackColor, App.path & "\config.ini")
            Call WriteToINI("config", "progtag" & j, pic_prog(i).Tag, App.path & "\config.ini")
            Call WriteToINI("config", "progtextcolor" & j, pic_prog(i).ForeColor, App.path & "\config.ini")
            j = j + 1
        End If
    Next i
    
    


End Sub

Sub create_ini()

    Dim j As Integer
        
    Call WriteToINI("config", "nbreprog", 2 & "", App.path & "\config.ini")

    Call WriteToINI("config", "formleft", "3000", App.path & "\config.ini")
    Call WriteToINI("config", "formtop", "1000", App.path & "\config.ini")
    Call WriteToINI("config", "formwidth", "8945", App.path & "\config.ini")
    Call WriteToINI("config", "formheight", "6215", App.path & "\config.ini")
    Call WriteToINI("config", "formcolor", "8421376", App.path & "\config.ini")
    
    Call WriteToINI("config", "progleft0", "120", App.path & "\config.ini")
    Call WriteToINI("config", "progtop0", "120", App.path & "\config.ini")
    Call WriteToINI("config", "progwidth0", "2400", App.path & "\config.ini")
    Call WriteToINI("config", "progheight0", "960", App.path & "\config.ini")
    Call WriteToINI("config", "progcolor0", "1163880", App.path & "\config.ini")
    Call WriteToINI("config", "progtag0", "%SystemRoot%\explorer.exe;Explorateur Windows", App.path & "\config.ini")

    Call WriteToINI("config", "progleft1", "120", App.path & "\config.ini")
    Call WriteToINI("config", "progtop1", "1095", App.path & "\config.ini")
    Call WriteToINI("config", "progwidth1", "2400", App.path & "\config.ini")
    Call WriteToINI("config", "progheight1", "960", App.path & "\config.ini")
    Call WriteToINI("config", "progcolor1", "2187751", App.path & "\config.ini")
    Call WriteToINI("config", "progtag1", "%SystemRoot%\notepad.exe;Bloc-Notes", App.path & "\config.ini")

    Call WriteToINI("config", "progleft2", "2540", App.path & "\config.ini")
    Call WriteToINI("config", "progtop2", "120", App.path & "\config.ini")
    Call WriteToINI("config", "progwidth2", "2400", App.path & "\config.ini")
    Call WriteToINI("config", "progheight2", "960", App.path & "\config.ini")
    Call WriteToINI("config", "progcolor2", "47321", App.path & "\config.ini")
    Call WriteToINI("config", "progtag2", "%SystemRoot%\system32\mspaint.exe;Paint", App.path & "\config.ini")

    'Explorateur Windows,%SystemRoot%\explorer.exe



End Sub

Sub savewallpaper(f As String)

If File_Exist(f) = True Then
    Call WriteToINI("config", "wallpaper", f, App.path & "\config.ini")
End If

End Sub

Private Function GetRGB(R As Integer, g As Integer, B As Integer, Color As Long)
    Dim R1, G1, B1
    Dim R2, G2, B2
    Dim a As Long
    
    'Variables:
    ' Output = R as Integer, G as Integer, B as Integer
    ' Input  = color as Long
    
    
    
    
    
    R2 = 0      'Reset variables
    G2 = 0
    B2 = 0
    
    If Color = -1 Then GoTo ExitGetRGB      'If not a color, exit
    
    For B1 = 0 To 255           'Loop to retrieve Blue
        DoEvents
        a = RGB(R2, G2, B1)
        If a > Color Then
            B2 = (B1 - 1)
            Exit For
        ElseIf a = Color Then
            B2 = B1
            Exit For
        End If
    Next B1
    
    For G1 = 0 To 255           'Loop to retrieve Green
        DoEvents
        a = RGB(R2, G1, B2)
        If a > Color Then
            G2 = (G1 - 1)
            Exit For
        ElseIf a = Color Then
            G2 = G1
            Exit For
        End If
    Next G1
    
    For R1 = 0 To 255           'Loop to retrieve Red
        DoEvents
        a = RGB(R1, G2, B2)
        If a = Color Then
            R2 = (R1)
            Exit For
        End If
    Next R1
    
ExitGetRGB:
    
    R = R2      'Returns the Red into the variable assigned
    g = G2      'Returns the Green into the variable assigned
    B = B2      'Returns the Blue into the variable assigned
End Function


Sub traitement_special(index As Integer)
Dim txt As String
'
'/n = explorateur
'/e = extended folder (liste des dossiers a gauche, et fichiers a droite)
'corbeille
'Shell "explorer /n,::{645FF040-5081-101B-9F08-00AA002F954E}", vbNormalFocus
'desktop
'Shell "explorer.exe /n, /root", vbNormalFocus
'ordnateur
'::{20D04FE0-3AEA-1069-A2D8-08002B30309D}
'my doc
'::{450D8FBA-AD25-11D0-98A8-0800361B1103}
'My Network Places (default view)
'            sFile = "explorer.exe"
'            sParams = "::{208D2C60-3AEA-1069-A2D7-08002B30309D}"
'Printers & Faxes (default view)
'            sFile = "explorer.exe"
'            sParams = "::{2227A280-3AEA-1069-A2DE-08002B30309D}"

        pic_prog(index).Cls
        
        If Check_gradient.Value = 1 Then Call PaintGradient(pic_prog(index).hdc, 0, 0, pic_prog(index).ScaleWidth / 15, pic_prog(index).ScaleHeight / 15, BlendColor(pic_prog(index).BackColor, vbWhite, Slider1.Value), pic_prog(index).BackColor, 132)  '150
        
        
        'txt = pic_prog(Index).ToolTipText
        If pic_prog(index).Tag = "special;::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then txt = "Poste de Travail"
        If pic_prog(index).Tag = "special;::{450D8FBA-AD25-11D0-98A8-0800361B1103}" Then txt = "Mes documents"
        If pic_prog(index).Tag = "special;::{208D2C60-3AEA-1069-A2D7-08002B30309D}" Then txt = "Connexions réseau"
        If pic_prog(index).Tag = "special;::{645FF040-5081-101B-9F08-00AA002F954E}" Then txt = "Corbeille"
        If pic_prog(index).Tag = "special;rundll32.exe shell32.dll,Control_RunDLL" Then txt = "Panneau de Configuration" '"Imprimantes et télécopieurs"
        
        'txt=
        
        actualise_caption index, txt, 770, True
        
'        pic_prog(index).CurrentX = 100: pic_prog(index).CurrentY = 770
'
'        If Check_centrer.Value = 1 Then pic_prog(index).CurrentX = (pic_prog(index).Width - TextWidth(txt)) / 2
'
'        pic_prog(index).Print txt
         
        DrawIcon mtag(index), index, 0, False
        pic_prog(index).Refresh


End Sub

Private Sub Check_cadre_Click()
Check_centrer_Click
End Sub

Private Sub Check_centrer_Click()

If Check_centrer.Value = 1 Then Check_ico_text_horizontal.Value = 0

Dim i As Integer
For i = 0 To pic_prog.UBound
    actualise_ico_caption i, ""
Next i

info_corbeille = "" 'pour forcer l'actualisation de l'info corbeille

End Sub


Private Sub Check_editable_Click()

If Check_editable.Value = 1 Then
    Frame_mode_edition.Visible = True: Frame_couleur.Visible = False
    DrawGrid
    Frame_mode_edition.Visible = True
    cmd_cree_button.Visible = True
Else
    Frame_mode_edition.Visible = False: Frame_couleur.Visible = True
    ctl.HideHandles
    
    info_corbeille = "" 'pour forcer le timer
    
    Me.Cls
    
'        If Dir$(App.path & "\papierpaint.bmp") <> "" Then
'            Me.Picture = LoadPicture(App.path & "\papierpaint.bmp")
'        ElseIf Dir$(App.path & "\papierpaint.jpg") <> "" Then
'            Me.Picture = LoadPicture(App.path & "\papierpaint.jpg")
'        Else
'            Me.Picture = Nothing
'        End If
    
    Frame_mode_edition.Visible = False
    cmd_cree_button.Visible = False
    pic_barre.ZOrder 0
End If
End Sub

Private Sub Check_gradient_Click()

Dim i As Integer
For i = 0 To pic_prog.UBound
    actualise_ico_caption i, ""
Next i

info_corbeille = "" 'pour forcer l'actualisation de l'info corbeille


End Sub

Private Sub Check_hide_desktop_icon_Click()

If Check_hide_desktop_icon.Value = 0 Then
    DesktopIconsShow
Else
    DesktopIconsHide
End If

End Sub

Function DesktopIconsShow()
    Dim hwnd As Long
    hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hwnd, 5
End Function
Function DesktopIconsHide()
    Dim hwnd As Long
    hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hwnd, 0
End Function

Private Sub Check_ico_text_horizontal_Click()
If Check_ico_text_horizontal.Value = 1 Then Check_centrer.Value = 0
Check_centrer_Click
End Sub

Private Sub check_indice_folder_Click()

Dim i As Integer
For i = 0 To pic_prog.UBound
    actualise_ico_caption i, ""
Next i


End Sub

Private Sub Check_roundcorner_Click()

Dim i As Integer
For i = 0 To pic_prog.UBound
    actualise_ico_caption i, ""
Next i

info_corbeille = "" 'pour forcer l'actualisation de l'info corbeille

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Frame_alignement.Visible = False
Else
    Frame_alignement.Visible = True
End If
End Sub

Private Sub Check2_Click()
Frame_magnetisme.Visible = Not Frame_magnetisme.Visible
End Sub

Private Sub cmd_recharger_param_Click()

Dim nbre As Integer, R As Integer

R = MsgBox("êtes-vous sûr de vouloir réstaurer la dernière Version?", vbYesNo + vbInformation, "Restauration")
If R = 7 Then Exit Sub
    
    
    
    nbre = GetFromINI("config", "nbreprog", App.path & "\config.ini")
    If nbre < nbre_total_prog Then
        R = MsgBox("Vous avez ajouter des programmes, si vous continuer, vous allez les perdre. continuer?", vbYesNo + vbCritical, "Attention")
        If R = 7 Then Exit Sub
    End If

Screen.MousePointer = 13
charge_param
ctl.HideHandles
Screen.MousePointer = 0
End Sub


Private Sub cmd_cree_button_Click()

    Load pic_prog(pic_prog.UBound + 1)
    pic_prog(pic_prog.UBound).Visible = True
    pic_prog(pic_prog.UBound).BackColor = RGB(NombreAleatoire(1, 255), NombreAleatoire(1, 255), NombreAleatoire(1, 255))
    pic_prog(pic_prog.UBound).Tag = ""
    pic_prog(pic_prog.UBound).Width = 1500: pic_prog(pic_prog.UBound).Height = 1000
    
    'si icone system (poste travail ...), respecter la position dragdrop
    If dragdropico = True Then
        pic_prog(pic_prog.UBound).Left = xlocation - (1500 / 2)
        pic_prog(pic_prog.UBound).Top = ylocation - (1500 / 2)
    Else
        pic_prog(pic_prog.UBound).Left = (Me.Width - pic_prog(pic_prog.UBound).Width) / 2
        pic_prog(pic_prog.UBound).Top = (Me.Height - pic_prog(pic_prog.UBound).Height) / 2
    End If
    
    pic_prog(pic_prog.UBound).ZOrder 0
    pic_prog(pic_prog.UBound).Cls
    pic_prog(pic_prog.UBound).Picture = Nothing
    pic_prog_MouseDown pic_prog.UBound, 1, 0, 0, 0  ' (pic_prog.UBound)
    
    charger_onglet 1
    pic_frame_config.Visible = True
    
    
End Sub

Private Sub cmd_save_config_Click()
Screen.MousePointer = 13
Call save_param
Screen.MousePointer = 0
End Sub

Private Sub Combo_dim_Click()

On Error Resume Next
Combo_dim_destination.ListIndex = Combo_dim.ListIndex

If Combo_dim = "Top+Height" Then Combo_dim_destination.ListIndex = 1
If Combo_dim = "Left+Width" Then Combo_dim_destination.ListIndex = 0

End Sub

Private Sub Command1_Click()
Dim i As Integer
For i = 0 To pic_prog.UBound
    pic_prog(i).Visible = True
Next i
End Sub

Private Sub cmd_deplacbottom_Click()
'pic_prog(indexencours).Top = pic_prog(indexencours).Top + 10
deplace_resize indexencours, 1, "10"


End Sub

Private Sub cmd_deplacleft_Click()
deplace_resize indexencours, 0, "-10"
'pic_prog(indexencours).Left = pic_prog(indexencours).Left - 10

End Sub



Private Sub cmd_deplacright_Click()
'pic_prog(indexencours).Left = pic_prog(indexencours).Left + 10
deplace_resize indexencours, 0, "+10"

End Sub

Private Sub cmd_deplactop_Click()
'pic_prog(indexencours).Top = pic_prog(indexencours).Top - 10
deplace_resize indexencours, 1, "-10"

End Sub


Private Sub cmd_height_Click()
'pic_prog(indexencours).Height = pic_prog(indexencours).Height + 10
deplace_resize indexencours, 3, "10"


End Sub

Private Sub cmd_heightminus_Click()
'pic_prog(indexencours).Height = pic_prog(indexencours).Height - 10
deplace_resize indexencours, 3, "-10"

End Sub

Private Sub cmd_width_Click()
'pic_prog(indexencours).Width = pic_prog(indexencours).Width + 10
deplace_resize indexencours, 2, "10"


End Sub

Private Sub cmd_widthminus_Click()
'pic_prog(indexencours).Width = pic_prog(indexencours).Width - 10
deplace_resize indexencours, 2, "-10"

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub


Private Sub Command7_Click()

End Sub


Private Sub Command8_Click()

If combo_valeur_espace = "Petit Espace" Then
    alignement_height_ou_width 15, "top+height"
Else
    alignement_height_ou_width 60, "top+height"
End If




    
End Sub

Private Sub Command9_Click()

If combo_valeur_espace = "Petit Espace" Then
    alignement_height_ou_width 15, "left+width"
Else
    alignement_height_ou_width 60, "left+width"
End If

End Sub


Private Sub Form_DblClick()
Dim i As Integer
    
    If Check_editable.Value = 1 Then Exit Sub
    
    
    
    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Me.hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Images JPG (*.jpg)" + Chr$(0) + "*.jpg" + Chr$(0) + "Images BMP (*.bmp)" + Chr$(0) + "*.bmp" + Chr$(0) + "Tous les Fichiers (*.*)" + Chr$(0) + "*.*"
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    'OFName.lpstrInitialDir = "C:\"
    OFName.lpstrTitle = "Sélectionnez une Image"
    OFName.flags = 0

    If GetOpenFileName(OFName) Then
        savewallpaper Trim$(OFName.lpstrFile) 'CommonDialog1.FileName
        loadwallpeper
        'simule_trans_panel True 'actialiser l'arrière plan des pic_conteneur
        If Check_roundcorner.Value = 1 Then
            For i = 0 To pic_prog.UBound
                If pic_prog(i).Visible Then rond_pic i, Check_cadre.Value
            Next i
        End If
    Else
        'cancel
    End If
    

End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
Dim i As Integer

If Source.Name = "Shape_pick_color" Then
    If Me.Picture <> 0 Then
        If MsgBox("Remplacer l'image d'arrière plan par cette couleur?", vbYesNo + vbInformation, "Couleur de fond") = vbYes Then
            Me.Picture = Nothing
            Me.BackColor = Shape_pick_color.BackColor
            'If Check_editable.Value = 1 Then DrawGrid
            savewallpaper "" ' la couleur remplace limage
            If Check_roundcorner.Value = 1 Then
                For i = 0 To pic_prog.UBound
                    If pic_prog(i).Visible Then rond_pic i, Check_cadre.Value
                Next i
            End If
        End If
    Else
        Me.BackColor = Shape_pick_color.BackColor
        If Check_roundcorner.Value = 1 Then
            For i = 0 To pic_prog.UBound
                If pic_prog(i).Visible Then rond_pic i, Check_cadre.Value
            Next i
        End If
        'If Check_editable.Value = 1 Then DrawGrid
        'savewallpaper "" ' la couleur remplace limage
    End If
    'mon_freecolor = cmd_free_lecteur.BackColor
'End If

ElseIf Source.Name = "pic_system" Then
    
    'Call pic_barre_Click
    'pic_onglet_Click 1
    dragdropico = True: xlocation = X: ylocation = Y
    Call cmd_cree_button_Click
    dragdropico = False
    'pic_prog_OLEDragDrop pic_prog.UBound, Data, Effect, Button, Shift, X, Y
    indexencours = pic_prog.Count - 1
    pic_prog(indexencours).CurrentX = 150: pic_prog(indexencours).CurrentY = 650
    'pic_prog(index).Print Tempo
    pic_prog(indexencours).Tag = Source.Tag
    'dessine_icone Tempo, Index
    actualise_ico_caption indexencours, ""
    
    pic_onglet_Click 1
    pic_prog_MouseDown indexencours, 1, 0, 0, 0
        
    
    'Me.Picture = Nothing
    'Me.BackColor = Shape_pick_color.BackColor
    'If Check_editable.Value = 1 Then DrawGrid
    ''mon_freecolor = cmd_free_lecteur.BackColor
'End If
'ElseIf Source.Name = "pic_prog" Then 'And indexencours <> -1 Then
'    Set Source.Container = Me 'pic_conteneur(Index)
'    Source.Left = x
'    Source.Top = y
End If




End Sub


Private Sub Form_Initialize()
InitCommonControls

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If indexencours = -1 And indexencours_conteneur = -1 Then Exit Sub 'aucun objet sélectionné

'if indexencours<>

'MsgBox KeyCode
If Check_editable.Value = 1 Then
    
    If Shift = 1 Then 'shift = resize
        If KeyCode = 37 Then cmd_widthminus_Click
        If KeyCode = 38 Then cmd_heightminus_Click
        If KeyCode = 39 Then cmd_width_Click
        If KeyCode = 40 Then cmd_height_Click
    ElseIf Shift = 2 Then 'ctrl = move
        If KeyCode = 37 Then cmd_deplacleft_Click
        If KeyCode = 38 Then cmd_deplactop_Click
        If KeyCode = 39 Then cmd_deplacright_Click
        If KeyCode = 40 Then cmd_deplacbottom_Click
    End If
    
''    'suppr
''        If KeyCode = 46 Then
''            show_number
''            Dim i As Integer
''            'appliquer sur tous les objets sélectionnés
''            For i = 0 To pic_prog.UBound
''                If pic_prog(i).WhatsThisHelpID = 1 Or i = indexencours Then 'déjà sel
''                    pic_prog(i).Visible = False: ctl.HideHandles
''                End If
''            Next i
''
''            Call suppr_panel
''
''        End If

Else
'    If KeyCode = 114 Then 'F3 -> search
'        'Picture5_Click
'        SetCursorPos txt_search.Width / 2, txt_search.Height / 2
'        txt_search.SetFocus
'        If Trim$(txt_search) <> "" Then Picture5_Click
'    End If
    If Shift = 4 Then 'pour lancer un prog avec les touche d'accès rapide "Alt" + caractère souligné
        'If KeyCode <> 18 Then MsgBox KeyCode & " - " & Chr(KeyCode)
        If KeyCode <> 18 Then key_lancer_prog Chr(KeyCode) '18 = touche alt
    End If
    
    If KeyCode = 112 Then frm_astuce.Show 1, Me 'F1
    
End If

End Sub


Private Sub Form_Load()

Screen.MousePointer = vbHourglass

Set ctl = New CControlSizer
ctl.GridSize = 8
ctl.AttachForm Me 'The form that is using the designer class

'DrawGrid 'Draw the GridSize by GridSize pixel grid on the background of the form


info_corbeille = ""
corbeille_vide = True

charge_param

verifi_pos_screen 'ou cas ou

Shape1.Width = pic_frame_config.Width - 5
Shape1.Height = pic_frame_config.Height - 5

label0.BackStyle = 0
pic_frame_config.Top = -1

'
'pic_frame_config.Visible = True
'Frame_mode_edition.Container = pic_frame_config
'Frame_mode_edition.Left = Frame_couleur.Left
'Frame_mode_edition.Top = Frame_couleur.Top
'pic_frame_config.Visible = False

Call initialise_combobox

Dim i As Integer
For i = 0 To 4
    'pic_system(0).Cls
    pic_system(i).Line (0, 0)-(pic_system(i).Width - 10, pic_system(i).Height - 10), pic_system(i).ForeColor, B
Next i

Frame_param.Visible = True
Frame_aspect.Visible = False
charger_onglet 0
charger_onglet2 3

barre_visible = True

combo_valeur_espace.AddItem "Petit Espace"
combo_valeur_espace.AddItem "Espace Moyen"
combo_valeur_espace.ListIndex = 0

combo_uniformiser_taille.AddItem "Petit"
combo_uniformiser_taille.AddItem "Grand"
combo_uniformiser_taille.ListIndex = 0

oldindex = -1


horloge

Screen.MousePointer = vbDefault

''indexencours_conteneur = -1

''simule_trans_panel True

''MsgBox "arret sur : ctrl+click pour desel l'élement en cours (indexencours) afin de ne plus le prendre en considération"

End Sub


Sub initialise_combobox()

Combo_dim.AddItem "Left"
Combo_dim.AddItem "Top"
Combo_dim.AddItem "Width"
Combo_dim.AddItem "Height"
Combo_dim.AddItem "Bordure_droite"
Combo_dim.AddItem "Bordure_inférieure"

Combo_dim.AddItem "Top+Height"
Combo_dim.AddItem "Left+Width"

Combo_dim_destination.AddItem "Left"
Combo_dim_destination.AddItem "Top"
Combo_dim_destination.AddItem "Width"
Combo_dim_destination.AddItem "Height"
Combo_dim_destination.AddItem "Bordure_droite"
Combo_dim_destination.AddItem "Bordure_inférieure"

Combo_dim.ListIndex = 0
Combo_dim_destination.ListIndex = 0

Combo_color.AddItem "Fond": Combo_color.AddItem "Texte"
Combo_color.ListIndex = 0


End Sub


Public Sub DrawGrid()

'Draws black pixels, GridSize pixels apart on both the X and Y axis.
'This is done in the Form_Load event BEFORE the form is shown -
'if we waited until the form is visible to do this, it would
'take too much time.
'SetPixelV is used because it's faster than a PSet method.

'Exit Sub

Dim X As Long
Dim Y As Long

Me.AutoRedraw = True

'Draw on the form's memory image
For Y = 0 To (Screen.Height \ Screen.TwipsPerPixelY) Step ctl.GridSize
    For X = 0 To (Screen.Width \ Screen.TwipsPerPixelX) Step ctl.GridSize
        SetPixelV Me.hdc, X, Y, vbBlack
    Next
Next
'Set the form's Picture property to what we've drawn -
'a grid that's the size of the screen.

'Set Me.Picture = Me.Image
Me.Refresh

'Me.AutoRedraw = False

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

indexencours = -1 '': indexencours_conteneur = -1
Text_name.Visible = False
'List_taille_ico.Visible = False

If Check_editable.Value = 1 Then
    ctl.HideHandles
    'Call Label_desell_all_Click
    StartX = X
    StartY = Y
    If Button = 1 Then
        blnMouseIsDown = True
        'ClearSelection
        blnControlSelected = False
        Exit Sub
    End If
End If

If Button = 1 Then
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    verifi_pos_screen
ElseIf Button = 2 Then
    'Call pic_barre_Click
    
    If pic_frame_config.Visible = True Then
        label0_Click
        Exit Sub
    End If
    
    'simule_trans_panel False 'pour voir les panels (pic_conteneur)

    If pic_frame_config.Top = -1 Then 'pour exécuter ce code une seule fois
        pic_frame_config.Left = Me.Width - pic_frame_config.Width - 60
        pic_frame_config.Top = 200 '60
    End If
    
    'pic_frame_config.Visible = Not
    
    If pic_frame_config.Left + pic_frame_config.Width > Me.Width - 60 Then pic_frame_config.Left = Me.Width - pic_frame_config.Width - 60
    If pic_frame_config.Top + pic_frame_config.Height > Me.Height Then pic_frame_config.Top = 200
    
    'If pic_frame_config.Visible = True Then Exit Sub
    
    pic_frame_config.Visible = True
    Call charger_onglet(0)
    
    
End If

End Sub


Private Function NombreAleatoire(ByVal lngInf As Long, ByVal lngSup As Long) As Long


Randomize 'initialise le générateur pseudo-aléatoire
NombreAleatoire = Int(Rnd() * (lngSup - lngInf + 1)) + lngInf

End Function

Sub verifi_pos_screen()
        

    'vérifier avec le vrai escpace libre de l'écran et non pas seulement screen.width et screen.height
    Dim lNewTop As Long, lNewLeft As Long
    Dim WA As RECT, lReturn As Long

    lReturn = SystemParametersInfo(SPI_GETWORKAREA, 0&, WA, 0&)
    
    WA.Left = WA.Left * Screen.TwipsPerPixelX
    WA.Right = WA.Right * Screen.TwipsPerPixelX
    WA.Top = WA.Top * Screen.TwipsPerPixelY
    WA.Bottom = WA.Bottom * Screen.TwipsPerPixelY
    
    If Me.Left < WA.Left Then Me.Left = WA.Left
    If Me.Top < WA.Top Then Me.Top = WA.Top
    If (Me.Top + Me.Height) > WA.Bottom - WA.Top Then Me.Top = WA.Bottom - WA.Top - Me.Height
    If (Me.Left + Me.Width) > WA.Right - WA.Left Then Me.Left = WA.Right - WA.Left - Me.Width
    
    


End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Debug.Print GetPixel(hdc, X, Y)

If Check_editable.Value = 1 Then
    If blnMouseIsDown Then
        If Button = 1 Then
            ExecuteState 0, 0, X, Y
        'ElseIf Button = 2 Then
        '    Exit Sub
        End If
    'Else
    End If
End If

pic_close_sur_mousemove.Visible = False

'si changement de backcolor sur mousemove
'If check_change_backcolor_mousemove.Value = 1 And Check_editable.Value = 0 Then
If Check_editable.Value = 0 Then
    If check_change_backcolor_mousemove.Value = 1 Then
        'restaurer color ancien
        If oldindex <> -1 Then
            pic_prog(oldindex).BackColor = old_backcolor
            actualise_ico_caption oldindex
            'si corbeille forcer refresh info
            If mcaption(oldindex) = "::{645FF040-5081-101B-9F08-00AA002F954E}" Then
                info_corbeille = "": Timer1_Timer
            End If
            oldindex = -1
        End If
    Else
        If oldindex <> -1 Then
            'pic_prog(oldindex).BackColor = old_backcolor
            actualise_ico_caption oldindex
            'si corbeille forcer refresh info
            If mcaption(oldindex) = "::{645FF040-5081-101B-9F08-00AA002F954E}" Then
                info_corbeille = "": Timer1_Timer
            End If
            oldindex = -1
        End If
    End If
End If

oldindex = -1

Shape_sel.Visible = False: Shape_sel2.Visible = False
'label_caption = "{Fenêtre} -> DblClick = Choix Papier Paint  | Click Droit = Afficher/Cacher la Boîte à outils | {Programme} -> Click Droit = Changer Nom | Ctrl+Click Droit = Taille Icône Perso"

label_caption = "F1 = Astuce"


End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

blnMouseIsDown = False

    If SelectBox.Visible Then
        SelectObjects 0, Shift 'Index
        SelectBox.Visible = False
    End If
    'lngState = Default


End Sub


Private Sub SelectObjects(GetSec As Integer, Optional ctrl As Integer)
Dim i As Integer, k As Integer
On Error Resume Next
    NumInGrp = -1
    If ctrl <> 2 And ctrl <> 1 Then ClearSelection 'si touche ctrl ou shift alors garder la selection
    'ClearSelection
    With SelectBox
        For i = 0 To pic_prog.UBound
            'Set ctlTest = Me.Controls(i)
                    'ctlTest.Enabled = True
                    'actualise_ico_caption i
                    If .Left < pic_prog(i).Left And _
                    .Left + .Width > pic_prog(i).Left + pic_prog(i).Width Then
                        If .Top < pic_prog(i).Top And _
                        .Top + .Height > pic_prog(i).Top + pic_prog(i).Height Then
                            'blnGroupSelected = True
                            'NumInGrp = NumInGrp + 1
                            'ReDim Preserve SelectedCtl(NumInGrp)
                            'Set SelectedCtl(NumInGrp).ctl = ctlTest
                            'ctlTest.Enabled = False
                            'pic_prog(i).WhatsThisHelpID = 1
                            If pic_prog(i).WhatsThisHelpID = 1 Then 'enlevé la sélection
                                pic_prog(i).WhatsThisHelpID = 0
                                actualise_ico_caption i
                            Else 'ajouter à la sélection
                                pic_prog(i).WhatsThisHelpID = 1
                                actualise_ico_caption i
                                pic_prog(i).ZOrder 0
                                ctl.AttachControl pic_prog(i)
                                indexencours = i
                            End If
                        End If
                    End If
            'blnSelectArrayInit = True
        Next i
    End With
End Sub
Sub ClearSelection()


    ctl.HideHandles
    Call Label_desell_all_Click

End Sub

Private Sub ExecuteState(SectionIndex As Integer, GetState As Long, GetX As Single, GetY As Single)
Dim i As Integer, j As Integer

    ReSizeSelectBox SectionIndex, GetX, GetY

End Sub



Private Sub ReSizeSelectBox(GetSec As Integer, MousX As Single, MousY As Single)
'sizes the selection box based on the mouse coordinates
Dim i As Integer, k As Integer

    'If blnGridOn Then
    '    ShowGrid
    'Else
    '    picSection(GetSec).Cls
    'End If
    
    If Not blnDragStarted Then
        SelectBox.ZOrder (0)
        blnDragStarted = True

    End If
    
    blnGroupSelected = False
    NumInGrp = -1
    
    With SelectBox
        If MousX >= StartX Then
            .Left = StartX
            .Width = MousX - StartX
        Else
            .Left = MousX
            .Width = StartX - MousX
        End If
        
        If MousY >= StartY Then
            .Top = StartY
            .Height = MousY - StartY
        Else
            .Top = MousY
            .Height = StartY - MousY
        End If
        .Visible = True
        
    End With
        
End Sub



Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

pic_onglet_Click 1

dragdropico = True: xlocation = X: ylocation = Y
Call cmd_cree_button_Click
dragdropico = False

pic_prog_OLEDragDrop pic_prog.UBound, Data, Effect, Button, Shift, X, Y

pic_prog_MouseDown pic_prog.UBound, 1, 0, 0, 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Me.Height = 300 Then Call label_min_MouseUp(2, 0, 0, 0)


save_param

End Sub

Private Sub Form_Resize()
On Error Resume Next

If Me.WindowState = 1 Then Exit Sub

If Me.Height <= 300 Then Exit Sub

If Me.Width < 5200 Then Me.Width = 5200
If Me.Height < 2800 Then Me.Height = 2800


'Me.Show
pic_barre.Left = 0
pic_barre.Top = Height - pic_barre.Height '- 5 '- 20 ' - pic_resize.Height
pic_barre.Width = Width


pic_resize.Left = Width - pic_resize.Width - 60 '- pic_resize.Width
pic_resize.Top = 180 'pic_resize.Height  '- 10 ' - pic_resize.Height

pic_show_desktop.Left = pic_resize.Left - pic_show_desktop.Width - 60 '- pic_resize.Width
pic_show_desktop.Top = 120 '30 'pic_resize.Height  '- 10 ' - pic_resize.Height


pic_min_close.Top = 15
pic_min_close.Left = Me.Width - pic_min_close.Width - 30

'Me.Refresh

If pic_frame_config.Visible Then
    If pic_frame_config.Left + pic_frame_config.Width > Me.Width - 60 Then pic_frame_config.Left = Me.Width - pic_frame_config.Width - 60
    If pic_frame_config.Top + pic_frame_config.Height > Me.Height Then pic_frame_config.Top = 200

    'If pic_frame_config.Left + pic_frame_config.Width < Me.Width - 60 Then pic_frame_config.Left = Me.Width - pic_frame_config.Width - 60
End If

Shape_form.Move 0, 0, Me.Width, Me.Height



Label_horloge1.Top = Me.Height - Label_horloge1.Height - 900 'Label_horloge2.Height - 600 - Label_horloge1.Height - 20
Label_horloge1.Left = Me.Width - Label_horloge1.Width - 400 '+ 80

Label_horloge2.Move Label_horloge1.Left - ((Label_horloge2.Width - Label_horloge1.Width) / 2), Label_horloge1.Top + Label_horloge1.Height - 100 '+ 5


End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, j As Integer
If Button = 1 Then
    If Image1.MousePointer = 0 Then
        Image1.MousePointer = 99
        Set Image1.Picture = Nothing
    End If
    Image1.Refresh
    Shape_pick_color.BackColor = GetColorAtCursor
    

End If


End Sub


Private Function GetColorAtCursor() As Long
Dim p As POINTAPI, C As Long
    GetCursorPos p
    C = GetPixel(GetWindowDC(GetDesktopWindow), p.X, p.Y)
    GetColorAtCursor = C
End Function

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If Image1.MousePointer = 99 Then
        Image1.MousePointer = 0
        Set Image1.Picture = imgcircle.Picture
        'Shape_pick_color.BackColor = vbWhite
        'txtColor.text = GetColorAtCursor
    End If

    'show_color_circle

End If


End Sub





Private Sub label_close_Click()
'Call Form_DblClick
Screen.MousePointer = 13
Unload Me

End Sub

Private Sub Label_desell_all_Click()

Dim i As Integer

'If Index <> 1 Then 'eliminer les selection
    For i = 0 To pic_prog.UBound
        pic_prog(i).WhatsThisHelpID = 0: actualise_ico_caption i
    Next i
'End If


End Sub

Private Sub Label_horloge1_DragDrop(Source As Control, X As Single, Y As Single)
If Source.Name = "Shape_pick_color" Then
    Label_horloge1.ForeColor = Shape_pick_color.BackColor
    Label_horloge2.ForeColor = Shape_pick_color.BackColor
End If

End Sub

Private Sub Label_horloge2_DragDrop(Source As Control, X As Single, Y As Single)
If Source.Name = "Shape_pick_color" Then
    Label_horloge1.ForeColor = Shape_pick_color.BackColor
    Label_horloge2.ForeColor = Shape_pick_color.BackColor
End If

End Sub

Private Sub label_min_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    Me.WindowState = 1
Else
    If Me.Height <= 300 Then
        Me.Height = mheight
        no_transparence
        For i = 0 To pic_prog.UBound
            pic_prog(i).Enabled = True
        Next i
    Else
        mheight = Me.Height
        Me.Height = 300
        'transparence
        Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(Me.hwnd, 0, 180, LWA_ALPHA)
        For i = 0 To pic_prog.UBound
            pic_prog(i).Enabled = False
        Next i
    End If
End If

End Sub


Private Sub Label_onglet_Click(index As Integer)

'charger_onglet index
pic_onglet_Click index

End Sub

Private Sub Label_sel_all_Click()
Dim i As Integer

If index <> 1 Then 'eliminer les selection
    For i = 0 To pic_prog.UBound
        pic_prog(i).WhatsThisHelpID = 1: actualise_ico_caption i
    Next i
End If

If indexencours = -1 Then 'si aucun objet sel alors sel le premier (il faut au moins un objet sel pour avior un indexencours)
    indexencours = 0
    'pic_prog(0).ZOrder 0
    'ctl.AttachControl pic_prog(0)
    
End If

End Sub

Private Sub label0_Click()
pic_frame_config.Visible = False
If Check_editable.Value = 1 Then
    Check_editable.Value = 0
    ctl.HideHandles
    Call Label_desell_all_Click
End If

End Sub




Private Sub pic_alignement_bas_Click()
alignement "bas"
End Sub

Private Sub pic_alignement_bas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_alignement_bas.BorderStyle = 1
End Sub

Private Sub pic_alignement_bas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_alignement_bas.BorderStyle = 0
End Sub


Private Sub pic_alignement_droit_Click()
alignement "droit"
End Sub

Private Sub pic_alignement_droit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_alignement_droit.BorderStyle = 1
End Sub


Private Sub pic_alignement_droit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_alignement_droit.BorderStyle = 0
End Sub


Private Sub pic_alignement_gauche_Click()
alignement "gauche"
End Sub


Private Sub pic_alignement_gauche_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_alignement_gauche.BorderStyle = 1

End Sub


Private Sub pic_alignement_gauche_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_alignement_gauche.BorderStyle = 0
End Sub


Private Sub pic_alignement_haut_Click()
alignement "haut"
End Sub

Private Sub pic_alignement_haut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_alignement_haut.BorderStyle = 1
End Sub


Private Sub pic_alignement_haut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_alignement_haut.BorderStyle = 0
End Sub



Private Sub pic_bordure_color_DragDrop(Source As Control, X As Single, Y As Single)

If Source.Name = "Shape_pick_color" Then
    
    pic_bordure_color.BackColor = Shape_pick_color.BackColor
    Shape_sel.BorderColor = Shape_pick_color.BackColor
    
End If



End Sub


Private Sub pic_close_desktopico_Click()
    AnimateForm Me, pic_desktop, -1, -1, aUnload, 20, 5, 2, 25
    pic_show_desktop.Enabled = True
    
End Sub

Private Sub pic_close_sur_mousemove_Click()
Dim i As Integer

If pic_close_sur_mousemove.Container.Name = "pic_prog" Then
    pic_prog(pic_close_sur_mousemove.Tag).Visible = False
End If
'suppr_prog pic_close_sur_mousemove.Tag
End Sub



Private Sub pic_frame_config_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

pic_frame_config.ZOrder 0

ReleaseCapture
SendMessage pic_frame_config.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

If pic_frame_config.Left < 0 Then pic_frame_config.Left = 0
If pic_frame_config.Top < 0 Then pic_frame_config.Top = 0


End Sub


Private Sub pic_onglet_Click(index As Integer)

If index < 3 Then
    Call charger_onglet(index)
Else
    Call charger_onglet2(index)
End If
End Sub

Public Sub pic_prog_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)

Dim i As Integer

If Source.Name = "Shape_pick_color" Then
    If Combo_color = "Fond" Then
        pic_prog(index).BackColor = Shape_pick_color.BackColor
    Else
        pic_prog(index).ForeColor = Shape_pick_color.BackColor
    End If
    'mon_freecolor = cmd_free_lecteur.BackColor
    
    'redessine l'icone
    actualise_ico_caption index
    
End If

If Source.Name = "txt_copier" Then
    If Combo_dim = "Left" Then txt_copier = pic_prog(index).Left
    If Combo_dim = "Top" Then txt_copier = pic_prog(index).Top
    If Combo_dim = "Width" Then txt_copier = pic_prog(index).Width
    If Combo_dim = "Height" Then txt_copier = pic_prog(index).Height
    
    'special
    If Combo_dim = "Bordure_droite" Then txt_copier = pic_prog(index).Left + pic_prog(index).Width
    If Combo_dim = "Bordure_inférieure" Then txt_copier = pic_prog(index).Top + pic_prog(index).Height
    If Combo_dim = "Top+Height" Then txt_copier = pic_prog(index).Top + pic_prog(index).Height + 15
    If Combo_dim = "Left+Width" Then txt_copier = pic_prog(index).Left + pic_prog(index).Width + 15
    
    
End If

If Source.Name = "pic_coller" Then
    If txt_copier = "" Then Exit Sub
    
    
    'appliquer sur tous les objets sélectionnés
    For i = 0 To pic_prog.UBound
        If pic_prog(i).WhatsThisHelpID = 1 Or i = index Then 'déjà sel
            If Combo_dim_destination = "Left" Then pic_prog(i).Left = txt_copier
            If Combo_dim_destination = "Top" Then pic_prog(i).Top = txt_copier
            If Combo_dim_destination = "Width" Then pic_prog(i).Width = txt_copier
            If Combo_dim_destination = "Height" Then pic_prog(i).Height = txt_copier
            
            If Combo_dim_destination = "Bordure_droite" Then
                'If pic_prog(i).Left + pic_prog(i).Width < txt_copier Then
                    If txt_copier > pic_prog(i).Left Then pic_prog(i).Width = txt_copier - pic_prog(i).Left
                'Else
                '    pic_prog(i).Width = txt_copier - pic_prog(i).Left + pic_prog(i).Width
                'End If
            End If
            If Combo_dim_destination = "Bordure_inférieure" Then
                If txt_copier > pic_prog(i).Top Then pic_prog(i).Height = txt_copier - pic_prog(i).Top
            End If
        
        End If
    Next i
    
    
    indexencours = index
    'pour avoir les même dimension
    pic_prog(index).ZOrder 0
    ctl.AttachControl pic_prog(index)
    
    
    'If Combo_dim = "Top+Height" Then txt_copier = pic_prog(Index).Top + pic_prog(Index).Height + 5
    'If Combo_dim = "Left+Width" Then txt_copier = pic_prog(Index).Left + pic_prog(Index).Width + 5
    
    
End If



End Sub


Public Sub pic_prog_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim txt As String
On Error Resume Next

Dim t As String
Dim chemin As String
Dim i As Integer
Dim icochemin As String



info_corbeille = "" 'pour forcer l'actualisation de l'info corbeille

'0 = front , 1 = send to back
If Button = vbMiddleButton Then
    pic_prog(index).ZOrder 1
    Exit Sub
End If


If Check_editable.Value = 0 Then
    
    indexencours = index
    
    If pic_prog(index).Tag = "" Then Exit Sub
    


        'animation
        pic_prog(index).FillStyle = 1
        
        For i = 1 To 10 '50 '50
            'cercle
            'pic_prog(Index).Circle ((pic_prog(Index).Width / 2), (pic_prog(Index).Height / 2)), i * 10, vbWhite
            pic_prog(index).Circle ((X), (Y)), i * 50, vbWhite
            pic_prog(index).Refresh
            Sleep 0.5 '1.5
            actualise_ico_caption index, ""
        Next i
        'actualise_ico_caption index, ""
        'Timer1_Timer 'refresh corbeille info
    
    If Button = 1 Then
        
        
        If Mid$(pic_prog(index).Tag, 1, Len("special")) = "special" Then
            
            If Shift = 0 Then
                'traitement spécial
                If Trim$(Mid$(pic_prog(index).Tag, Len("special") + 2, 2)) <> "::" Then
                    Shell Trim$(Mid$(pic_prog(index).Tag, Len("special") + 2, Len(pic_prog(index).Tag))), vbNormalFocus
                Else
                    'Shell "explorer.exe /n," & Trim$(Mid$(pic_prog(index).Tag, Len("special") + 2, Len(pic_prog(index).Tag))), vbNormalFocus
                    'corbeille
                    'If (mcaption(index) = "::{645FF040-5081-101B-9F08-00AA002F954E}" And corbeille_vide = False) Then
                    If mcaption(index) = "::{645FF040-5081-101B-9F08-00AA002F954E}" Then
                        If corbeille_vide = False Then Call open_corbeille(index)
                    Else
                        'Shell Trim$(Mid$(pic_prog(index).Tag, Len("special") + 2, Len(pic_prog(index).Tag))), vbNormalFocus
                        Shell "explorer.exe /n," & Trim$(Mid$(pic_prog(index).Tag, Len("special") + 2, Len(pic_prog(index).Tag))), vbNormalFocus
                    End If
                
                End If
            End If
        Else
            
            'launch
            If Shift = 1 Then
                Shell "explorer.exe /select, " & mtag(index) & "", vbNormalFocus
            ElseIf Shift = 3 Then 'ctrl+shift
                chemin = mtag(index)
                If LCase(Right$(chemin, 4)) = ".lnk" Then chemin = GetTarget(chemin)
                If File_Exist(chemin) = True Then ', vbArchive + vbHidden + vbNormal + vbNormal + vbSystem
                    Shell "explorer.exe /select, " & chemin, vbNormalFocus
                ElseIf File_Exist(chemin, True) = True Then ', vbDirectory + vbSystem + vbHidden + vbSystem + vbHidden
                    Shell "explorer.exe /select, " & chemin, vbNormalFocus
                End If
            ElseIf Shift = 2 Then 'touche ctrl
                'si dossier alors
                If File_or_folder(mtag(index)) = "Folder" Then
                    traitement_special_dossier mtag(index)
                Else
                    ShellExecute 0&, vbNullString, mtag(index), vbNullString, vbNullString, vbNormalFocus
                End If
            ElseIf Shift = 0 Then
                ShellExecute 0&, vbNullString, mtag(index), vbNullString, vbNullString, vbNormalFocus
                'actualise_ico_caption Index, ""
            End If
        End If
        
    Else
        
        
        Load frm_param
        
        If Mid$(pic_prog(index).Tag, 1, Len("special")) = "special" Then
            If (mcaption(index) = "::{645FF040-5081-101B-9F08-00AA002F954E}" And corbeille_vide = False) Then Call empty_recycle 'corbeille indexiconcorbeille = i: Exit For
            'If Shift <> 2 Then Exit Sub 'pas de caption pour les prog systeme
            If pic_prog(index).Tag = "special;::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then txt = "Poste de Travail"
            If pic_prog(index).Tag = "special;::{450D8FBA-AD25-11D0-98A8-0800361B1103}" Then txt = "Mes documents"
            If pic_prog(index).Tag = "special;::{208D2C60-3AEA-1069-A2D7-08002B30309D}" Then txt = "Connexions réseau"
            If pic_prog(index).Tag = "special;::{645FF040-5081-101B-9F08-00AA002F954E}" Then txt = "Corbeille"
            If pic_prog(index).Tag = "special;rundll32.exe shell32.dll,Control_RunDLL" Then txt = "Panneau de Configuration" '"Imprimantes et télécopieurs"
            frm_param.text1 = txt
            frm_param.text1.Enabled = False
            frm_param.Command1.Enabled = False
        Else
            frm_param.text1 = mcaption(index)
        End If
        
        If pic_prog(index).Tag = "special;::{645FF040-5081-101B-9F08-00AA002F954E}" Then Exit Sub 'si corbeille alors rien a faire
        
        icochemin = GetFromINI("config", "progico" & index, App.path & "\config.ini")
        If icochemin <> "" And File_Exist(icochemin) = True Then 'icone perso exist
            frm_param.Height = 3120
            frm_param.List_taille_ico.Visible = True
            frm_param.Command2.Visible = True
        Else
            frm_param.Height = 1095
            'si prog system et pas d'icone perso alors ne rien faire
            If frm_param.text1.Enabled = False Then Unload frm_param: Exit Sub
        End If
        move_frm_param index  'frm_param.Move Me.Left + pic_prog(index).Left + pic_prog(index).Width - 30, Me.Top + pic_prog(index).Top
        frm_param.Show 1, Me

    End If

Else 'editable

    If Button = 1 Then
        'MsgBox "probleme tjs frmmain pkoi ??!!: " & pic_prog(Index).Container.Name
''        mparent = pic_prog(index).Container.Name
''            If mparent <> "frmmain" Then
''                mparent = mparent & "(" & pic_prog(index).Container.index & ")"
''            End If
            'pic_prog(index).ZOrder 0
            'ctl.AttachControl pic_prog(index)
            
            If Shift <> 2 And Shift <> 1 Then Label_desell_all_Click 'si pas de touche ctrl et shift alors deselect all
            
            'marquer la selection
            If pic_prog(index).WhatsThisHelpID = 1 Then 'déjà sel
                pic_prog(index).WhatsThisHelpID = 0 'desel
                'indexencours = -1
                actualise_ico_caption index
                ctl.HideHandles
            Else
                pic_prog(index).WhatsThisHelpID = 1
                actualise_ico_caption index
                
                pic_prog(index).ZOrder 0
                ctl.AttachControl pic_prog(index)
                indexencours = index
            End If
            
            
    Else
    End If

End If





End Sub

Sub open_corbeille(index As Integer)
Dim err As Integer
Dim i As Integer

    'pic_show_desktop.Enabled = False 'pour empêcher le click alors que la procédure est toujours en cours

    err = shellfoldernav(&HA)
    
    If err = 0 Then
        pic_desktop.Width = Me.Width
        pic_desktop.Height = (Me.Height / 2) '- pic_barre.Height
        pic_desktop.Top = (Me.Height / 2) - pic_barre.Height
        WebBrowser1.Width = pic_desktop.Width - 60
        WebBrowser1.Height = pic_desktop.Height - 60
        
        pic_close_desktopico.Left = pic_desktop.Width - pic_close_desktopico.Width - 100
        
        'animation
        AnimateForm Me, pic_desktop, -1, -1, aload, 15, 5, 2, 25
        pic_desktop.Visible = True
        WebBrowser1.SetFocus
    Else 'erreur d'ouvreture de la corbeille sous seven
        Shell "explorer.exe /n," & Trim$(Mid$(pic_prog(index).Tag, Len("special") + 2, Len(pic_prog(index).Tag))), vbNormalFocus
    End If
End Sub

Private Sub pic_prog_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'caption
'    If indexencours = Index Then Exit Sub
'
'    indexencours = Index
    'label_caption = ""
    
If oldindex = index Then Exit Sub 'And Shape_sel2.Visible Then Exit Sub
    
If Check_editable.Value = 1 Then
    Shape_sel.Visible = False
    'sur mousemove mettre les controle move/resize
    If oldindex <> index And pic_prog(index).WhatsThisHelpID = 1 Then
        ctl.AttachControl pic_prog(index)
        oldindex = index
        indexencours = index
    End If
    Exit Sub
End If

If Check_caption_statusbar.Value = 1 Then 'And Index <> indexencours Then
    If Mid$(pic_prog(index).Tag, 1, Len("special")) = "special" Then
        label_caption = label_special(index)
    Else
        'If (label_caption = "" Or label_caption <> captionencours) Then
            label_caption = mtag(index)
            If LCase(Right$(label_caption, 3)) = "lnk" Then label_caption = label_caption & " (Raccourcis)"
        'End If
    End If
End If

    
'If check_change_backcolor_mousemove.Value = 1 And oldindex = index Then Exit Sub
    



    'si changement de backcolor sur mousemove
    If check_change_backcolor_mousemove.Value = 1 Then
        'restaurer color ancien
        If oldindex <> -1 Then
            pic_prog(oldindex).BackColor = old_backcolor
            actualise_ico_caption oldindex
            'si corbeille forcer refresh info
            If mcaption(oldindex) = "::{645FF040-5081-101B-9F08-00AA002F954E}" Then
                info_corbeille = "": Timer1_Timer
            End If
        End If
        'save backcolor actuelle
        old_backcolor = pic_prog(index).BackColor
        pic_prog(index).BackColor = 2407566: actualise_ico_caption index
        'pour forcer l'actualisation de l'info corbeille
        If mcaption(index) = "::{645FF040-5081-101B-9F08-00AA002F954E}" Then info_corbeille = "":         Timer1_Timer 'refresh corbeille
    Else
        Set Shape_sel.Container = pic_prog(index)
        Shape_sel.Visible = True
        Shape_sel.Left = 0 - (Shape_sel.Width / 2)
        Shape_sel.Top = 0 - (Shape_sel.Height / 2)
        Shape_sel.FillColor = BlendColor(pic_prog(index).BackColor, vbBlack, 150)
    
        If oldindex <> -1 Then
            'pic_prog(oldindex).BackColor = old_backcolor
            actualise_ico_caption oldindex
            'si corbeille forcer refresh info
            If mcaption(oldindex) = "::{645FF040-5081-101B-9F08-00AA002F954E}" Then
                info_corbeille = "": Timer1_Timer
            End If
        End If
    End If
    
    
Set pic_close_sur_mousemove.Container = pic_prog(index)
pic_close_sur_mousemove.Move pic_prog(index).Width - pic_close_sur_mousemove.Width - 30, 30 '15
pic_close_sur_mousemove.Tag = index
pic_close_sur_mousemove.Visible = True
    
    

'If pic_prog(index).BackColor = &H80FF& Then
'    Shape_sel.FillColor = vbBlack
'Else
'    Shape_sel.FillColor = &H80FF&
'End If




If Shape_sel2.Visible = True And oldindex = index Then Exit Sub

If simule_3d.Value = 1 Then
    If Check_roundcorner.Value = 1 Then
        'Set Shape_sel2.Container = pic_prog(index)
        'Shape_sel2.Left = 0 'pic_prog(index).Left + 200
        'Shape_sel2.Top = 0 'pic_prog(index).Top + 200
        'Shape_sel2.Width = pic_prog(index).Width + 30 ' 10 '10 '30
        'Shape_sel2.Height = pic_prog(index).Height + 30 '10 '10 '30
        'Shape_sel2.Visible = True
        If Check_cadre.Value = 1 Then
            pic_prog(index).Line (15, 0)-(pic_prog(index).Width - 45, 0), vbWhite, B
            pic_prog(index).Line (0, 15)-(0, pic_prog(index).Height - 45), vbWhite, B
            pic_prog(index).PSet (15, 15), vbWhite
        Else
            pic_prog(index).Line (15, 0)-(pic_prog(index).Width - 45, 0), vbWhite, B
            pic_prog(index).Line (0, 15)-(0, pic_prog(index).Height - 45), vbWhite, B
            pic_prog(index).PSet (15, 15), vbWhite
        End If
    Else
        Set Shape_sel2.Container = pic_prog(index)
        Shape_sel2.Left = 0 'pic_prog(index).Left + 200
        Shape_sel2.Top = 0 'pic_prog(index).Top + 200
        Shape_sel2.Width = pic_prog(index).Width + 15 ' 10 '10 '30
        Shape_sel2.Height = pic_prog(index).Height + 15 '10 '10 '30
        Shape_sel2.Visible = True
    End If
Else
    Shape_sel2.Visible = False
End If

pic_prog(index).ZOrder 0

oldindex = index

'autre méthode flêche avec replacecouleur ========
''pic_fleche.Picture = pic_fleche2.Picture
'ReplaceColor pic_fleche, vbBlack, pic_prog(index).BackColor
'pic_fleche.Visible = True
'pic_fleche.Left = (pic_prog(index).Width - pic_fleche.Width) / 2 '- 15
'pic_fleche.Top = 30  '- 10



indexencours = index


End Sub

Sub ReplaceColor(Pic As PictureBox, C1 As OLE_COLOR, C2 As OLE_COLOR)

Dim i As Integer, j As Integer
Dim pixel As Long
For i = 0 To Pic.Width
    For j = 0 To Pic.Height
        pixel = GetPixel(Pic.hdc, i, j)
        If pixel = C1 Then
            SetPixel Pic.hdc, i, j, C2
        End If
    Next j
Next i

End Sub

Public Sub pic_prog_OLEDragDrop(index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Tempo As String
    Dim capt_encours As String
    'Recupère le chemin déposé
    'param = ""
    
    Dim affiche_param_taille_icone As Boolean
    Dim taille As String * 3
    On Error GoTo err1
    'Dim fname As Variant
    
    Dim i As Integer
    
    capt_encours = mcaption(index)
    
    pic_prog(index).Cls
    
    'For Each fname In Data.Files
    'For i = 1 To Data.Files.Count
        If Data.GetFormat(vbCFFiles) Then
            Tempo = Data.Files(1)
            'si raccourcis alors retrouver le fichier d'origine
            'If LCase(Right$(Tempo, 4)) = ".lnk" Then Tempo = GetTarget(Tempo)
            'il vaut mieux travailler avec les raccourcis (cas de exeplorer.exe ,/e... ) utilisation
            'des param ds les raccourcis :)
            If Tempo <> "" Then
                
                If LCase(Right$(Tempo, 4)) = ".ico" Then 'ico perso
                    Call WriteToINI("config", "progico" & index, Trim$(Tempo), App.path & "\config.ini")
                    
                    'savoir si le prog doit afficher l'écran taille icone ou non
                        taille = GetFromINI("config", "progicotaille" & index, App.path & "\config.ini")
                        If Trim$(taille) = "" Then affiche_param_taille_icone = True Else affiche_param_taille_icone = False
                        
                    actualise_ico_caption index
                    'If List_taille_ico.Visible = True And List_taille_ico.Left = pic_prog(index).Left + pic_prog(index).Width Then
                    'Else
                    '    List_taille_ico.Visible = False
                    'End If
                    If affiche_param_taille_icone Then Call pic_prog_MouseDown(index, 2, Shift, X, Y)
                    Exit Sub
                End If
                
                'pic_prog(Index).CurrentX = 150: pic_prog(Index).CurrentY = 650
                'pic_prog(Index).Print Tempo
                pic_prog(index).Tag = Tempo
                'dessine_icone Tempo, Index
                If capt_encours <> "" Then
                    If MsgBox("Garder le même Nom du Programme?", vbYesNo + vbInformation) = vbYes Then
                        pic_prog(index).Tag = pic_prog(index).Tag & ";" & capt_encours
                        actualise_ico_caption index, capt_encours
                    Else
                        actualise_ico_caption index
                    End If
                Else
                    actualise_ico_caption index
                End If
            End If
        End If
    'Next i
    
'If Check_editable.Value = 1 Then
'    'indexencours = Index
'    ''pour avoir les même dimension
'    'pic_prog(Index).ZOrder 0
'    ctl.AttachControl pic_prog(Index)
'End If
    
    Exit Sub

err1:
MsgBox "Erreur : " & err.Number & " " & err.Description, vbCritical, App.Title
Exit Sub

End Sub

Private Sub pic_prog_Resize(index As Integer)

actualise_ico_caption index

End Sub

Private Sub pic_resize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'ResizeMe = True

    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0&
     
End Sub


Private Sub pic_resize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
'If ResizeMe Then
'
'    Width = X + Width
'    Height = Y + Height
'End If
End Sub


Private Sub pic_resize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'ResizeMe = False
End Sub



Private Sub pic_show_desktop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

label_caption = "Btn Gauche= Afficher les éléments du Bureau | Btn Droit= Option d'Arrêt de l'Ordinateur"

End Sub


Private Sub pic_show_desktop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim i As Integer

    If Button = 1 Then
        pic_show_desktop.Enabled = False 'pour empêcher le click alors que la procédure est toujours en cours
            shellfoldernav (&H0)
            pic_desktop.Width = Me.Width
            pic_desktop.Height = (Me.Height / 2) '- pic_barre.Height
            pic_desktop.Top = (Me.Height / 2) - pic_barre.Height
            WebBrowser1.Width = pic_desktop.Width - 60
            WebBrowser1.Height = pic_desktop.Height - 60
            'pic_close_desktopico.Left = pic_desktop.Width - pic_close_desktopico.Width - 100
            pic_close_desktopico.Move pic_desktop.Width - pic_close_desktopico.Width - 280, 15

            'animation
    '        For i = Me.Height To (Me.Height / 2) - pic_barre.Height Step -40
    '            pic_desktop.Top = i
    '            DoEvents
    '        Next i
            'pic_desktop.Top = (Me.Height / 2) - pic_barre.Height
            AnimateForm Me, pic_desktop, -1, -1, aload, 15, 5, 2, 25
            pic_desktop.Visible = True
            pic_desktop.ZOrder 0
            WebBrowser1.SetFocus
            WebBrowser1.ZOrder 0
            pic_close_desktopico.ZOrder 0
        pic_show_desktop.Enabled = True
    Else
        
        'transparence
        Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(Me.hwnd, 0, 120, LWA_ALPHA)
        frm_shutdown.Show , Me
    End If

End Sub


Private Sub pic_uniforme_hauteur_Click()
uniformiser "hauteur"
End Sub

Private Sub pic_uniforme_hauteur_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_uniforme_hauteur.BorderStyle = 1
End Sub


Private Sub pic_uniforme_hauteur_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_uniforme_hauteur.BorderStyle = 0
End Sub


Private Sub pic_uniforme_largeur_Click()
uniformiser "largeur"
End Sub

Private Sub pic_uniforme_largeur_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_uniforme_largeur.BorderStyle = 1
End Sub


Private Sub pic_uniforme_largeur_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_uniforme_largeur.BorderStyle = 0
End Sub


Private Sub pic_uniforme_les_deux_Click()
uniformiser "lesdeux"

End Sub

Private Sub pic_uniforme_les_deux_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_uniforme_les_deux.BorderStyle = 1
End Sub


Private Sub pic_uniforme_les_deux_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic_uniforme_les_deux.BorderStyle = 0
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    Shape_pick_color.BackColor = Picture1.Point(X, Y)
    'imgMarker.Visible = True
    imgMarker.Move X - imgMarker.Width / 2, Y - imgMarker.Height / 2
End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    Shape_pick_color.BackColor = Picture1.Point(X, Y)
    'imgMarker.Visible = True
    imgMarker.Move X - imgMarker.Width / 2, Y - imgMarker.Height / 2
End If

End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'imgMarker.Visible = False

End Sub


Private Sub Picture5_Click()

If Trim$(Me.txt_search) = "" Then Exit Sub

Load frm_search

If frm_search.Height > Me.Height Then Me.Height = frm_search.Height + 500

frm_search.Move Me.Left + 60, Me.Top + Me.Height - frm_search.Height - 150 'pic_barre.Height '- 150

Me.Enabled = False
Timer1.Enabled = False: Timer2.Enabled = False
    'transparence
    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(Me.hwnd, 0, 120, LWA_ALPHA)
    'If Trim$(txt_search) <> "" Then
    frm_search.text1 = Trim$(txt_search)
    frm_search.Show 1, Me

Timer1.Enabled = True: Timer2.Enabled = True

no_transparence
Me.Enabled = True

End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
label_caption = "Rechercher dans le Bureau et le Menu Démarrer"
End Sub


Private Sub Slider1_Change()

Dim i As Integer
For i = 0 To pic_prog.UBound
    actualise_ico_caption i, ""
Next i

info_corbeille = "" 'pour forcer l'actualisation de l'info corbeille


End Sub




Private Sub Text_name_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Text_name.Visible = False
    'pic_prog(index).Enabled = True
ElseIf KeyAscii = 13 Then
    Text_name.Visible = False
    't = InputBox("Veuillez spécifier le Nom du Programme", "Nom du Programme", mcaption(Index))
    'pic_prog(index).Enabled = True
    If Trim$(Text_name) <> "" Then
        pic_prog(indexencours).Tag = mtag(indexencours) & ";" & Trim$(Text_name)
        actualise_ico_caption indexencours, mcaption(indexencours)
    End If
End If
End Sub


Private Sub Text_name_LostFocus()
    Text_name.Visible = False
End Sub



Private Sub Timer1_Timer()





Dim index_corbeille As Integer
Dim nbre As Variant 'Integer
Dim taille As String


'Dim kb As Single

    QueryRecycleBin , c_size, num_items
    taile = c_size / 1024
    'lblSize.Caption = Format$(bin_size) & _
        " (" & Format(kb, "0.0") & " KB)"
    'lblNumItems.Caption = Format$(num_items)
    
    
    taille = c_size
    nbre = num_items
    

index_corbeille = indexiconcorbeille

If index_corbeille = -1 Then Exit Sub


taille = Format(taille / 1024 / 1024, "# ##0.00") & " Mo"

If info_corbeille = nbre & taille Then Exit Sub

info_corbeille = nbre & taille

    If nbre = 0 Then
        corbeille_vide = True
    Else
        corbeille_vide = False
    End If

    
actualise_ico_caption index_corbeille

    If nbre = 0 Then
        pic_prog(index_corbeille).CurrentX = 1000: pic_prog(index_corbeille).CurrentY = 80
        pic_prog(index_corbeille).Print "Corbeille Vide"
        'corbeille_vide = True
    Else
        'corbeille_vide = False
        pic_prog(index_corbeille).CurrentX = 1000: pic_prog(index_corbeille).CurrentY = 80
        pic_prog(index_corbeille).Print nbre & " Objet(s)"
        pic_prog(index_corbeille).CurrentX = 950: pic_prog(index_corbeille).CurrentY = 280
        pic_prog(index_corbeille).Print taille '& " Objet(s)"
    End If
'If pic_prog(index_corbeille).Height > 2400 Then
'    pic_prog(index_corbeille).CurrentX = 10: pic_prog(index_corbeille).CurrentY = 10
'    pic_prog(index_corbeille).Print nbre & " | " & taille
'Else



End Sub

Private Sub Timer2_Timer()
'test_souris_out

If WindowState = 1 Then Exit Sub

If Me.Height <= 300 Then pic_min_close.Visible = True: Exit Sub

Dim p As POINTAPI, X1 As Integer, Y1 As Integer
    
    GetCursorPos p

    X1 = p.X * Screen.TwipsPerPixelX
    Y1 = p.Y * Screen.TwipsPerPixelY

If X1 > Me.Left And Y1 > Me.Top And X1 < Me.Left + Me.Width And Y1 < Me.Top + Me.Height Then
'souris in
    'pic_min_close.Visible = True
    'pic_barre.Visible = True
    If Not barre_visible Then animation_barre 1 '1 show/0 hide
    
Else 'souris out
    
    If barre_visible Then animation_barre 0 '1 show/0 hide
    
'    top_pic_barre = pic_barre.Top
'    barre_visible = False
'    pic_min_close.Visible = False
'    pic_barre.Visible = False
    

End If


horloge

End Sub

Private Sub txt_search_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Picture5_Click
End Sub


Private Sub txt_search_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
label_caption = "Rechercher dans le Bureau et le Menu Démarrer"
End Sub


Private Sub WebBrowser1_LostFocus()
pic_desktop.Visible = False
pic_close_desktopico_Click
End Sub

