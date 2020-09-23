VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmFontList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CyberLat Font List Creator"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5610
   FillColor       =   &H80000004&
   ForeColor       =   &H8000000C&
   Icon            =   "FrmFontList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   7740
      Top             =   2970
   End
   Begin VB.ListBox Temp_list 
      Height          =   1230
      Left            =   6450
      TabIndex        =   65
      Top             =   3540
      Width           =   1725
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   7170
      Top             =   2940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox HTM2 
      Height          =   855
      Left            =   6510
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   47
      Text            =   "FrmFontList.frx":0E42
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox HTM1 
      Height          =   795
      Left            =   6480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "FrmFontList.frx":118B
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame CUADRO1 
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   5595
      Begin ComctlLib.ProgressBar Progress 
         Height          =   135
         Left            =   150
         TabIndex        =   66
         Top             =   3390
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton HTML_List 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3660
         Picture         =   "FrmFontList.frx":127C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "This will make an HTML formatted list."
         Top             =   3630
         Width           =   1815
      End
      Begin VB.CommandButton RTF_List 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1860
         Picture         =   "FrmFontList.frx":13D7
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "This will make a RTF file format"
         Top             =   3630
         Width           =   1815
      End
      Begin VB.CommandButton Cancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3630
         Width           =   1755
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Loading fonts...(you can use any code to load fonts)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   67
         Top             =   3120
         Width           =   5295
      End
      Begin VB.Image Logo 
         Height          =   2280
         Left            =   1080
         Picture         =   "FrmFontList.frx":151D
         Top             =   690
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CyberLat ""Formatted FontList"" Wizard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   5430
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame CUADRO2 
      Height          =   4095
      Left            =   0
      TabIndex        =   6
      Top             =   -60
      Width           =   5595
      Begin VB.Frame Frame2 
         Caption         =   "Select Fonts"
         Enabled         =   0   'False
         Height          =   2115
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   5355
         Begin VB.ListBox List1 
            Height          =   1635
            ItemData        =   "FrmFontList.frx":450B
            Left            =   180
            List            =   "FrmFontList.frx":450D
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   13
            Top             =   300
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   1635
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "In Progress...Wait..."
            Top             =   300
            Width           =   2655
         End
         Begin VB.Label TEST 
            Alignment       =   2  'Center
            Caption         =   "Font Test"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2910
            TabIndex        =   68
            Top             =   1320
            Width           =   2355
         End
      End
      Begin VB.CommandButton Next 
         Caption         =   "Next >>"
         Height          =   315
         Index           =   0
         Left            =   4260
         TabIndex        =   11
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton Back 
         Caption         =   "<< Back"
         Height          =   315
         Index           =   0
         Left            =   3060
         TabIndex        =   10
         Top             =   3660
         Width           =   1215
      End
      Begin VB.OptionButton Make_Selection 
         Caption         =   "Make a selection"
         Height          =   195
         Left            =   1020
         TabIndex        =   8
         ToolTipText     =   "You can make a font list on  CyberLat Fonter"
         Top             =   1140
         Width           =   1995
      End
      Begin VB.OptionButton List_Installed 
         Caption         =   "List all installed fonts"
         Height          =   195
         Left            =   1020
         TabIndex        =   7
         Top             =   840
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.CommandButton Exit 
         Caption         =   "Exit"
         Height          =   315
         Index           =   1
         Left            =   1860
         TabIndex        =   48
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Image ICONO 
         Height          =   480
         Left            =   300
         Top             =   840
         Width           =   480
      End
      Begin VB.Image ICO1 
         Height          =   480
         Left            =   3900
         Picture         =   "FrmFontList.frx":450F
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ICO2 
         Height          =   480
         Left            =   4620
         Picture         =   "FrmFontList.frx":4DD9
         Top             =   780
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         X1              =   180
         X2              =   5400
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label Step 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1: Selecting Fonts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   240
         Width           =   3225
      End
   End
   Begin VB.Frame CUADRO3 
      Height          =   4095
      Left            =   0
      TabIndex        =   15
      Top             =   -60
      Width           =   5595
      Begin VB.Frame Frame5 
         Height          =   675
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   5355
         Begin VB.CommandButton BGCOLOR 
            Caption         =   "BgColor"
            Height          =   315
            Left            =   3420
            TabIndex        =   69
            Top             =   240
            Width           =   915
         End
         Begin VB.ComboBox Font_Size 
            Height          =   315
            ItemData        =   "FrmFontList.frx":56A3
            Left            =   4410
            List            =   "FrmFontList.frx":56B9
            Style           =   2  'Dropdown List
            TabIndex        =   25
            ToolTipText     =   "Font size"
            Top             =   240
            Width           =   795
         End
         Begin VB.CommandButton Btn_Color 
            Caption         =   "FontColor"
            Height          =   315
            Left            =   2550
            TabIndex        =   21
            ToolTipText     =   "Background and font color"
            Top             =   240
            Width           =   885
         End
         Begin VB.OptionButton Opt_right 
            Height          =   315
            Left            =   900
            Picture         =   "FrmFontList.frx":56D4
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Right aligment"
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton Opt_Center 
            Height          =   315
            Left            =   540
            Picture         =   "FrmFontList.frx":581E
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Center text"
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton Opt_Left 
            Height          =   315
            Left            =   180
            Picture         =   "FrmFontList.frx":5968
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Left aligment"
            Top             =   240
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.CheckBox Underline 
            Height          =   315
            Left            =   2100
            Picture         =   "FrmFontList.frx":5AB2
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Underline"
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox Italic 
            Height          =   315
            Left            =   1740
            Picture         =   "FrmFontList.frx":5BFC
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Italic"
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox Bold 
            Height          =   315
            Left            =   1380
            Picture         =   "FrmFontList.frx":5D46
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Bold"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.CommandButton Next 
         Caption         =   "Next >>"
         Height          =   315
         Index           =   1
         Left            =   4260
         TabIndex        =   16
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton Back 
         Caption         =   "<< Back"
         Height          =   315
         Index           =   1
         Left            =   3060
         TabIndex        =   17
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Height          =   2295
         Left            =   120
         TabIndex        =   19
         Top             =   1260
         Width           =   5355
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   62
            Text            =   "Example 1"
            Top             =   1500
            Width           =   5055
         End
         Begin VB.CheckBox SPACE 
            Caption         =   "Add Double Space"
            Height          =   255
            Left            =   180
            TabIndex        =   46
            Top             =   660
            Width           =   1935
         End
         Begin VB.CheckBox UpperLabel 
            Caption         =   "Add a Small Caption"
            Height          =   255
            Left            =   180
            TabIndex        =   30
            ToolTipText     =   "Add a small font name on the top of the example text"
            Top             =   300
            Width           =   2055
         End
         Begin VB.ComboBox Font_Text 
            Height          =   315
            ItemData        =   "FrmFontList.frx":5E90
            Left            =   180
            List            =   "FrmFontList.frx":5EA9
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Image Image1 
            Height          =   1110
            Left            =   3300
            Picture         =   "FrmFontList.frx":5F7A
            Top             =   300
            Width           =   1905
         End
      End
      Begin VB.CommandButton Exit 
         Caption         =   "Exit"
         Height          =   315
         Index           =   2
         Left            =   1860
         TabIndex        =   49
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Label Step 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2: Formating Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   240
         Width           =   3165
      End
   End
   Begin VB.Frame CUADRO4 
      Height          =   4095
      Left            =   0
      TabIndex        =   31
      Top             =   -60
      Width           =   5595
      Begin VB.CheckBox BORDER 
         Enabled         =   0   'False
         Height          =   195
         Left            =   3000
         TabIndex        =   52
         Top             =   1020
         Width           =   195
      End
      Begin VB.OptionButton COL3 
         Height          =   795
         Left            =   240
         Picture         =   "FrmFontList.frx":665D
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2520
         Width           =   2535
      End
      Begin VB.OptionButton COL2 
         Height          =   795
         Left            =   240
         Picture         =   "FrmFontList.frx":6C05
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1740
         Width           =   2535
      End
      Begin VB.OptionButton COL1 
         Height          =   795
         Left            =   240
         Picture         =   "FrmFontList.frx":711F
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   960
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.CommandButton Next 
         Caption         =   "Next >>"
         Height          =   315
         Index           =   2
         Left            =   4260
         TabIndex        =   32
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton Back 
         Caption         =   "<< Back"
         Height          =   315
         Index           =   2
         Left            =   3060
         TabIndex        =   33
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton Exit 
         Caption         =   "Exit"
         Height          =   315
         Index           =   3
         Left            =   1860
         TabIndex        =   50
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "(Just for two or three cols)"
         Height          =   255
         Left            =   3300
         TabIndex        =   56
         Top             =   1260
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "If you choose three cols, we don't recomend you to fix a big font size."
         Height          =   615
         Left            =   3180
         TabIndex        =   55
         Top             =   2280
         Width           =   2040
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE:"
         Height          =   255
         Left            =   3000
         TabIndex        =   54
         Top             =   1980
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Add borders to table."
         Enabled         =   0   'False
         Height          =   195
         Left            =   3300
         TabIndex        =   53
         Top             =   1020
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4620
         Picture         =   "FrmFontList.frx":759F
         ToolTipText     =   "Only for HTML format"
         Top             =   1620
         Width           =   480
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000003&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   2940
         Shape           =   4  'Rounded Rectangle
         Top             =   3060
         Width           =   2535
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   5460
         Y1              =   3540
         Y2              =   3540
      End
      Begin VB.Label Step 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3: Formating Cols"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   180
         TabIndex        =   34
         Top             =   240
         Width           =   3210
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   180
         X2              =   5400
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000005&
         FillStyle       =   0  'Solid
         Height          =   2475
         Left            =   4140
         Shape           =   4  'Rounded Rectangle
         Top             =   900
         Width           =   1335
      End
   End
   Begin VB.Frame CUADRO5 
      Height          =   4095
      Left            =   0
      TabIndex        =   38
      Top             =   -60
      Width           =   5595
      Begin ComctlLib.ProgressBar Progress_Bar2 
         Height          =   165
         Left            =   300
         TabIndex        =   64
         Top             =   3360
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   291
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar Progress_Bar 
         Height          =   165
         Left            =   300
         TabIndex        =   63
         Top             =   3150
         Visible         =   0   'False
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   291
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Frame OPCIONES 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         TabIndex        =   57
         Top             =   1980
         Width           =   5175
         Begin VB.OptionButton OP 
            Caption         =   "Open the list on the default text editor and close this program"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   61
            Top             =   120
            Width           =   4635
         End
         Begin VB.OptionButton OP 
            Caption         =   "Open the list on the default text editor"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   3135
         End
         Begin VB.OptionButton OP 
            Caption         =   "Close this program."
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   1995
         End
         Begin VB.OptionButton OP 
            Caption         =   "Nothing"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.TextBox FilePath 
         Height          =   375
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1080
         Width           =   4815
      End
      Begin VB.CommandButton Save 
         Height          =   375
         Left            =   270
         Picture         =   "FrmFontList.frx":7E69
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Save list of selected items"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Next 
         Caption         =   "Create List !!!"
         Height          =   315
         Index           =   3
         Left            =   4260
         TabIndex        =   39
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton Back 
         Caption         =   "<< Back"
         Height          =   315
         Index           =   3
         Left            =   3060
         TabIndex        =   40
         Top             =   3660
         Width           =   1215
      End
      Begin VB.CommandButton Exit 
         Caption         =   "Exit"
         Height          =   315
         Index           =   0
         Left            =   1860
         TabIndex        =   51
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What to do after the Formatted List is created?"
         Height          =   195
         Left            =   1230
         TabIndex        =   45
         Top             =   1680
         Width           =   3315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Where do you want to save the Formatted List?"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   780
         Width           =   3360
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   5460
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Step 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 4: Saving the Formatted List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   180
         TabIndex        =   41
         Top             =   240
         Width           =   4635
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000009&
         X1              =   180
         X2              =   5400
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000005&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   300
         Shape           =   4  'Rounded Rectangle
         Top             =   1620
         Width           =   5175
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000003&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   5175
      End
   End
   Begin VB.Image Con_Sub 
      Height          =   1110
      Left            =   6480
      Picture         =   "FrmFontList.frx":7FB3
      Top             =   1740
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Sin_Sub 
      Height          =   555
      Left            =   6480
      Top             =   2880
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "FrmFontList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code was created by Mauricio Castelazo Gamboa
'email: castelazo@cyberlatino.com.mx
'this code is part of a fonter that I made, which has a lot of features,
'and I like this part because is easy to create these lists.

'I allow you to use the code, but please, always put in your aplications:

'      "THE FORMATTED FONTLIST CREATOR WAS MADE BY
'   MAURICIO CASTELAZO GAMBOA, FROM WWW.CYBERLATINO.COM.MX"

'remember that I'm just a begginer programmer  from México, and here is not
'so easy to study vb or a visual stuff. E-mail me your comments. visit my home
'page, and download the entire Freeware Fonter at:
'www.cyberlatino.com.mx/software

'ps. Most of the objects names are in spanish, and you won't find any
'comments....sorry :)


Dim INICIO_BAN As Boolean
Dim FORMATO As String


Private Sub Back_Click(Index As Integer)
On Error GoTo fin:
    Select Case Index
        Case 0:
            CUADRO1.Visible = True
        Case 1:
            CUADRO2.Visible = True
        Case 2:
            CUADRO3.Visible = True
        Case 3:
            If FORMATO = "RTF" Then CUADRO3.Visible = True
            CUADRO4.Visible = True
            Me.Next(2).Caption = "Next >>"
            Me.Next(2).Enabled = True
    End Select
Exit Sub
fin:
    MsgBox Err.Description, vbExclamation, "CyberLat Fonter 1.0: Error"
    Err.Clear
End Sub

Private Sub BGCOLOR_Click()
    Dim U As Long
    C.CancelError = False
    U = C.Color
    C.ShowColor
    If U <> C.Color Then Text3.BackColor = C.Color
End Sub

Private Sub Bold_Click()
    Text3.Font.Bold = CBool(Bold.Value)
End Sub

Private Sub Btn_Color_Click()
    Dim U As Long
    C.CancelError = False
    U = C.Color
    C.ShowColor
    If U <> C.Color Then Text3.ForeColor = C.Color
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub


Private Sub COL1_Click()
    Label2.Enabled = False
    BORDER.Enabled = False
    Label7.Enabled = False
End Sub

Private Sub COL2_Click()
    Label2.Enabled = True
    BORDER.Enabled = True
    Label7.Enabled = True
End Sub

Private Sub COL3_Click()
    Label2.Enabled = True
    Label7.Enabled = True
    BORDER.Enabled = True
End Sub



Private Sub Exit_Click(Index As Integer)
    Dim Resp As Integer
    Resp = MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion)
    If Resp = vbYes Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Font_Size_Click()
    Text3.FontSize = Font_Size.Text
End Sub

Private Sub Font_Text_Click()
On Error GoTo fin:
    If Font_Text.ListIndex = 0 Then
        Text3.Text = "Use the Font Name"
        Text3.Locked = True
    ElseIf Font_Text.ListIndex >= 1 And Font_Text.ListIndex < 6 Then
        Text3.Text = Font_Text.Text
        Text3.Locked = True
    ElseIf Font_Text.ListIndex = 6 Then
        Text3.Locked = False
        Text3.Text = "Introduce a text"
    End If
Exit Sub
fin:
    MsgBox Err.Description, vbExclamation, "CyberLat Fonter 1.0: Error"
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo fin:
    Font_Text.ListIndex = 0
    Font_Size.ListIndex = 2
    Sin_Sub.Picture = Image1.Picture
    ICONO.Picture = ICO1.Picture
    INICIO_BAN = False
    Timer1.Enabled = True

Exit Sub
fin:
    MsgBox Err.Description, vbExclamation, "CyberLat Fonter 1.0: Error"
    Err.Clear
End Sub


Private Sub Iniciar_Progress(Bar As ProgressBar, Minimo As Integer, Maximo As Integer, Valor As Integer)
On Error GoTo fin:
    Bar.Value = Valor
    Bar.Min = Minimo
    Bar.Max = Maximo
    Progress_Bar.Visible = True
    Progress_Bar2.Visible = True
Exit Sub
fin:
    MsgBox Err.Description, vbExclamation, "Fonter Error"
    Err.Clear
End Sub


Private Sub HTML_List_Click()
    FORMATO = "HTML"
    Me.Caption = "CyberLat Font List Creator: HTML List"
    CUADRO1.Visible = False
    BGCOLOR.Enabled = True
End Sub

Private Sub Italic_Click()
    Text3.Font.Italic = CBool(Italic.Value)
End Sub

Private Sub Label2_Click()
    BORDER.Value = (BORDER.Value - 1) * (-1)
End Sub

Private Sub Label7_Click()
    Label2_Click
End Sub



Private Sub List_Installed_Click()
    Frame2.Enabled = False
    ICONO.Picture = ICO1.Picture
End Sub



Private Sub List1_Click()
    Err.Clear
    On Error GoTo fin:
    TEST.Font.Name = List1.Text
    If TEST.Font.Bold = True Then TEST.Font.Bold = False
    If TEST.Font.Italic = True Then TEST.Font.Italic = False
    If TEST.Font.Size <> 12 Then TEST.Font.Size = 12
    Exit Sub
fin:
MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Load_Click()
    FrmFonter.Show
End Sub

Private Sub Make_Selection_Click()
    Frame2.Enabled = True
    ICONO.Picture = ICO2.Picture
End Sub



Private Sub Opt_Center_Click()
    Text3.Alignment = 2
End Sub

Private Sub Opt_Left_Click()
    Text3.Alignment = 0
End Sub

Private Sub Opt_right_Click()
    Text3.Alignment = 1
End Sub

Private Sub RTF_List_Click()
    FORMATO = "RTF"
    Text3.BackColor = vbWhite
    Me.Caption = "visit www.cyberlatino.com.mx: RTF List"
    CUADRO1.Visible = False
    BGCOLOR.Enabled = False
End Sub


Private Sub Save_Click()
    Dim A As String
    C.FileName = ""
    C.CancelError = False
    C.DialogTitle = "Where to save the file"
    Select Case FORMATO
        Case "RTF":
            C.Filter = "*.RTF format|*.rtf"
        Case "HTML":
            C.Filter = "*.HTM format|*.htm"
    End Select
    C.ShowSave
    If Trim(C.FileName) <> "" Then
        FilePath.Text = C.FileName
    End If
End Sub


Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Dim i As Integer
    Iniciar_Progress Progress, 0, Screen.FontCount, 0
    For i = 0 To Screen.FontCount - 1
        DoEvents
        List1.AddItem Screen.Fonts(i)
        Progress.Value = i
    Next i
    Progress.Visible = False
    RTF_List.Enabled = True
    HTML_List.Enabled = True
End Sub

Private Sub Underline_Click()
    Text3.Font.Underline = CBool(Underline.Value)
End Sub

Private Sub UpperLabel_Click()
    If UpperLabel.Value = 1 Then
        Image1.Picture = Con_Sub.Picture
        SPACE.Value = 1
    Else
        Image1.Picture = Sin_Sub.Picture
    End If
End Sub

Private Sub Execute(File As String)
    Dim W As Long
    On Error GoTo fin:
    W = Shell("rundll32.exe url.dll,FileProtocolHandler " & (FilePath), vbMaximizedFocus)
    Exit Sub
fin:      MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Next_Click(Index As Integer)
On Error GoTo fin:
    Select Case Index
        Case 0:
            CUADRO2.Visible = False
        Case 1:
            If FORMATO = "RTF" Then CUADRO4.Visible = False
            CUADRO3.Visible = False
        Case 2:
            CUADRO4.Visible = False
        Case 3:
            If FORMATO = "RTF" And Trim(Right(FilePath.Text, 3)) = "rtf" Then
                CargarRTF
            ElseIf FORMATO = "HTML" And Trim(Right(FilePath.Text, 3)) = "htm" Then
                CargarHTML
            Else
                FilePath.Text = ""
                Y = MsgBox("You have not choosen where do you want to save your list file?", vbInformation, "CyberLat Fonter 1.0")
                Save_Click
                Exit Sub
            End If
            If OP(0).Value = True Then
                Execute (FilePath.Text)
                Unload Me
                Exit Sub
            ElseIf OP(1).Value = True Then
                Execute (FilePath.Text)
            ElseIf OP(2).Value = True Then
                Unload Me
                Exit Sub
            Else
                Exit Sub
            End If
            Me.Exit(0).Default = True
    End Select
Exit Sub
fin:
    MsgBox Err.Description, vbExclamation, "CyberLat Fonter 1.0: Error"
    Err.Clear
End Sub

Private Sub CargarRTF()
    Dim i As Integer
    Dim LETRAS As String
    Dim Inicio As String
    Dim CONTENIDO As String
    Dim Color As String
    Dim red As Integer
    Dim green As Integer
    Dim blue As Integer
    Dim FINAL As String
    Dim USER_TEXT As String

    USER_TEXT = Text3.Text

    If Make_Selection.Value = True Then
        If List1.SelCount = 0 Then
            MsgBox ("You didn't choose any font")
            Exit Sub
        End If
        Temp_list.Clear
        For i = 0 To List1.ListCount - 1
            If List1.Selected(i) = True Then Temp_list.AddItem List1.List(i)
        Next i
    Else
        Temp_list.Clear
        For i = 0 To List1.ListCount - 1
            Temp_list.AddItem List1.List(i)
        Next i
    End If


    PERMITIR False
    Inicio = ""
    CONTENIDO = ""
    On Error GoTo fin:
    Err.Clear
    Iniciar_Progress Progress_Bar, 0, Temp_list.ListCount, 0
    Iniciar_Progress Progress_Bar2, 0, 3, 0
    Progress_Bar2.Value = 1

    LETRAS = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl" & vbCrLf
    
    If Font_Text.ListIndex = 0 Then
        For i = 0 To Temp_list.ListCount - 1
            DoEvents
            LETRAS = LETRAS & "{\f" & i & "\fnil\fprq2\fcharset0 " & Temp_list.List(i) & ";}" & vbCrLf
            If i > O And UpperLabel.Value = 0 Then CONTENIDO = CONTENIDO & "\f" & i & " " & Temp_list.List(i) & "\par"
            If i > O And UpperLabel.Value = 1 Then CONTENIDO = CONTENIDO & "\fs14\f1001 " & Temp_list.List(i) & "\par" & "\fs" & Int(Font_Size.Text) * 2 & "\f" & i & " " & Temp_list.List(i) & "\par"
            If SPACE.Value = 1 Then CONTENIDO = CONTENIDO & "\par"
            CONTENIDO = CONTENIDO & vbCrLf
            Progress_Bar.Value = i
        Next i
    ElseIf Font_Text.ListIndex > 0 And Font_Text.ListIndex < 6 Then
        For i = 0 To Temp_list.ListCount - 1
            DoEvents
            LETRAS = LETRAS & "{\f" & i & "\fnil\fprq2\fcharset0 " & Temp_list.List(i) & ";}" & vbCrLf
            If i > O And UpperLabel.Value = 0 Then CONTENIDO = CONTENIDO & "\f" & i & " " & Font_Text.Text & "\par"
            If i > O And UpperLabel.Value = 1 Then CONTENIDO = CONTENIDO & "\fs14\f1001 " & Temp_list.List(i) & "\par" & "\fs" & Int(Font_Size.Text) * 2 & "\f" & i & " " & Font_Text.Text & "\par"
            If SPACE.Value = 1 Then CONTENIDO = CONTENIDO & "\par"
            CONTENIDO = CONTENIDO & vbCrLf
            Progress_Bar.Value = i
        Next i
    ElseIf Font_Text.ListIndex = 6 Then
        For i = 0 To Temp_list.ListCount - 1
            DoEvents
            LETRAS = LETRAS & "{\f" & i & "\fnil\fprq2\fcharset0 " & Temp_list.List(i) & ";}" & vbCrLf
            If i > O And UpperLabel.Value = 0 Then CONTENIDO = CONTENIDO & "\f" & i & " " & USER_TEXT & "\par"
            If i > O And UpperLabel.Value = 1 Then CONTENIDO = CONTENIDO & "\fs14\f1001 " & Temp_list.List(i) & "\par" & "\fs" & Int(Font_Size.Text) * 2 & "\f" & i & " " & USER_TEXT & "\par"
            If SPACE.Value = 1 Then CONTENIDO = CONTENIDO & "\par"
            CONTENIDO = CONTENIDO & vbCrLf
            Progress_Bar.Value = i
        Next i
    End If
    Progress_Bar2.Value = 2

    LETRAS = LETRAS & "{\f1001" & "\fnil\fprq2\fcharset0 MS Sans Serif;}" & vbCrLf

    SEPARAR Text3.ForeColor, red, green, blue
    Color = "{\colortbl ;\red" & red & "\green" & green & "\blue" & blue & ";}"
    Inicio = Color & vbCrLf

'HERE STARTS EVERYTHING ABOUT THE BOLD, ITALIC AND UNDERLINE
    Inicio = Inicio & "\viewkind4\uc1\pard"
    If Bold.Value = 1 Then
        Inicio = Inicio & "\b "
        CONTENIDO = CONTENIDO & "\b0"
    End If
    If Italic.Value = 1 Then
        Inicio = Inicio & "\i "
        CONTENIDO = CONTENIDO & "\i0"
    End If
    If Underline.Value = 1 Then
        Inicio = Inicio & "\ul "
        CONTENIDO = CONTENIDO & "\ulnone"
    End If
    If Opt_right = True Then Inicio = Inicio & "\qr"
    If Opt_Center = True Then Inicio = Inicio & "\qc"
    Inicio = Inicio & "\fs" & Int(Font_Size.Text) * 2
    
    'AQUI TERMINA LO DE BOLD ITALIC ETC
    If Font_Text.ListIndex = 0 Then
        Inicio = Inicio & "\cf1\fs14\f1001 " & Temp_list.List(0) & "\par\fs" & Int(Font_Size.Text) * 2 & "\f0 " & Temp_list.List(0) & "\par"
        If SPACE.Value = 1 Then Inicio = Inicio & "\par"
        Inicio = Inicio & vbCrLf
    ElseIf Font_Text.ListIndex > 0 And Font_Text.ListIndex < 6 Then
        Inicio = Inicio & "\cf1\fs14\f1001 " & Temp_list.List(0) & "\par\fs" & Int(Font_Size.Text) * 2 & "\f0 " & Font_Text.Text & "\par"
        If SPACE.Value = 1 Then Inicio = Inicio & "\par"
        Inicio = Inicio & vbCrLf
    ElseIf Font_Text.ListIndex = 6 Then
        Inicio = Inicio & "\cf1\fs14\f1001 " & Temp_list.List(0) & "\par\fs" & Int(Font_Size.Text) * 2 & "\f0 " & USER_TEXT & "\par"
        If SPACE.Value = 1 Then Inicio = Inicio & "\par"
        Inicio = Inicio & vbCrLf
    End If
    LETRAS = LETRAS & "}"
    CONTENIDO = CONTENIDO & "}"
    FINAL = LETRAS & vbCrLf & Inicio & vbCrLf & CONTENIDO
    SAVE_TEXT FINAL
    Progress_Bar2.Value = 3
    Progress_Bar.Visible = False
    Progress_Bar2.Visible = False
    PERMITIR True
    Exit Sub
fin:
    PERMITIR True
    MsgBox Err.Description, vbExclamation, "CyberLat Fonter 1.0: Error"
    Err.Clear
End Sub

Private Sub CargarHTML()
    Dim i As Integer
    Dim FINAL As String
    Dim Inicio As String
    Dim CONTENIDO As String
    Dim red As Integer
    Dim green As Integer
    Dim blue As Integer
    Dim Color As String
    Dim UPER As String
    Dim TEMP As String
    Dim CONT As Integer
    Dim ABRIR As String
    Dim CERRAR As String
    Dim USER_TEXT As String

    If Make_Selection.Value = True Then
        If List1.SelCount = 0 Then
            MsgBox ("You didn't choose any font")
            Exit Sub
        End If
        Temp_list.Clear
        For i = 0 To List1.ListCount - 1
            If List1.Selected(i) = True Then Temp_list.AddItem List1.List(i)
        Next i
    Else
        Temp_list.Clear
        For i = 0 To List1.ListCount - 1
            Temp_list.AddItem List1.List(i)
        Next i
    End If

    USER_TEXT = Text3.Text
    ABRIR = ""
    CERRAR = ""
    CONT = 0
    Inicio = ""
    CONTENIDO = ""
    On Error GoTo fin:
    PERMITIR False
    Err.Clear
    Iniciar_Progress Progress_Bar, 0, Temp_list.ListCount, 0
    Iniciar_Progress Progress_Bar2, 0, 3, 0

    Progress_Bar2.Value = 1

    Color = HEXA(Text3.BackColor)
    Inicio = HTM1.Text & "<BODY bgcolor=" & Chr(34) & "#" & Color & Chr(34) & ">" & HTM2 & vbCrLf
    Inicio = Inicio & "<script> alert (" & Chr(34) & "CyberLat Fonter 1.0 \n\nThis list is registered to: Planet Source Users" & Chr(34) & ")</script>" & vbCrLf
    Color = HEXA(Text3.ForeColor)

    If Opt_right.Value = True Then
        Inicio = Inicio & "<P ALIGN=RIGHT>"
        ABRIR = ABRIR & "<P ALIGN=RIGHT>"
        CERRAR = "</P>" & CERRAR
    End If
    If Opt_Center.Value = True Then
        Inicio = Inicio & "<CENTER>"
        ABRIR = ABRIR & "<CENTER>"
        CERRAR = "</CENTER>" & CERRAR
    End If
    If Bold.Value = 1 Then
        Inicio = Inicio & "<B>"
        ABRIR = ABRIR & "<B>"
        CERRAR = "</B>" & CERRAR
    End If
    If Italic.Value = 1 Then
        Inicio = Inicio & "<I>"
        ABRIR = ABRIR & "<I>"
        CERRAR = "</I>" & CERRAR
    End If
    If Underline.Value = 1 Then
        Inicio = Inicio & "<U>"
        ABRIR = ABRIR & "<U>"
        CERRAR = "</U>" & CERRAR
    End If

    Inicio = Inicio & "<FONT COLOR=" & Color & " SIZE=" & Int(Font_Size.ListIndex) + 1 & ">"
    If COL2.Value = True Or COL3.Value = True Then Inicio = Inicio & "<TABLE BORDER=" & BORDER.Value * 3 & " WIDTH=100%><TR>"


   UPER = "<FONT NAME=" & Chr(34) & "ARIAL,VERDANA" & Chr(34) & " SIZE=1 COLOR=" & Chr(34) & Color & Chr(34) & ">"
    If COL1.Value = True Then
            If Font_Text.ListIndex = 0 Then
                For i = 0 To Temp_list.ListCount - 1
                    DoEvents
                    If UpperLabel.Value = 0 Then CONTENIDO = CONTENIDO & "<FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & ">" & Temp_list.List(i) & "</FONT><BR>"
                    If UpperLabel.Value = 1 Then CONTENIDO = CONTENIDO & UPER & Temp_list.List(i) & "</FONT><BR>" & "<FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & ">" & Temp_list.List(i) & "</FONT><BR>"
                    If SPACE.Value = 1 Then CONTENIDO = CONTENIDO & "<BR>"
                    CONTENIDO = CONTENIDO & vbCrLf
                    Progress_Bar.Value = i
                Next i
            ElseIf Font_Text.ListIndex > 0 Then
                For i = 0 To Temp_list.ListCount - 1
                    DoEvents
                    If UpperLabel.Value = 0 Then CONTENIDO = CONTENIDO & "<FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & ">" & USER_TEXT & "</FONT><BR>"
                    If UpperLabel.Value = 1 Then CONTENIDO = CONTENIDO & UPER & Temp_list.List(i) & "</FONT><BR>" & "<FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & ">" & USER_TEXT & "</FONT><BR>"
                    If SPACE.Value = 1 Then CONTENIDO = CONTENIDO & "<BR>"
                    CONTENIDO = CONTENIDO & vbCrLf
                    Progress_Bar.Value = i
                Next i
            End If
    ElseIf COL2.Value = True Then
            If Font_Text.ListIndex = 0 Then
                For i = 0 To Temp_list.ListCount - 1
                    CONT = CONT + 1
                    DoEvents
                    If UpperLabel.Value = 0 Then TEMP = "<TD width=50%><FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & " COLOR=" & Chr(34) & Color & Chr(34) & ">" & ABRIR & Temp_list.List(i) & CERRAR & "</FONT>"
                    If UpperLabel.Value = 1 Then TEMP = "<TD width=50%>" & UPER & Temp_list.List(i) & "</FONT><BR>" & "<FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & " COLOR=" & Chr(34) & Color & Chr(34) & ">" & ABRIR & Temp_list.List(i) & CERRAR & "</FONT>"
                    If SPACE.Value = 1 Then TEMP = TEMP & "<BR><BR>"
                    TEMP = TEMP & "</TD>"
                    If CONT = 2 Then
                        CONT = 0
                        TEMP = TEMP & vbCrLf & "</TR><TR>"
                    End If
                    CONTENIDO = CONTENIDO & TEMP & vbCrLf
                    Progress_Bar.Value = i
                Next i
            ElseIf Font_Text.ListIndex > 0 Then
                For i = 0 To Temp_list.ListCount - 1
                    CONT = CONT + 1
                    DoEvents
                    If UpperLabel.Value = 0 Then TEMP = "<TD width=50%><FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & " COLOR=" & Chr(34) & Color & Chr(34) & ">" & ABRIR & USER_TEXT & CERRAR & "</FONT></TD>"
                    If UpperLabel.Value = 1 Then TEMP = "<TD width=50%>" & UPER & Temp_list.List(i) & "</FONT><BR>" & "<FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & " COLOR=" & Chr(34) & Color & Chr(34) & ">" & ABRIR & USER_TEXT & CERRAR & "</FONT>"
                    If SPACE.Value = 1 Then TEMP = TEMP & "<BR><BR>"
                    TEMP = TEMP & "</TD>"
                    If CONT = 2 Then
                        CONT = 0
                        TEMP = TEMP & vbCrLf & "</TR><TR>"
                    End If
                    CONTENIDO = CONTENIDO & TEMP & vbCrLf
                    Progress_Bar.Value = i
                Next i
            End If
            If CONT = 1 Then CONTENIDO = CONTENIDO & "<TD width=50%></TD></TR>"
    ElseIf COL3.Value = True Then
            If Font_Text.ListIndex = 0 Then
                For i = 0 To Temp_list.ListCount - 1
                    CONT = CONT + 1
                    DoEvents
                    If UpperLabel.Value = 0 Then TEMP = "<TD width=33%><FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & " COLOR=" & Chr(34) & Color & Chr(34) & ">" & ABRIR & Temp_list.List(i) & CERRAR & "</FONT></TD>"
                    If UpperLabel.Value = 1 Then TEMP = "<TD width=33%>" & UPER & Temp_list.List(i) & "</FONT><BR>" & "<FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & " COLOR=" & Chr(34) & Color & Chr(34) & ">" & ABRIR & Temp_list.List(i) & CERRAR & "</FONT>"
                    If SPACE.Value = 1 Then TEMP = TEMP & "<BR><BR>"
                    TEMP = TEMP & "</TD>"
                    If CONT = 3 Then
                        CONT = 0
                        TEMP = TEMP & vbCrLf & "</TR><TR>"
                    End If
                    CONTENIDO = CONTENIDO & TEMP & vbCrLf
                    Progress_Bar.Value = i
                Next i
            ElseIf Font_Text.ListIndex > 0 Then
                For i = 0 To Temp_list.ListCount - 1
                    CONT = CONT + 1
                    DoEvents
                    If UpperLabel.Value = 0 Then TEMP = "<TD width=33%><FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & " COLOR=" & Chr(34) & Color & Chr(34) & ">" & ABRIR & USER_TEXT & CERRAR & "</FONT></TD>"
                    If UpperLabel.Value = 1 Then TEMP = "<TD width=33%>" & UPER & Temp_list.List(i) & "</FONT><BR>" & "<FONT SIZE=" & Trim(Str(Font_Size.ListIndex + 1)) & " FACE=" & Chr(34) & Temp_list.List(i) & Chr(34) & " COLOR=" & Chr(34) & Color & Chr(34) & ">" & ABRIR & USER_TEXT & CERRAR & "</FONT>"
                    If SPACE.Value = 1 Then TEMP = TEMP & "<BR><BR>"
                    TEMP = TEMP & "</TD>"
                    If CONT = 3 Then
                        CONT = 0
                        TEMP = TEMP & vbCrLf & "</TR><TR>"
                    End If
                    CONTENIDO = CONTENIDO & TEMP & vbCrLf
                    Progress_Bar.Value = i
                Next i
            End If
            If CONT = 1 Then CONTENIDO = CONTENIDO & "<TD width=33%></TD><TD width=33%></TD></TR>"
            If CONT = 2 Then CONTENIDO = CONTENIDO & "<TD width=33%></TD></TR>"
    End If

   Progress_Bar2.Value = 2
    FINAL = vbCrLf & "</TABLE></FONT>"
    If Bold.Value = 1 Then FINAL = FINAL & "</B>"
    If Italic.Value = 1 Then FINAL = FINAL & "</I>"
    If Underline.Value = 1 Then FINAL = FINAL & "</U>"
    If Opt_right = True Then FINAL = FINAL & "</P>"
    If Opt_Center = True Then FINAL = FINAL & "</P>"

    FINAL = FINAL & vbCrLf & "</BODY><HTML>"
    Inicio = Inicio & vbCrLf & CONTENIDO & vbCrLf & FINAL
 
    SAVE_TEXT Inicio
    Progress_Bar2.Value = 3
    PERMITIR True
    Progress_Bar.Visible = False
    Progress_Bar2.Visible = False
    Exit Sub
fin:
    PERMITIR True
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub SAVE_TEXT(TEXTO As String)
    On Error GoTo fin:
    Open FilePath.Text For Output As #1
    Print #1, TEXTO
    Close #1
    Beep
    Exit Sub
fin:
    MsgBox Err.Description, vbExclamation, "Error"
    Err.Clear
End Sub

Private Sub PERMITIR(Resp As Boolean)
On Error GoTo fin:
    Me.Next(3).Enabled = Resp
    Me.Back(3).Enabled = Resp

    Me.Exit(0).Enabled = Resp
    Save.Enabled = Resp
    OPCIONES.Enabled = Resp
Exit Sub
fin:
    MsgBox Err.Description, vbExclamation, "CyberLat Fonter 1.0: Error"
    Err.Clear
End Sub

Private Sub SEPARAR(ByVal Color As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
    Dim TEMP As Long
    TEMP = (Color And 255)
    red = TEMP And 255
    TEMP = Int(Color / 256)
    green = TEMP And 255
    TEMP = Int(Color / 65536)
    blue = TEMP And 255
       
End Sub

Private Function HEXA(Color As Long) As String
On Error GoTo fin:
    Dim HexRed As String
    Dim HexGreen As String
    Dim HexBlue As String
    Dim R As Integer
    Dim G As Integer
    Dim b As Integer

    SEPARAR Color, R, G, b
    HexRed = Hex$(R)
    If Len(HexRed) = 1 Then HexRed = "0" & HexRed
    HexGreen = Hex$(G)
    If Len(HexGreen) = 1 Then HexGreen = "0" & HexGreen
    HexBlue = Hex$(b)
    If Len(HexBlue) = 1 Then HexBlue = "0" & HexBlue
    HEXA = HexRed & HexGreen & HexBlue
Exit Function
fin:
    MsgBox Err.Description, vbExclamation, "CyberLat Fonter 1.0: Error"
    Err.Clear
End Function
