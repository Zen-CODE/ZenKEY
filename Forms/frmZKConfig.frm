VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmZKConfig 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "| ZenKEY configuration |"
   ClientHeight    =   9360
   ClientLeft      =   3840
   ClientTop       =   1335
   ClientWidth     =   15420
   ClipControls    =   0   'False
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmZKConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   624
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   Begin VB.CommandButton zbExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7500
      TabIndex        =   32
      Top             =   3690
      Width           =   1695
   End
   Begin VB.CommandButton zbSave 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   31
      Top             =   3690
      Width           =   1695
   End
   Begin VB.TextBox txtMode 
      Height          =   315
      Index           =   3
      Left            =   3900
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   8760
      Width           =   630
   End
   Begin VB.TextBox txtMode 
      Height          =   315
      Index           =   2
      Left            =   4620
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   8760
      Width           =   630
   End
   Begin VB.TextBox txtMode 
      Height          =   315
      Index           =   1
      Left            =   5280
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   8760
      Width           =   630
   End
   Begin VB.TextBox txtMode 
      Height          =   315
      Index           =   0
      Left            =   3180
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8760
      Width           =   630
   End
   Begin VB.Timer tmrEdit 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2220
      Top             =   8760
   End
   Begin VB.PictureBox picItems 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3435
      Left            =   6480
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   549
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Move the mouse over an item to find out more about it."
      Top             =   360
      Width           =   8235
      Begin VB.CommandButton zbMove 
         Caption         =   "&Move"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7140
         TabIndex        =   36
         Top             =   3000
         Width           =   915
      End
      Begin VB.CommandButton zbDel 
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6180
         TabIndex        =   35
         Top             =   3000
         Width           =   915
      End
      Begin VB.CommandButton zbEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5220
         TabIndex        =   34
         Top             =   3000
         Width           =   915
      End
      Begin VB.CommandButton zbNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4260
         TabIndex        =   33
         Top             =   3000
         Width           =   915
      End
      Begin MSComctlLib.TreeView tvTree 
         Height          =   2835
         Left            =   60
         TabIndex        =   7
         ToolTipText     =   "Right click for options, double-click to edit/expand/collapse."
         Top             =   540
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   5001
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlTree"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkEnabled 
         BackColor       =   &H00FFFFFF&
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
         Height          =   210
         Left            =   7800
         TabIndex        =   1
         ToolTipText     =   "Enable or disable this group or item. This will prevent the item/group from showing in 'ZenKEY'."
         Top             =   660
         Width           =   210
      End
      Begin VB.CheckBox chkRClick 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Right click menu"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5760
         TabIndex        =   2
         ToolTipText     =   "Set this property to make this menu appear when you right click on the form or system tray."
         Top             =   2430
         Width           =   1635
      End
      Begin VB.Label lblZKProp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Properties"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   38
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label lblZKMenu 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ZenKEY Menu"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   120
         Width           =   3495
      End
      Begin VB.Image imiZKMenu 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Left            =   60
         Picture         =   "frmZKConfig.frx":058A
         Stretch         =   -1  'True
         Top             =   60
         Width           =   3870
      End
      Begin VB.Label lblAdd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Addtional"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4260
         TabIndex        =   14
         ToolTipText     =   "Lists any Operating Systems the action will not work on."
         Top             =   2220
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parameters"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4260
         TabIndex        =   9
         ToolTipText     =   "This is the parameter that will be passed to the executable file."
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4260
         TabIndex        =   8
         ToolTipText     =   "The caption of the item as it appears in the menu"
         Top             =   720
         Width           =   585
      End
      Begin VB.Image imiGroupOpen 
         Height          =   480
         Left            =   7500
         Picture         =   "frmZKConfig.frx":0C87
         ToolTipText     =   "Indicates this item is a group. Click to collapse this menu in the left pane."
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image imiItem 
         Height          =   480
         Left            =   7500
         Picture         =   "frmZKConfig.frx":18C9
         ToolTipText     =   "Click here to test the action."
         Top             =   1740
         Width           =   480
      End
      Begin VB.Image imiGroup 
         Height          =   480
         Left            =   7500
         Picture         =   "frmZKConfig.frx":250B
         ToolTipText     =   "Indicates this item is a group. Click to expand this menu in the left pane."
         Top             =   2220
         Width           =   480
      End
      Begin VB.Label lblHotkey 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hotkey"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4260
         TabIndex        =   5
         ToolTipText     =   "The key combination that will fire the action"
         Top             =   1620
         Width           =   540
      End
      Begin VB.Label lblAction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4260
         TabIndex        =   4
         ToolTipText     =   "The 'ZenKEY' description of the items' action"
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblActionType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4260
         TabIndex        =   3
         ToolTipText     =   "Describes the type of action to be performed"
         Top             =   1020
         Width           =   870
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   2835
         Index           =   5
         Left            =   3960
         Top             =   540
         Width           =   4215
      End
      Begin VB.Image imiZKProp 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Left            =   4080
         Stretch         =   -1  'True
         Top             =   60
         Width           =   4095
      End
   End
   Begin VB.PictureBox picNewSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3435
      Left            =   1020
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   549
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Move the mouse over an item to find out more about it."
      Top             =   4140
      Width           =   8235
      Begin VB.Frame frmSetting 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2685
         Left            =   4080
         TabIndex        =   20
         Top             =   570
         Visible         =   0   'False
         Width           =   4005
         Begin VB.CommandButton zbBackup 
            Cancel          =   -1  'True
            Caption         =   "Backup"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2640
            TabIndex        =   44
            Top             =   1680
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CommandButton zbSetFileRemove 
            Caption         =   "&Remove"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2520
            TabIndex        =   43
            Top             =   2100
            Width           =   915
         End
         Begin VB.CommandButton zbSetFileAdd 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1440
            TabIndex        =   42
            Top             =   2100
            Width           =   915
         End
         Begin VB.CommandButton zbQuoteNow 
            Caption         =   "&Now"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   41
            Top             =   2100
            Width           =   915
         End
         Begin VB.TextBox txtSet 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1620
            MaxLength       =   3
            TabIndex        =   25
            Top             =   1620
            Width           =   375
         End
         Begin VB.ComboBox cmbSet 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1140
            Width           =   1575
         End
         Begin VB.Shape shpSetColour 
            BorderColor     =   &H00008080&
            FillColor       =   &H00008080&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1080
            Top             =   720
            Width           =   195
         End
         Begin VB.Label lblSetColour 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Active "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   300
            TabIndex        =   27
            Top             =   720
            Width           =   525
         End
         Begin VB.Label lblTxtPostfix 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Postfix"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2100
            TabIndex        =   26
            Top             =   1740
            Width           =   525
         End
         Begin VB.Label lblSetNotes 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Notes"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -255
            TabIndex        =   24
            Top             =   1740
            Width           =   2985
            WordWrap        =   -1  'True
         End
         Begin VB.Image imiSet 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   180
            Top             =   1740
            Width           =   555
         End
         Begin VB.Label lblSetValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current value"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   300
            TabIndex        =   23
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label lblSetDescrip 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   300
            TabIndex        =   22
            Top             =   240
            Width           =   3195
            WordWrap        =   -1  'True
         End
      End
      Begin MSComctlLib.TreeView tvSettings 
         Height          =   2835
         Left            =   60
         TabIndex        =   19
         ToolTipText     =   "Select the group of settings, and then the setting that you wish to edit."
         Top             =   540
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   5001
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlSettings"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblZKNewSetting 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   120
         Width           =   3495
      End
      Begin VB.Image imiZKNewSetting 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Left            =   4080
         Stretch         =   -1  'True
         Top             =   60
         Width           =   4080
      End
      Begin VB.Label lblZKSettings 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ZenKEY Settings"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   120
         Width           =   3495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   2715
         Index           =   9
         Left            =   4080
         Top             =   540
         Width           =   4035
      End
      Begin VB.Image imiZKSettings 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Left            =   60
         Stretch         =   -1  'True
         Top             =   60
         Width           =   3870
      End
   End
   Begin MSComctlLib.ImageList imlTree 
      Left            =   780
      Top             =   8700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":314D
            Key             =   "Action"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":349F
            Key             =   "Moving"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":37F1
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":3B43
            Key             =   "FolderOpen"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSettings 
      Left            =   1440
      Top             =   8700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":3E95
            Key             =   "One"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":4206
            Key             =   "Two"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":458A
            Key             =   "Three"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":4909
            Key             =   "Four"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":4C87
            Key             =   "Five"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":4FFE
            Key             =   "Six"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":5367
            Key             =   "Seven"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZKConfig.frx":56E4
            Key             =   "Ying"
         EndProperty
      EndProperty
   End
   Begin VB.Image imiItems 
      Height          =   660
      Left            =   240
      Picture         =   "frmZKConfig.frx":5A61
      ToolTipText     =   "Edit the items on the ZenKEY menus and assign Hotkeys"
      Top             =   180
      Width           =   660
   End
   Begin VB.Image imiSettings 
      Height          =   660
      Left            =   240
      Picture         =   "frmZKConfig.frx":621D
      ToolTipText     =   "Edit the ZenKEY settings"
      Top             =   1185
      Width           =   660
   End
   Begin VB.Image imiAbout 
      Height          =   660
      Left            =   240
      Picture         =   "frmZKConfig.frx":6A9C
      ToolTipText     =   "Find out more about ZenKEY"
      Top             =   2175
      Width           =   660
   End
   Begin VB.Label lblItems 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Items"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   345
      TabIndex        =   12
      ToolTipText     =   "Edit the items on the ZenKEY menus and assign Hotkeys"
      Top             =   840
      Width           =   420
   End
   Begin VB.Label lblSettings 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Settings"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   285
      TabIndex        =   15
      ToolTipText     =   "Edit the ZenKEY settings"
      Top             =   1845
      Width           =   630
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   300
      TabIndex        =   17
      ToolTipText     =   "Find out more about ZenKEY"
      Top             =   2835
      Width           =   510
   End
   Begin VB.Shape shpFocus 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   975
      Left            =   135
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   360
      TabIndex        =   18
      ToolTipText     =   "Find out more about ZenKEY"
      Top             =   3840
      Width           =   390
   End
   Begin VB.Image imiHelp 
      Height          =   660
      Left            =   240
      Picture         =   "frmZKConfig.frx":723C
      ToolTipText     =   "Find out more about ZenKEY"
      Top             =   3180
      Width           =   660
   End
   Begin VB.Label lblFooter 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZenKEY Configuration"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3885
      TabIndex        =   11
      Top             =   3720
      Width           =   1785
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   4200
      TabIndex        =   10
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Label lblBy 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "by ZenCODE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   7440
      TabIndex        =   6
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Image imiYinYang 
      Height          =   2295
      Left            =   3600
      Picture         =   "frmZKConfig.frx":79AE
      ToolTipText     =   "Move the mouse over an item to find out more about it."
      Top             =   660
      Width           =   2310
   End
   Begin VB.Menu mnuMov 
      Caption         =   "Move"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit item"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "New item"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Remove"
         Visible         =   0   'False
         Begin VB.Menu mnuDelete 
            Caption         =   "Remove item"
         End
         Begin VB.Menu mnuHKDel 
            Caption         =   "Remove Hotkey"
         End
         Begin VB.Menu mnuHKDelAll 
            Caption         =   "Remove all Hotkeys"
         End
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveUp 
         Caption         =   "Move item up"
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "Move item down"
      End
      Begin VB.Menu mnuMoveBefore 
         Caption         =   "Move to before another item"
      End
      Begin VB.Menu mnuMoveAfter 
         Caption         =   "Move to after another item"
      End
      Begin VB.Menu mnuGroup 
         Caption         =   "Group"
         Begin VB.Menu mnuClearDead 
            Caption         =   "Clear dead links"
         End
         Begin VB.Menu mnuSortGroup 
            Caption         =   "Sort this group"
         End
         Begin VB.Menu mnuSortAll 
            Caption         =   "Sort all groups"
         End
         Begin VB.Menu mnuExpand 
            Caption         =   "Expand group"
         End
         Begin VB.Menu mnuCollapse 
            Caption         =   "Collapse group"
         End
      End
   End
End
Attribute VB_Name = "frmZKConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim booChanged As Boolean
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Dim objSelected As Control
Dim booSetChanged As Boolean
Dim booLoading As Boolean
Dim ModeLabel As Label
Dim lblClicked As Label
Dim ZIndex As Long
Private Const ColDisabled = &H808080   '&HC0C0C0
Private booMoving As Boolean
Private lngSource As Long
Private booMoveAfter As Boolean
Private booRightClick As Boolean
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Rem ---------------------------------------------------------------------------------------------------------------------------
Rem - For the colour dialog box
Rem ---------------------------------------------------------------------------------------------------------------------------
Private Type CHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type PRINTDLG_TYPE
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Dim CustomColors() As Byte
Const List_Null = "<None>"
Dim SET_CurIndex As Long
Dim nodDrag As Node, booDrag As Boolean
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Sub Array_Up(ByVal Start As Long, ByVal Num As Long)
Dim k As Integer
Dim max As Integer

    max = UBound(ZKMenu())
    
    For k = Start To max
        Set ZKMenu(k - Num) = ZKMenu(k)
    Next k
    
    ReDim Preserve ZKMenu(max - Num)
    
    
End Sub

Private Sub chkEnabled_Click()
Dim booEnabled As Boolean

    If Not booLoading Then
        If ZKMenu(ZIndex)("Disabled") = "True" Then
            ZKMenu(ZIndex)("Disabled") = "False"
        Else
            ZKMenu(ZIndex)("Disabled") = "True"
        End If
        Call Item_Selected
        booChanged = True
    End If
    
End Sub

Private Sub cmbSet_Click()

    If Not booLoading Then Call Set_SetValue(tvSettings.SelectedItem.Tag)

End Sub


Private Sub chkRClick_Click()
    
    If Not booLoading Then
        Dim lngPrev As Long, k As Long
        Dim max As Long
        
        Rem - Check if the right click menu is already in use....
        max = UBound(ZKMenu())
        For k = 2 To max
            If ZKMenu(k)("RightClickMenu") = "True" Then
                lngPrev = k
                Exit For
            End If
        Next k
        
        Dim booEnabled As Boolean
        booEnabled = CBool(chkRClick.Value = 1)
        
        If (lngPrev > 0) And booEnabled Then
            If ZenMB("The 'Right Click' menu is already enabled for '" & ZKMenu(lngPrev)("Caption") & "'. Do you wish to set it to this Menu instead?", "Yes", "No") = 1 Then
                booLoading = True
                chkRClick.Value = 0
                booLoading = False
                Exit Sub
            End If
            ZKMenu(lngPrev)("RightClickMenu") = vbNullString
        End If
        
        ZKMenu(ZIndex)("RightClickMenu") = IIf(booEnabled, "True", vbNullString)
        booChanged = True
    End If
    
End Sub


Private Sub Form_Click()

    If booMoving Then Call Move_Selected(booMoveAfter)
End Sub

Private Sub Form_Initialize()
Dim X As Long
X = InitCommonControls

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Static Typed As String

    Typed = Right$(Typed, 12) & UCase(Chr$(KeyAscii))
    Select Case True
        Case Right$(Typed, 7) = "BETTINA": Call ZenMB("Too easy, could be anyone. Try your animal name. You know, like Ouch, Bobbin..", "OK")
        Case Right$(Typed, 4) = "PTAT": Call ZenMB("Hey 'tat. Missed you everyday for a year, you know dat? You iz a most bewtiful kitten! Love you forever...", "OK")
        Case Right$(Typed, 5) = "JOHAN": Call ZenMB("Snotklap! Ha ha ha ha....Big up, piesang!", "OK")
        Case Right$(Typed, 4) = "PAUL": Call ZenMB("Whatsuuuuup? Die oue wat 'three pointed stars' poep. We may be different, but character is bigger than difference. Respect! ", "OK")
        Case Right$(Typed, 8) = "GIOVANNI": Call ZenMB("Yo stinkfoot! For an overgrown Italian prattboy, you're still a chop! Love ya, knuckle...", "OK")
        Case Right$(Typed, 8) = "KAROLINE": Call ZenMB("The princess!!! My hero & saviour - Karoline the beautiful! You rule Kazzy. I love you like a lion...", "OK")
        Case Right$(Typed, 6) = "MARTIN": Call ZenMB("Hey stinky! What are you doing using my program? Get your own, sick-boy! Love ya, you chopstick...", "OK")
        Case Right$(Typed, 9) = "CATHERINE": Call ZenMB("Tveetles!!! Oh darhling, you have become such a gem. Love you baby!", "OK")
        Case Right$(Typed, 5) = "DANNY": Call ZenMB("Danny boy! Dude of note! Much love...", "OK")
        Case Right$(Typed, 7) = "TRENTON": Call ZenMB("Trendoid! Dude, you've grown to be one of the biggest people I know. Respekt!. Much love ...(Tuneage forever!!!)", "OK")
        Case Right$(Typed, 5) = "JAMIE": Call ZenMB("Time has marched endlessly between us. Love you still, forever... ", "OK")
        Case Right$(Typed, 6) = "CASPAR": Call ZenMB("Yo boet! Still gonna kick your butt for forcing me off the track in GP1...", "OK")
        Case Right$(Typed, 3) = "MUM": Call ZenMB("Mum! What can be said. You are my world...", "OK")
        Case Right$(Typed, 4) = "JOHN": Call ZenMB("Yo brudda man!! Proof that madness is is my blood! Big up!", "OK")
        Case Right$(Typed, 4) = "ERIC": Call ZenMB("Ello Monsieur! Vanilla fantastique. onion pudding, sushi - and your heart is as good as your cooking! Much love..", "OK")
        Case Right$(Typed, 3) = "LEE": Call ZenMB("Monsieur Lee! Maniac Madely, more fried than a banana split! Much love..", "OK")
        Case Right$(Typed, 6) = "CLAIRE": Call ZenMB("Clairyuffski! Pekeneese disguised as pit-bull, or is it the other way round. You scare me silly, but I luvz ya anywayz... ", "OK")
        Case Right$(Typed, 7) = "CLAUDIA": Call ZenMB("Rowdy Claudie! Damn near as close to kool as frostbite.Rrrrrr. Big lurv...", "OK")
        Case Right$(Typed, 7) = "JESSICA": Call ZenMB("Jump up jump up Jessie! (and get down..) Ta for the big heart. Much love to ya...", "OK")
        Case Right$(Typed, 7) = "STEWART": Call ZenMB("Who? Looks like an overgrown teddy bear you say? Can only be da bindy-boy!!. Much love", "OK")
        Case Right$(Typed, 7) = "KENDALL": Call ZenMB("Boetman! Kendallsan! Damn, its been way to long.....Big up!", "OK")
        Case Right$(Typed, 5) = "TESSA": Call ZenMB("Hola Tessa! What could have been? Have never talked to anyone like I've talked to you. Will always carry you with me..", "OK")
        Case Right$(Typed, 6) = "JUDITH": Call ZenMB("Hey Jude! Love you kiddo! My first. Damn, how clueless was I...", "OK")
        Case Right$(Typed, 7) = "SPENCER": Call ZenMB("Friggin maniac! What are you doing using a computer? To the man who showed me madness, music and mayhem. Much love...", "OK")
        Case Right$(Typed, 6) = "LOUISA": Call ZenMB("Ahh, my Lou lou. What can I say? You are my love, my Queen, my destiny....", "OK")
        Case Right$(Typed, 6) = "MAGGIE": Call ZenMB("Hey sista! You ARE my sister. And you deserve the world! Love ya kiddo...", "OK")
    End Select
    

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Set_Focus(Nothing)
End Sub




Private Sub imiAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Set_Focus(lblAbout)
End Sub







Private Sub imiGroup_Click()
    Call imiGroupOpen_Click
End Sub

Private Sub imiHelp_Click()
    
    'Call Mode_Set("HELP")
    txtMode(3).SetFocus
    Call ShellExe(App.Path & "\Help\Index.htm")
    
End Sub

Private Sub imiHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Set_Focus(lblHelp)
End Sub


Private Sub imiItem_Click()
Dim zTest As clsZenDictionary
    Rem - Add a windows handle for Window commands
    Set zTest = ZKMenu(ZIndex).Copy
    zTest("HWnd") = Me.hwnd
    Call TestAction(zTest)
    
End Sub

Private Sub imiItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Set_Focus(lblItems)
End Sub

Private Sub imiSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Set_Focus(lblSettings)
End Sub






Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Set_Focus(lblAbout)
End Sub














Private Sub lblSetColour_Click()
Dim colNew As OLE_COLOR

    If GetColour(colNew) Then
        booSetChanged = True
        shpSetColour.FillColor = colNew
                
        Dim strItem As String, strSetName As String
        strItem = tvSettings.Nodes(SET_CurIndex).Tag
        strSetName = Prop_Get("SetName", strItem)
        settings(strSetName) = CStr(colNew)
    End If

End Sub

Private Sub lblHelp_Click()

    Call Mode_Set("HELP")
    Call ShellExe(App.Path & "\Help\Index.htm")
    
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Set_Focus(lblHelp)
End Sub


Private Sub mnuClearDead_Click()
Rem - Scans through a group and removes items that no longer point to valid paths...
Dim lngEnd As Long, k As Long
Dim strFile As String
Dim booWork As Boolean


    Rem - Ensure that they do not delete everything
    If ZKMenu(ZIndex)("Class") = "Group" Then
        Rem - Delete the group
        lngEnd = Item_GetGroupEnd(ZIndex)
        For k = ZIndex + 1 To lngEnd
            If ZKMenu(k)("Class") = "File" Then
                strFile = ZKMenu(k)("Action")
                If InStr(strFile, "\") > 0 Then
                    Rem - The file has a path ie. Is not just 'msconfig.exe'
                    Rem - If it is a special folder, insert the folder
                    If InStr(strFile, "%") > 0 Then strFile = InsertSpecialFolder(strFile)
                    On Error Resume Next
                    If Len(Dir(strFile)) = 0 Then
                        Rem - The path is not valid. Remove it.
                        Call Array_Up(k + 1, 1)
                        k = k - 1 ' Move back
                        lngEnd = lngEnd - 1
                        booWork = True
                    End If
                End If
            End If
        Next k
        
        If booWork Then
            booChanged = True
            Rem - Refresh the tree
            Call Tree_Load
            Call Tree_SetFocus(ZIndex)
            Call Item_Selected
        End If
    Else
        Call ZenMB("Sorry, but you can only scan groups for dead items.")
    End If

End Sub

Private Sub mnuHKDel_Click()
    
    ZKMenu(ZIndex)("Hotkey") = ""
    ZKMenu(ZIndex)("Shiftkey") = ""
    booChanged = True
    Call Item_Selected
    Call Tree_SetFocus(ZIndex)

End Sub

Private Sub mnuHKDelAll_Click()
    
    If 1 = ZenMB("You are about to remove all the Hotkeys. Are you sure?", "Yes", "No") Then Exit Sub

    Dim k As Long, max As Long, strProp As String

    max = UBound(ZKMenu())
    strProp = tvTree.SelectedItem.key
    
    For k = 1 To max
        If Len(ZKMenu(k)("Hotkey")) > 0 Then ZKMenu(k)("Hotkey") = ""
        If Len(ZKMenu(k)("Shiftkey")) > 0 Then ZKMenu(k)("ShiftKey") = ""
    Next k
    
    Call Tree_Load
    Call Tree_SetFocus(Prop_Get("Index", strProp))
    Call Item_Selected
    
    booChanged = True

End Sub





Private Sub mnuMoveDown_Click()
    
    If tvTree.SelectedItem.Index < tvTree.Nodes.Count Then
        Call Move_Selected(True)
        Set tvTree.SelectedItem = tvTree.Nodes(tvTree.SelectedItem.Index + 1)
        Call tvTree_Click
    Else
        Call ZenMB("Sorry, but this is already the last item.", "OK")
    End If

End Sub

Private Sub mnuMoveUp_Click()
    
    If tvTree.SelectedItem.Index > 1 Then
        Call Move_Selected(False)
        Set tvTree.SelectedItem = tvTree.Nodes(tvTree.SelectedItem.Index - 1)
        Call tvTree_Click
    Else
        Call ZenMB("Sorry, but this is already the first item.", "OK")
    End If
    
End Sub

Private Sub tmrEdit_Timer()

    tmrEdit.Enabled = False
    If tmrEdit.Tag = "Edit" Then
        Call zbEdit_Click
    Else
        Call zbNew_Click
    End If
    
    
End Sub

Private Sub tvSettings_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then txtMode(1).SetFocus
End Sub

Private Sub tvSettings_NodeClick(ByVal Node As MSComctlLib.Node)

    SET_CurIndex = Node.Index
    Call Set_Selected

End Sub


Private Sub tvTree_DragDrop(Source As Control, X As Single, Y As Single)
    #If Dev = 1 Then
        If tvTree.DropHighlight Is Nothing Then
            Set tvTree.DropHighlight = Nothing
            booDrag = False
            Exit Sub
        Else
            If nodDrag = tvTree.DropHighlight Then Exit Sub
            'Call MsgBox(nodDrag.Text & " dropped on " & tvTree.DropHighlight.Text)
            Dim strSource As String, strDest As String
            strSource = nodDrag.key
            strDest = tvTree.DropHighlight.key
            tvTree.Nodes.Remove strSource
            tvTree.Nodes.Add strDest, tvwPrevious, strSource, nodDrag.Text, nodDrag.Image, nodDrag.SelectedImage
            
            
    '        Dim lngSource As Long, lngDest As Long
    '        lngSource = nodDrag.Index
    '        lngDest = tvTree.DropHighlight.Index
    '        tvTree.Nodes.Remove lngSource
    '        tvTree.Nodes.Add lngDest, tvwPrevious, , nodDrag.Text, nodDrag.Image, nodDrag.SelectedImage
            
            Set tvTree.DropHighlight = Nothing
            booDrag = False
        End If
    #End If
End Sub

Private Sub tvTree_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    #If Dev = 1 Then
        If booDrag Then
            Rem - Set DropHighlight to the mouse's coordinates.
            Set tvTree.DropHighlight = tvTree.HitTest(X, Y)
        End If
    #End If
End Sub


Private Sub tvTree_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 93 Then
        booRightClick = True
        Call tvTree_Click
        booRightClick = False
    ElseIf KeyCode = vbKeyEscape Then
        txtMode(0).SetFocus
    End If
End Sub

Private Sub tvTree_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call tvTree_DblClick

End Sub


Private Sub tvTree_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    #If Dev = 1 Then
        If Button = vbLeftButton Then ' Signal a Drag operation.
            booDrag = True ' Set the flag to true.
            With tvTree
                Rem - Set the drag icon with the CreateDragImage method.
                .DragIcon = .SelectedItem.CreateDragImage
                .Drag vbBeginDrag
            End With
        End If
    #End If
End Sub

Private Sub txtMode_Change(Index As Integer)
    txtMode(Index).Text = ""
End Sub

Private Sub txtMode_GotFocus(Index As Integer)
    
    shpFocus.BorderColor = RGB(100, 200, 100)
    Select Case Index
        Case 0
            Call Mode_Set("ITEMS")
        Case 1
            Call Mode_Set("SETTINGS")
        Case 2
            Call Mode_Set("ABOUT")
        Case 3
            Call Mode_Set("HELP")
    End Select
    
End Sub


Private Sub txtMode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Dim intIndex As Integer
    Select Case KeyCode
        Case vbKeyUp
            If Index = 0 Then intIndex = txtMode.UBound Else intIndex = Index - 1
        Case vbKeyDown
            If Index = txtMode.UBound Then intIndex = 0 Else intIndex = Index + 1
        Case vbKeyRight, vbKeyReturn, vbKeySpace
            intIndex = -1
            Select Case ModeLabel.Name
                Case "lblItems"
                    tvTree.SetFocus
                Case "lblSettings"
                    tvSettings.SetFocus
                Case "lblAbout"
                    
                Case "lblHelp"
                    Call ShellExe(App.Path & "\Help\Index.htm")
            End Select
        
    End Select
    If intIndex >= 0 Then txtMode(intIndex).SetFocus
    
End Sub

Private Sub txtMode_LostFocus(Index As Integer)
    shpFocus.BorderColor = COL_Zen
End Sub


Private Sub txtSet_LostFocus()

    If Not booLoading Then Call Set_SetValue(txtSet.Tag)
    
End Sub



Private Sub zbBackup_Click()
Dim strFolder As String
    
    Select Case zbBackup.Tag
        Case "Backup"
            If FBR_BrowseForFolder("Backup folder", strFolder) Then
                If Len(Dir(strFolder & "\ZenKEY.ini")) > 0 Then
                    If ZenMB("This folder already contains a backup. Are you sure you wish to overwrite this backup?", "Yes", "No") = 1 Then Exit Sub
                End If
                On Error Resume Next
                Call FileCopy(settings("SavePath") & "\ZenKEY.ini", strFolder & "\ZenKEY.ini")
                Call FileCopy(settings("SavePath") & "\Settings.ini", strFolder & "\Settings.ini")
                If Err.Number = 0 Then
                    Call ZenMB("Menu and settings successfully backed up.")
                Else
                    Call ZenMB("Unable to backup to this folder. Please make sure you have write permissions to the folder " & strFolder)
                End If
                Err.Clear
            End If
        Case "Restore"
            If FBR_BrowseForFolder("Restore folder", strFolder) Then
                If Len(Dir(strFolder & "\ZenKEY.ini")) > 0 Then
                    If ZenMB("Restoring this backup will overwrite your current menu and settings. Are you sure you wish to restore this backup?", "Yes", "No") = 1 Then Exit Sub
                Else
                    Call ZenMB("This folder does not contain a valid backup.", "OK")
                    Exit Sub
                End If
                On Error Resume Next
                Call FileCopy(strFolder & "\ZenKEY.ini", settings("SavePath") & "\ZenKEY.ini")
                Call FileCopy(strFolder & "\Settings.ini", settings("SavePath") & "\Settings.ini")
                If Err.Number = 0 Then
                    Dim lngHandle As Long
                    lngHandle = Val(Registry.GetRegistry(HKCU, "SOFTWARE\ZenCODE\ZenKEY", "WindowHandle"))
                    If lngHandle <> 0 Then
                        If 0 = ZenMB("Menu and settings successfully restored. The Configuration utility will now close and ZenKEY needs to be restarted. Restart ZenKEY now", "Yes", "No") Then Call ZK_Restart
                    Else
                        Call ZenMB("Menu and settings successfully restored. The Configuration utility will now close.", "OK")
                    End If
                    Unload Me
                Else
                    Call ZenMB("The selected folder does not appear to contain a valid backup. (" & strFolder & ")")
                End If
                Err.Clear
            End If
        Case "ShowHotkeys"
            Call ShowHotkeys
    End Select
    
End Sub

Private Sub zbDel_Click()
    Call PopupMenu(mnuDel)
End Sub




Private Sub zbEdit_Click()
    
    If Item_Edit(ZKMenu(ZIndex)) Then
        booChanged = True
        tvTree.SelectedItem.Text = ZKMenu(ZIndex)("Caption")
        tvTree.SelectedItem.Checked = Not (ZKMenu(ZIndex)("Disabled") = "True")
        Call Item_Selected
        Call Tree_SetFocus(ZIndex)
    End If
        
End Sub
















































Private Sub imiGroupOpen_Click()

    tvTree.SelectedItem.Expanded = Not tvTree.SelectedItem.Expanded
    Call Item_Selected
    
End Sub

Private Sub lblItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Set_Focus(lblItems)
End Sub












Private Sub zbSetFileAdd_Click()
Dim strExe As String, booSelected As Boolean

    If SelectFileDlg(Me, strExe, zbSetFileAdd.Tag = "ClassList") Then
        strExe = GetFileName(strExe)
        Call Filelist_Update(strExe, True)
        Call Set_Selected
        booSetChanged = True
    End If
    
    
End Sub

Private Sub zbSetFileRemove_Click()
    
    With cmbSet
        If (.ListCount > 1) Or (.List(0) <> List_Null) Then
            Call Filelist_Update(cmbSet.Text, False)
            Call Set_Selected
            booSetChanged = True
        End If
    End With

End Sub

Private Sub zbMove_Click()
    
    mnuNew.Visible = False
    mnuEdit.Visible = False
    mnuSep.Visible = False
    mnuGroup.Visible = False
    Call PopupMenu(mnuMov)
    mnuNew.Visible = True
    mnuEdit.Visible = True
    mnuSep.Visible = True
    mnuGroup.Visible = True

End Sub

Private Sub zbNew_Click()
Dim zdAct As clsZenDictionary

    Set zdAct = New clsZenDictionary
    If Item_Edit(zdAct) Then
        booChanged = True
                
        If zdAct("Class") = "Group" Then
            Rem - Adding a group
            Call Array_Down(ZIndex, 3)
            Set ZKMenu(ZIndex) = zdAct.Copy
            Set ZKMenu(ZIndex + 1) = zenDic("Class", "ZenKEY", "Action", "About", "Caption", "Item in new group")
            Set ZKMenu(ZIndex + 2) = zenDic("ENDGROUP", "True")
            
        Else
            Rem - Adding a singular item
            Call Array_Down(ZIndex, 1)
            Set ZKMenu(ZIndex) = zdAct.Copy
        End If
        
        Call Tree_Load
        Call Tree_SetFocus(ZIndex)
        Call Item_Selected
    End If
    
End Sub




Private Sub lblSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Set_Focus(lblSettings)
End Sub















Private Sub imiAbout_Click()
    'Call Mode_Set("ABOUT")
    txtMode(2).SetFocus
End Sub


Private Sub imiItems_Click()
    'Call Mode_Set("ITEMS")
    txtMode(0).SetFocus
End Sub

Private Sub imiSettings_Click()
    'Call Mode_Set("SETTINGS")
    txtMode(1).SetFocus
End Sub

Private Sub lblAbout_Click()

    Call Mode_Set("ABOUT")
End Sub





























Private Sub zbExit_Click()
    
    If UnLoadMe Then Unload Me
    
End Sub




Private Sub lblItems_Click()
    Call Mode_Set("ITEMS")
End Sub









Private Sub zbQuoteNow_Click()
Dim booMore As Boolean, strPrev As String
Dim strPrevQuotes As String

    strPrev = settings("HideQuotes")
    strPrevQuotes = settings("Quotes")
    settings("HideQuotes") = "False"
    settings("Quotes") = cmbSet.Text
    Do
        booMore = CBool(1 = ZenMB(ZenKEYCap, "OK", "More"))
    Loop While booMore
    settings("HideQuotes") = strPrev
    settings("Quotes") = strPrevQuotes
    
End Sub

Private Sub zbSave_Click()
    
    If booChanged Then Call SaveToINI
    If booSetChanged Then Call Set_Save(True)
    If booChanged Or booSetChanged Then
        Dim lngHandle As Long
        lngHandle = Val(Registry.GetRegistry(HKCU, "SOFTWARE\ZenCODE\ZenKEY", "WindowHandle"))
        If lngHandle <> 0 Then
            If 0 = ZenMB("The changes have been saved. An instance of ZenKEY is active and needs to be restarted for the changes to take effect. Would you like to restart ZenKEY now?", "Yes", "No") Then Call ZK_Restart
        Else
            Call ZenMB("Settings saved.", "OK")
        End If
    End If
    booChanged = False
    booSetChanged = False

End Sub

Public Sub Display()
    
    booLoading = True
    Call Menu_LoadINI("ZenKEY.ini", True)
    Call SetGraphics
    Call CentreForm(Me)
    Set ModeLabel = Nothing
    
    Rem - Initialise the colour array
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    Dim i As Integer
    For i = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(i) = 0
    Next i
    
    
    Me.AutoRedraw = False
    Set objSelected = Nothing
    booLoading = False
    booChanged = False
    Me.Show
    If txtMode(2).Visible Then txtMode(2).SetFocus

    Select Case Command$
        Case "SETTINGS"
            Call Mode_Set("SETTINGS")
            Call imiSettings_Click
        Case "ITEMS"
            Call Mode_Set("ITEMS")
            Call imiItems_Click
        Case "MAP"
            Call Mode_Set("SETTINGS")
            Call imiSettings_Click
            Dim k As Long
            DoEvents
            With tvSettings
                For k = .Nodes.Count To 1 Step -1
                    If .Nodes(k).Text = "Desktop map" Then
                        .SelectedItem = .Nodes(k)
                        SET_CurIndex = k
                        Call Set_Selected
                        .SelectedItem.Expanded = True
                        Exit For
                    End If
                Next k
            End With
    End Select
    
End Sub






Private Function Dynalist_Load(ByVal ListName As String) As String
Dim strName As String, strTemp As String
Dim k As Long

    k = 1
    Select Case ListName
        Case "Skin": strName = Dir(App.Path & "\Skins\*.ico", vbNormal)
        Case "Quotes": strName = Dir(App.Path & "\Quotes\*.txt", vbNormal)
        Case "TransActive", "TransInactive", "SET_Trans"
            Call Prop_Set("Item1", "Opaque", strTemp)
            Call Prop_Set("Val1", "Opaque", strTemp)
            For k = 1 To 10
                Call Prop_Set("Item" & CStr(12 - k), CStr(10 * k) & "%", strTemp)
                Call Prop_Set("Val" & CStr(12 - k), CStr(10 * k) & "%", strTemp)
            Next k
            
            Select Case ListName
                Case "TransActive"
                    Call Prop_Set("Default", "-1", strTemp) ' Default opaque
                Case "TransInactive"
                    Call Prop_Set("Default", "70", strTemp) ' Default 20% for inactive
                Case "SET_Trans"
                    Call Prop_Set("Default", "Opaque", strTemp) ' Default opaque
            End Select
            
            strName = vbNullString ' Skip loading in step below
    End Select
    
    While Len(strName) > 0
        Call Prop_Set("Item" & CStr(k), left$(strName, Len(strName) - 4), strTemp)
        strName = Dir
        k = k + 1
    Wend
    Dynalist_Load = strTemp

End Function








Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode <> vbFormCode Then
        If Not UnLoadMe Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    Dim k As Integer
    For k = Forms.Count - 1 To 0 Step -1
        Unload Forms(k)
    Next k
    
End Sub


Private Sub lblSettings_Click()
    Call Mode_Set("SETTINGS")
End Sub




Private Sub SaveToINI()
Dim intFNum As Integer
Dim k As Integer
Dim intItemMax As Integer
    
    Rem ======================     Open file and start loading
    Rem - Initialise variables
    intFNum = FreeFile
    intItemMax = UBound(ZKMenu())
    
    Open settings("SavePath") & "\Zenkey.ini" For Output As #intFNum
            Print #intFNum, "====================== Zenkey initialisation file ======================"
            'For intSection = 0 To SecIndexes
            For k = 2 To intItemMax - 2
                Rem - Initialise for a new section
                If ZKMenu(k)("Class") = "Group" Then Print #intFNum, "====================== " & ZKMenu(k)("Caption") & " ======================"
                Print #intFNum, ZKMenu(k).ToProp
            Next k
    Close #intFNum

End Sub


Private Sub SetGraphics()
Dim StartY As Single, Y As Single
Dim COL_Surround As OLE_COLOR
    
    COL_Surround = RGB(220, 220, 255)


    With Me
        .AutoRedraw = True
        .Width = 9500
        .Height = 4650 '5100
        Call TileMe(Me, LoadPicture(App.Path & "\Help\cloudsdark.jpg"))
        Set .Picture = .Image
        .AutoRedraw = False
    End With
    lblVersion.Caption = "V " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    
    Rem ===================== Moving ===========================
    Rem - Layout the item properties labels
    
    Y = 19
    StartY = 50
    chkEnabled.Top = StartY
    lblCaption.Top = StartY
    lblActionType.Top = StartY + Y
    lblAction.Top = StartY + 2 * Y
    lblHotkey.Top = StartY + 3 * Y
    lblParam.Top = StartY + 4 * Y
    lblAdd.Top = StartY + 5 * Y
    chkRClick.Top = StartY + 6 * Y
    
    picNewSettings.Move 68, 10 ', 550, 234
    picItems.Move 68, 10 ', 550, 234
    imiItem.Move imiGroup.left, imiGroup.Top
    imiGroupOpen.Move imiGroup.left, imiGroup.Top
    
    Rem ===================== Colour and graphics ===========================
    Rem - Picture box for backgrounds
    With picNewSettings
        .AutoRedraw = True
        Call .PaintPicture(Me.Picture, 0, 0, .Width, .Height, .left, .Top, .Width, .Height)
        Set .Picture = .Image
       .AutoRedraw = False
    End With
    With picItems
       .AutoRedraw = True
        Call .PaintPicture(Me.Picture, 0, 0, .Width, .Height, .left, .Top, .Width, .Height)
        Set .Picture = .Image
        .AutoRedraw = False
    End With
    
    imiZKMenu.Top = 1
    imiZKProp.Top = 1
    imiZKNewSetting.Top = 1
    imiZKSettings.Top = 1
    
    shpFocus.BorderColor = COL_Zen 'RGB(80, 80, 255)
    Me.AutoRedraw = True
    Me.Font.Size = 26
    Call PrintZen(3, 2, COL_Surround)
    Call PrintZen(-3, 2, COL_Surround)
    Call PrintZen(-3, --2, COL_Surround)
    Call PrintZen(3, -2, COL_Surround)
    Call PrintZen(0, 0, RGB(76, 166, 203))
    'Call PrintZen(0, 0, vbBlack)
    
    Dim It As Control
    For Each It In Me.Controls
        If TypeOf It Is Label Or TypeOf It Is CheckBox Or TypeOf It Is OptionButton Then
            If It.Font.Bold Then
                It.ForeColor = COL_Zen
            Else
                It.ForeColor = vbBlack
            End If
            It.BackColor = vbWhite
        ElseIf TypeOf It Is Shape Then
            It.FillColor = vbWhite
            
        End If
    Next It
    
    Rem - Set the grahics for the larger buttons
    Set imiZKProp.Picture = imiZKMenu.Picture
    Set imiZKNewSetting.Picture = imiZKMenu.Picture
    Set imiZKSettings.Picture = imiZKMenu.Picture
    Set zbExit.Picture = zbSave.Picture
    
    Rem - Set the graphics for the smaller buttons
    Set zbEdit.Picture = zbNew.Picture
    Set zbDel.Picture = zbNew.Picture
    Set zbMove.Picture = zbNew.Picture
    Set zbQuoteNow.Picture = zbNew.Picture
    Set zbSetFileAdd.Picture = zbNew.Picture
    Set zbSetFileRemove.Picture = zbNew.Picture
        
    Rem - Fill the buttons
    'zbSave.BackColor = RGB(220, 235, 255)
    'zbExit.BackColor = zbSave.BackColor
    'zbSelItem.BackColor = RGB(240, 245, 255)
    'zbMenu.BackColor = zbSelItem.BackColor
    'zbNewSetting.BackColor = zbSelItem.BackColor
    'zbNewSettings.BackColor = zbSelItem.BackColor


    Rem - Floodfills. Just looks to fancy.
'    Call zbSave.FloodBack(RGB(220, 235, 255), vbWhite, 1)
'    Set zbExit.Picture = zbSave.Picture
'
'    Call zbNew.FloodBack(RGB(220, 235, 255), vbWhite, 1)
'    Set zbEdit.Picture = zbNew.Picture
'    Set zbDel.Picture = zbNew.Picture
'    Set zbMove.Picture = zbNew.Picture
    
'    Call zbSelItem.FloodBack(RGB(180, 210, 255), vbWhite, 0)
'    Set zbMenu.Picture = zbSelItem.Picture
'    Set zbNewSetting.Picture = zbSelItem.Picture
'    Set zbNewSettings.Picture = zbSelItem.Picture
    
End Sub





Private Function UnLoadMe() As Boolean
    If booChanged Or booSetChanged Then
        UnLoadMe = CBool(0 = ZenMB("You have made changes. Are you sure you wish to exit without saving these changes?", "Yes", "No"))
    Else
        UnLoadMe = True
    End If
End Function

Private Sub Set_Focus(Optional ByRef DaLabel As Control = Nothing)
Dim colSelected As Long  '&HC000&

    colSelected = RGB(0, 155, 0) 'vbBlue
    If Not DaLabel Is Nothing Then
        If DaLabel = ModeLabel Then Exit Sub
        
        If Not objSelected Is Nothing Then
            If ModeLabel <> objSelected Then objSelected.ForeColor = COL_Zen
        End If
        
        Set objSelected = DaLabel
        objSelected.ForeColor = colSelected '
    
    ElseIf DaLabel Is Nothing Then
    
        If Not objSelected Is Nothing Then
            If objSelected <> ModeLabel Then
                objSelected.ForeColor = COL_Zen
                Set objSelected = Nothing
            End If
        End If
    End If
        
End Sub












Private Sub mnuCollapse_Click()
    
    tvTree.SelectedItem.Expanded = False
    Call Item_Selected

End Sub

Private Sub mnuDelete_Click()
Dim lngParent As Long
Dim lngEnd As Long
Dim booGroup As Boolean

    Rem - Ensure that they do not delete everything
    booGroup = CBool(ZKMenu(ZIndex)("Class") = "Group")
        
    If ZIndex = 2 Then ' The first item is being deleted
        If booGroup Then
            lngEnd = Item_GetGroupEnd(ZIndex)
            If lngEnd > UBound(ZKMenu()) - 3 Then
                Call ZenMB("Sorry, but you cannot delete the last item in ZenKEY. Otherwise, why bother, really!", "OK")
                Exit Sub
            End If
        ElseIf UBound(ZKMenu()) < 5 Then
            Call ZenMB("Sorry, but you cannot delete the last item in ZenKEY. Otherwise, why bother, really!", "OK")
            Exit Sub
        End If
    End If
    
    If ZKMenu(ZIndex)("Class") = "Group" Then
        If ZenMB("You are deleting a group, which will delete all the items inside the group. Are you sure you wish to do this?", "Yes", "No") = 0 Then
            Rem - Delete the group
            lngEnd = Item_GetGroupEnd(ZIndex)
            Rem - Deletre Group
            Call Array_Up(lngEnd + 1, lngEnd - ZIndex + 1)
            Rem - Refresh the tree
            Call Tree_Load
            Call Tree_SetFocus(ZIndex)
            booChanged = True
        End If
    Else
        Rem - Delete the item if not the only item in the group
        
        lngParent = Item_GetGroup(ZIndex)
        lngEnd = Item_GetGroupEnd(lngParent)
        
        If lngEnd - lngParent < 3 Then
            Rem - It is the last item
            Call ZenMB("You cannot delete the last item in a group. Rather just delete the group itself?", "OK")
        Else
            Rem - Delete Item
            Call Array_Up(ZIndex + 1, 1)
            Rem - Refresh the tree
            Call Tree_Load
            Call Tree_SetFocus(ZIndex)
            booChanged = True
        End If
        
    End If
    Call Item_Selected

End Sub

Private Sub mnuEdit_Click()
    
    Rem - Calling the edit sub here prevents the menus from being shown in the edit dialog
    tmrEdit.Tag = "Edit"
    tmrEdit.Enabled = True
    
End Sub

Private Sub mnuExpand_Click()
    
    tvTree.SelectedItem.Expanded = True
    Call Item_Selected

End Sub

Private Sub mnuMoveAfter_Click()

    Call Move_Selected(True)

End Sub

Private Sub mnuMoveBefore_Click()
    Call Move_Selected(False)
End Sub


Private Sub mnuNew_Click()

    Rem - Calling the edit sub here prevents the menus from being shown in the edit dialog
    tmrEdit.Tag = "New"
    tmrEdit.Enabled = True
    
    
End Sub

Private Sub mnuSortAll_Click()

    If ZenMB("This will sort all the items in all the groups. Are you sure?", "Yes", "No") = 0 Then
        Dim k As Long, strProp As String
            
        
        With tvTree
            Dim lngIndex As Long
            lngIndex = CLng(Prop_Get("Index", .SelectedItem.key))
            .Sorted = True
            For k = .Nodes.Count To 1 Step -1
                If .Nodes(k).Children > 0 Then .Nodes(k).Sorted = True
            Next k
            .Sorted = False
            
            Call Tree_ReadFrom
            Call Tree_Load
            Call Tree_SetFocus(lngIndex)
            booChanged = True
        End With
    End If
    
End Sub

Private Sub mnuSortGroup_Click()
Dim lngStart As Long
    
    'tvTree.SelectedItem.Sorted = True
    lngStart = Prop_Get("Index", tvTree.SelectedItem.key)
    If ZKMenu(lngStart)("Class") <> "Group" Then
        Call ZenMB("Only groups can be sorted. Please select a group to sort first.", "OK")
    Else
        tvTree.SelectedItem.Sorted = True
        Call Tree_ReadFrom
        Call Tree_Load
        Call Tree_SetFocus(lngStart)
        booChanged = True
    End If

End Sub











Public Sub Mode_Set(ByVal TheMode As String)

    Screen.MousePointer = vbHourglass
    
    TheMode = UCase(TheMode)
    If Not ModeLabel Is Nothing Then
        ModeLabel.ForeColor = COL_Zen ' vbBlue
        ModeLabel.Move ModeLabel.left - 2, ModeLabel.Top - 2
    End If
    Select Case TheMode
        Case "ITEMS"
            Set ModeLabel = lblItems
            If lblItems.Tag <> "Dirty" Then
                Rem - Load the stuff here to speed up the initial load
                lblItems.Tag = "Dirty"
                booLoading = True
                Set tvTree.ImageList = imlTree
                Call Tree_Load
                booLoading = False
            End If
            
            If tvTree.SelectedItem Is Nothing Then Set tvTree.SelectedItem = tvTree.Nodes(1)
        Case "SETTINGS"
            Set ModeLabel = lblSettings
            If lblSettings.Tag <> "Dirty" Then
                Rem - Load the stuff here to speed up the initial load
                booLoading = True
                lblSettings.Tag = "Dirty"
                Set tvSettings.ImageList = imlSettings
                Call Set_Init
                booLoading = False
            End If
        Case "ABOUT"
            Set ModeLabel = lblAbout
        Case "HELP"
            Set ModeLabel = lblHelp
            'Call ShellExe(App.Path & "\Help\Index.htm")
    End Select
    ModeLabel.ForeColor = COL_Zen '&H00808000& 'RGB(0, 150, 0)  'vbCyan
    ModeLabel.Move ModeLabel.left + 2, ModeLabel.Top + 2
    
    picNewSettings.Visible = CBool((TheMode = "SETTINGS"))
    If picNewSettings.Visible Then
        If SET_CurIndex = 0 Then SET_CurIndex = 1
        Call Set_Selected
        tvSettings.SelectedItem.Expanded = False
    End If
    
    picItems.Visible = CBool((TheMode = "ITEMS"))
    If picItems.Visible Then Call Item_Selected
    'shpFocus.Top = ModeLabel.Top - 48
    shpFocus.Width = imiAbout.Width * 1.31
    Select Case TheMode
        Case "ITEMS"
            shpFocus.Top = imiItems.Top - 4
        Case "SETTINGS"
            shpFocus.Top = imiSettings.Top - 4
        Case "ABOUT"
            shpFocus.Top = imiAbout.Top - 4
        Case "Help"
            shpFocus.Top = imiHelp.Top - 4
    End Select
    
    
    Screen.MousePointer = vbDefault
    
End Sub













Private Sub picItems_Click()
    If booMoving Then Call Move_Selected(booMoveAfter)
End Sub

Private Sub picItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Set_Focus(Nothing)
End Sub



Private Sub tvTree_Click()
  
    If booRightClick Then
        If Not booMoving Then
            If ZKMenu(ZIndex)("Class") = "Group" Then
                mnuGroup.Visible = True
                mnuExpand.Visible = Not tvTree.SelectedItem.Expanded
                mnuCollapse.Visible = Not mnuExpand.Visible
                
            Else
                mnuGroup.Visible = False
            End If
            mnuDel.Visible = True
            Call PopupMenu(mnuMov)
            mnuDel.Visible = False
            Exit Sub
        End If
    End If

    If booMoving Then
        tvTree.Nodes(lngSource).SelectedImage = tvTree.Nodes(lngSource).Image
                
        Dim Source As Long, dest As Long
        Dim prop As clsZenDictionary
        
        Source = Val(Prop_Get("Index", tvTree.Nodes(lngSource).key))
        dest = Val(Prop_Get("Index", tvTree.SelectedItem.key))
        
        Rem - Check that they are allowed to move the item
        Dim lngGroup As Long
        lngGroup = Item_GetGroup(Source)
        If Group_ItemCount(lngGroup) < 2 Then
            Call ZenMB("Sorry, but you cannot move the last item in a group.", "OK")
        Else
            Rem - Okay, move it!
            If (ZKMenu(dest)("Class") = "Group") And booMoveAfter Then dest = Item_GetGroupEnd(dest)
            Set prop = ZKMenu(Source).Copy
            Call Array_Move(Source, dest, booMoveAfter)
            Call Tree_Load
            Call Tree_SetFocus(Source)
        End If
        
        Call Move_Selected(booMoveAfter)
        booChanged = True
    End If
    
End Sub

Private Sub tvTree_Collapse(ByVal Node As MSComctlLib.Node)
    
    If Not (tvTree.SelectedItem Is Nothing) Then
        If Node.Index = tvTree.SelectedItem.Index Then Call Item_Selected
    End If
    

End Sub


Private Sub tvTree_DblClick()
    If ZKMenu(ZIndex)("Class") <> "Group" Then Call zbEdit_Click
End Sub

Private Sub tvTree_Expand(ByVal Node As MSComctlLib.Node)
    
    imiGroup.Visible = False
    imiGroupOpen.Visible = True

End Sub


Private Sub tvTree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    booRightClick = CBool(Button = 2)
    #If Dev = 1 Then
        Set nodDrag = tvTree.HitTest(X, Y)
    #End If
    
End Sub

Private Sub tvTree_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Call Item_Selected
End Sub


Public Sub Tree_Load()
Dim k As Long
Dim max As Long
Dim nodX As Node
Dim strCap As String
Dim nParent As Node

    max = UBound(ZKMenu()) - 2
    Set nParent = Nothing
    
    With tvTree
        .Visible = False
        .Nodes.Clear
        For k = 2 To max
            Rem - Add the group item
            If ZKMenu(k)("Class") = "Group" Then
                Rem - Start a new group
                k = Tree_Load_Group(-1, k)
            Else
                Rem - Just add to the current group
                Set nodX = .Nodes.Add(, tvwLast, "|Index=" & CStr(k) & "|", ZKMenu(k)("Caption"))
                nodX.ForeColor = IIf(ZKMenu(k)("Disabled") = "True", ColDisabled, vbBlack)
                'nodX.EnsureVisible
                nodX.Image = "Action"
            End If
        Next k
        .Visible = True
    End With
    
End Sub

Public Sub Item_Selected()
Dim strTemp As String
Dim booGroup As Boolean
    
    booLoading = True
    Rem - Set the group and item no accordingly
    ZIndex = CLng(Prop_Get("Index", tvTree.SelectedItem.key))
    booGroup = CBool(ZKMenu(ZIndex)("Class") = "Group")

    lblAdd.Visible = False
    chkRClick.Visible = booGroup
    imiItem.Visible = Not booGroup
    If booGroup Then
        imiGroup.Visible = Not CBool(tvTree.SelectedItem.Expanded)
        imiGroupOpen.Visible = Not imiGroup.Visible
    Else
        imiGroup.Visible = False
        imiGroupOpen.Visible = False
    End If
    
    lblParam.Visible = False 'booGroup
    lblParam.Caption = vbNullString
    If booGroup Then chkRClick.Value = IIf(ZKMenu(ZIndex)("RIGHTCLICKMENU") = "True", 1, 0)
    
    If booGroup Then
        Rem - They clicked on a group
        lblActionType.Caption = "Type : " & Actions_GetDescrip(zenDic("Class", "Group"))
        lblAction.Caption = "Action : Opens a ZenKEY Menu" 'Action type : ZenKEY Menu"
    Else
        Rem - They clicked on an item
        
        Select Case ZKMenu(ZIndex)("Class")
            Case "File"
                Dim strFile As String
                If ZKMenu(ZIndex)("Action") = "rundll32.exe" Then
                    Rem - A control panel applet
                    lblActionType.Caption = "Action type : Control panel applet"
                    lblAction.Caption = "Applet : " & Mid$(ZKMenu(ZIndex)("Param"), 28)
                Else
                    Rem - A normal file
                    lblActionType.Caption = "Action type : File open / run"
                    strTemp = ZKMenu(ZIndex)("Action")
                    If InStr(strTemp, "?") > 0 Then strTemp = left$(strTemp, InStr(strTemp, "?") - 1) ' Strip out the parameter if applicable
                    lblAction.Caption = "File name : " & GetFileName(strTemp)
                    If Len(ZKMenu(ZIndex)("Param")) > 0 Then
                        strTemp = ZKMenu(ZIndex)("Param")
                        If Len(strTemp) > 10 Then strTemp = "Parameter : ..." & Mid$(strTemp, Len(strTemp) - 10) Else strTemp = "Parameter : " & strTemp
                        lblParam.Caption = strTemp
                    End If
                    If ZKMenu(ZIndex)("NewInstance") = "True" Then
                        strTemp = "New instance, "
                    Else
                        strTemp = "Activate, "
                    End If

                    Rem - Values for ChangeDir
                    Rem - If ChangeDir = "No" - Stay in current dir
                    Rem - If InStr(ChangeDir, "\") > 0 - Changes to the specified dir
                    Rem - Else changes to App dir
                    Dim strDir As String
                    strDir = ZKMenu(ZIndex)("ChangeDir")
                    If strDir = "No" Then
                        Rem - Stay in current folder
                        strTemp = strTemp & "start in current folder"
                    Else
                        If InStr(strDir, "\") > 0 Then
                            Rem - Changes to the specified dir
                            strTemp = strTemp & "start in custom folder"
                        Else
                            Rem - Else changes to App dir
                            strTemp = strTemp & "start in app's folder"
                        End If
                    End If
                    lblAdd.Caption = strTemp
                    lblAdd.Visible = True
                    
                    'Select Case Prop_Get("StartUp", ZKMenu(ZIndex))
                    '    Case "Min": lblParam.Caption = "When ZenKEY starts, launch minimized"
                    '    Case "Max": lblParam.Caption = "When ZenKEY starts, launch maximized"
                    '    Case "Normal": lblParam.Caption = "When ZenKEY starts, launch normally"
                    '    Case Else: lblParam.Caption = vbNullString
                    'End Select
                    
                End If
            Case "Winamp"
                lblActionType.Caption = "Action type : Winamp instruction"
                lblAction.Caption = "Action : " & Actions_GetDescrip(ZKMenu(ZIndex))
            Case "Media"
                lblActionType.Caption = "Action type : Windows Media command"
                lblAction.Caption = "Action : " & Actions_GetDescrip(ZKMenu(ZIndex))
                Select Case ZKMenu(ZIndex)("Window Class")
                    Case "Active", vbNullString
                        lblParam.Caption = "Target : Active"
                    Case "Sonique2 Window Class"
                        lblParam.Caption = "Target : Sonique 2"
                    Case "WMPlayerApp"
                        lblParam.Caption = "Target : Win Media player"
                    Case Else ' User defined window class
                        lblParam.Caption = "Target class: '" & ZKMenu(ZIndex)("Window Class") & "'"
                End Select
            Case "Folder"
                lblActionType.Caption = "Action type : Opens a Folder"
                strTemp = ZKMenu(ZIndex)("Action")
                If InStr(strTemp, "%") > 0 Then strTemp = InsertSpecialFolder(strTemp)
                If Len(strTemp) > 20 Then strTemp = "..." & Right$(strTemp, 20)
                lblAction.Caption = "Folder : " & strTemp
            Case "SpecialFolder"
                lblActionType.Caption = "Action type : Opens a Special Folder"
                strTemp = ZKMenu(ZIndex)("Action")
                lblAction.Caption = "Folder : " & SpecialFolderCaption(Val(Mid(strTemp, 2)))
            Case "URL"
                Dim strPrefix As String
                strTemp = ZKMenu(ZIndex)("Action")
                Select Case True ' ---- First detect the  current type of address
                    Case left$(strTemp, 7) = "http://"
                        Rem - A normal www url
                        strPrefix = "http://"
                    Case left$(strTemp, 8) = "https://"
                        Rem - Case 2 '"Secure Web address (https)"
                        strPrefix = "https://"
                    Case left$(strTemp, 6) = "ftp://"
                        Rem -  ftp Site (ftp)
                        strPrefix = "ftp://"
                End Select
                lblActionType.Caption = "Action type : Open Internet location"
                If Len(strPrefix) > 0 Then
                    strTemp = Mid$(strTemp, (Len(strPrefix) + 1))
                    strTemp = "Action : " & strTemp & " (" & strPrefix & ")"
                Else
                    strTemp = "Action : " & strTemp
                End If
                If Len(strTemp) < 34 Then
                    lblAction.Caption = strTemp
                Else
                    lblAction.Caption = left$(strTemp, 32) & "..."
                End If
            Case "Search"
                strTemp = ZKMenu(ZIndex)("Action")
                If Len(strTemp) < 31 Then
                    lblAction.Caption = "Address : " & strTemp
                Else
                    lblAction.Caption = "Address : ..." & Right(strTemp, 29)
                End If
                lblActionType.Caption = "Action type : Perform internet search"
                'lblAction.Caption = "Action : " & Actions_GetDescrip(ZKMenu(ZIndex)) 'Prop_Get("Action", ZKMenu(ZIndex))
            Case "Keystrokes"
                strTemp = KS_GetDescription(ZKMenu(ZIndex)("Action"))
                If Len(strTemp) < 27 Then
                    lblAction.Caption = "Key sequence : " & strTemp
                Else
                    lblAction.Caption = "Key sequence : " & left(strTemp, 26) & "..."
                End If
                lblActionType.Caption = "Action type : Simulate a series of Keystrokes"
            Case Else
                lblActionType.Caption = "Action type : " & ZKMenu(ZIndex)("Class")
                lblAction.Caption = "Action : " & Actions_GetDescrip(ZKMenu(ZIndex)) 'Prop_Get("Action", ZKMenu(ZIndex))
        End Select
            
        Rem - Startup
        strTemp = ZKMenu(ZIndex)("StartUp")
        Select Case True
            Case strTemp = "0": strTemp = "Fire on startup"
            Case IsNumeric(strTemp): strTemp = "Fire " & strTemp & " seconds after startup"
            Case Else: strTemp = vbNullString 'Need this to eliminate legacy "Startup" values
        End Select
        If Len(strTemp) > 0 Then
            If Len(lblParam.Caption) > 0 Then
                lblParam.Caption = lblParam.Caption & ". " & strTemp
            Else
                lblParam.Caption = strTemp
            End If
            lblParam.Visible = True
        End If

    End If
        
    chkEnabled.Value = IIf(ZKMenu(ZIndex)("Disabled") = "True", 0, 1)
    tvTree.SelectedItem.ForeColor = IIf(chkEnabled.Value = 1, vbBlack, ColDisabled)
    
    Rem - Common to items and groups
    strTemp = ZKMenu(ZIndex)("Caption")
    If Len(strTemp) > 22 Then
        lblCaption.Caption = "Caption : " & left$(strTemp, 20) & "..."
    Else
        lblCaption.Caption = "Caption : " & strTemp
    End If
    
    strTemp = HotKeys.GetCaption(ZKMenu(ZIndex))
    If Len(strTemp) < 1 Then
        lblHotkey.Caption = "Hotkey : None"
    Else
        lblHotkey.Caption = "Hotkey : " & strTemp
    End If
    
    booLoading = False
    
End Sub







Private Function Item_Edit(ByRef prop As clsZenDictionary) As Boolean
Dim ActForm As frmAction
    
    Set ActForm = New frmAction
    With ActForm
        Set .CallingForm = Me
        Set .prop = prop.Copy
        .EditIndex = ZIndex
        .Init
        Me.Visible = False
        '.Show vbModal
        .Show
        While Not .booDone
            DoEvents
        Wend
        
        If .booValid Then Set prop = .prop
        Item_Edit = .booValid
        Me.Visible = True
        
    End With
    
    Unload ActForm
    Set ActForm = Nothing


End Function

Public Function Tree_Load_Group(ByVal ParentIndex As Long, ByVal startIndex As Long) As Long
Rem - Returns the Ending index of the group

Dim k As Long
Dim max As Long
Dim nodX As Node

    Rem - Start a new group
    max = UBound(ZKMenu())
    
    With tvTree
        If ParentIndex = -1 Then
            Rem - In root menu
            Set nodX = .Nodes.Add(, tvwLast, "|Index=" & CStr(startIndex) & "|", ZKMenu(startIndex)("Caption"))
        Else
            Rem - In a submenu
            Set nodX = .Nodes.Add("|Index=" & CStr(ParentIndex) & "|", tvwChild, "|Index=" & CStr(startIndex) & "|", ZKMenu(startIndex)("Caption"))
        End If
        'nodX.Checked = Not (Prop_Get("Disabled", ZKMenu(StartIndex)) = "True")
        nodX.ForeColor = IIf(ZKMenu(startIndex)("Disabled") = "True", ColDisabled, vbBlack)
                
        nodX.Image = "Folder"
        nodX.ExpandedImage = "FolderOpen"

    
        Rem - Now add all its sub items
        For k = startIndex + 1 To max
            Rem - Add the group item
            If ZKMenu(k)("EndGroup") = "True" Then
                Set nodX = nodX.Parent
                nodX.Expanded = False
                Exit For
            End If
            If ZKMenu(k)("Class") = "Group" Then
                Rem - Add another sub group
                k = Tree_Load_Group(startIndex, k)
            Else
                Rem - Add the item to the group
                Set nodX = .Nodes.Add("|Index=" & CStr(startIndex) & "|", tvwChild, "|Index=" & CStr(k) & "|", ZKMenu(k)("Caption"))
                nodX.Image = "Action"
            End If
            nodX.ForeColor = IIf(ZKMenu(k)("Disabled") = "True", ColDisabled, vbBlack)
            nodX.BackColor = vbWhite
            nodX.EnsureVisible
            
        Next k
        Tree_Load_Group = k
    End With

End Function

Private Sub Array_Down(ByVal Start As Long, ByVal Num As Long)
Dim k As Integer
Dim max As Integer

    max = UBound(ZKMenu())
    
    ReDim Preserve ZKMenu(max + Num)
    
    For k = max + Num To Start + Num Step -1
        Set ZKMenu(k) = ZKMenu(k - Num).Copy
    Next k
    

End Sub


Private Function Item_GetGroupEnd(ByVal GrpIndex As Long) As Long
Dim max As Long, k As Long

    max = UBound(ZKMenu())
    k = GrpIndex + 1
    
    Do
        If ZKMenu(k)("Class") = "Group" Then
            k = Item_GetGroupEnd(k)
        ElseIf ZKMenu(k)("EndGroup") = "True" Then
            Item_GetGroupEnd = k
            Exit Function
        End If
        k = k + 1
    Loop While k < max
    Item_GetGroupEnd = k
    
End Function


Public Sub Array_Swap(ByVal Item1 As Long, ByVal Item2 As Long)
'Dim End1 As Long
'Dim End2 As Long
'Dim ArrTemp() As String
'Dim k As Long
'Dim Count As Long
'Dim Max As Long
'
'    Rem - Determine the bounds of the items to be moved
'    If Prop_Get("Class", ZKMenu(Item1)) = "Group" Then End1 = Item_GetGroupEnd(Item1) Else End1 = Item1
'    If Prop_Get("Class", ZKMenu(Item2)) = "Group" Then End2 = Item_GetGroupEnd(Item2) Else End2 = Item2
'
'    Rem - Ensure smaller item is first
'    If Item1 > Item2 Then
'        Call Swap(Item1, Item2)
'        Call Swap(End1, End2)
'    End If
'
'    Rem - Copy into new array in correct order
'    Max = UBound(ZKMenu())
'    ReDim ArrTemp(0 To Max)
'    Rem - Pre 1
'    For k = 0 To Item1 - 1
'        ArrTemp(k) = ZKMenu(k)
'    Next k
'    Count = Item1
'    Rem - Copy 2 to 1
'    For k = 0 To End2 - Item2
'        ArrTemp(Count) = ZKMenu(k + Item2)
'        Count = Count + 1
'    Next k
'    Rem - Post 1, Pre 2
'    For k = 0 To Item2 - End1 - 2
'        ArrTemp(Count) = ZKMenu(End1 + k + 1)
'        Count = Count + 1
'    Next k
'    Rem - Copy 1 to 2
'    For k = 0 To End1 - Item1
'        ArrTemp(Count) = ZKMenu(Item1 + k)
'        Count = Count + 1
'    Next k
'    Rem - Post 2
'    For k = 0 To Max - End2 - 1
'        ArrTemp(Count) = ZKMenu(End2 + k + 1)
'        Count = Count + 1
'    Next k
'
'    Rem - Copy back into original array
'    For k = 0 To Max
'        ZKMenu(k) = ArrTemp(k)
'    Next k

End Sub

Public Sub Array_Move(ByVal Source As Long, ByVal dest As Long, ByVal After As Boolean)
Dim EndSource As Long
Dim k As Long
Dim Count As Long
Dim max As Long

    Rem - Determine the bounds of the items to be moved
    If ZKMenu(Source)("Class") = "Group" Then EndSource = Item_GetGroupEnd(Source) Else EndSource = Source
    If dest > Source And dest < EndSource Then
        Call ZenMB("Sorry, you cannot move a group into itself? Can you put something in itself, or is it already there?", "OK")
        Exit Sub
    End If
    
    Rem - Copy into new array in correct order
    If Source < dest Then
        If After Then
            Rem - Copy to after destination
            Call Array_Down(dest + 1, EndSource - Source + 1)
            For k = 0 To EndSource - Source
                Set ZKMenu(dest + k + 1) = ZKMenu(Source + k)
            Next
        Else
            Rem - Copy to before destination
            Call Array_Down(dest, EndSource - Source + 1)
            For k = 0 To EndSource - Source
                Set ZKMenu(dest + k) = ZKMenu(Source + k)
            Next
        End If
        Call Array_Up(EndSource + 1, EndSource - Source + 1)
    Else
        If After Then
            Rem - Copy to after destination
            Call Array_Down(dest + 1, EndSource - Source + 1)
            For k = 0 To EndSource - Source
                Set ZKMenu(dest + k + 1) = ZKMenu(Source + k + EndSource - Source + 1)
            Next
        Else
            Rem - Copy to before destination
            Call Array_Down(dest, EndSource - Source + 1)
            For k = 0 To EndSource - Source
                Set ZKMenu(dest + k) = ZKMenu(Source + k + EndSource - Source + 1)
            Next
        End If
        Call Array_Up(Source + k + EndSource - Source + 1, EndSource - Source + 1)
    
    End If
        
    
End Sub

Public Sub Move_Selected(ByVal After As Boolean)
    
    booMoving = Not booMoving
    If booMoving Then
        Rem - Begin the moving process
        booMoveAfter = After
        Rem - Check that they are not moving out the last item of the group
        If Not ZKMenu(ZIndex)("Class") = "Group" Then
            Dim lngParent As Long, lngEnd As Long
        
            lngParent = Item_GetGroup(ZIndex)
            lngEnd = Item_GetGroupEnd(lngParent)
            If lngEnd - lngParent < 2 Then
                Rem - It is the last item
                Call ZenMB("You cannot move the last item in a group. Rather just move the group itself?", "OK")
                booMoving = False
                Exit Sub
            End If
        End If
        tvTree.SelectedItem.SelectedImage = "Moving"
        lngSource = tvTree.SelectedItem.Index
    Else
        Rem - End the moving process
        tvTree.Nodes(lngSource).SelectedImage = tvTree.Nodes(lngSource).Image
        
    End If

    Rem - Set visibility of form controls
    imiItems.Visible = Not booMoving
    imiSettings.Visible = Not booMoving
    imiAbout.Visible = Not booMoving
    lblItems.Visible = Not booMoving
    lblSettings.Visible = Not booMoving
    lblParam.Visible = Not booMoving
    lblAbout.Visible = Not booMoving
    lblHelp.Visible = Not booMoving
    imiGroup.Visible = Not booMoving
    imiGroupOpen.Visible = Not booMoving
    imiHelp.Visible = Not booMoving
    
    imiItem.Visible = Not booMoving
    zbNew.Visible = Not booMoving
    zbDel.Visible = Not booMoving
    zbEdit.Visible = Not booMoving
    
    lblActionType.Visible = Not booMoving
    lblHotkey.Visible = Not booMoving
    chkRClick.Visible = Not booMoving
    zbSave.Visible = Not booMoving
    zbExit.Visible = Not booMoving
    chkEnabled.Visible = Not booMoving
    lblCaption.Visible = Not booMoving
    'shpCaption.Visible = Not booMoving
    zbMove.Visible = Not booMoving
    shpFocus.Visible = Not booMoving
    
    If Not booMoving Then
        Call Item_Selected
    Else
        lblAction.Caption = "Select where to move the item"
        'If booMoveAfter Then
            'lblAction.Caption = "Select the item after which to move it"
        'Else
        '    lblAction.Caption = "Select the item before which to move it"
        'End If
    End If

End Sub


Public Sub Tree_ReadFrom()
Rem - Recontructs the ZKMenu() Array based on the tree contents

Dim nodX As Node
Dim lngIndex As Long
Dim ArrTemp() As clsZenDictionary
Dim max As Long

    max = UBound(ZKMenu())
    ReDim ArrTemp(0 To max)
    
    Set nodX = tvTree.Nodes(1)
    Do While Not (nodX.Previous Is Nothing)
        Set nodX = nodX.Previous
    Loop
    
    Do
        Set ArrTemp(lngIndex) = ZKMenu(CLng(Prop_Get("Index", nodX.key)))
        If nodX.Children > 0 Then Call Tree_ReadNode(nodX, ArrTemp, lngIndex)
        lngIndex = lngIndex + 1
        If lngIndex > 240 Then
            lngIndex = lngIndex
        End If
        
        Set nodX = nodX.Next
    Loop While Not nodX Is Nothing
    
    ReDim ZKMenu(0 To max) ' Force a clearing
    For lngIndex = 2 To max - 2
        Set ZKMenu(lngIndex) = ArrTemp(lngIndex - 2)
    Next lngIndex

End Sub

Private Sub Tree_ReadNode(ByRef nodX As Node, ByRef ArrTemp() As clsZenDictionary, ByRef lngIndex As Long)
Dim nodChild As Node

    Set nodChild = nodX.Child
    lngIndex = lngIndex + 1
    Do
        Set ArrTemp(lngIndex) = ZKMenu(CLng(Prop_Get("Index", nodChild.key)))
        If nodChild.Children > 0 Then Call Tree_ReadNode(nodChild, ArrTemp, lngIndex)
        If lngIndex > 240 Then
            lngIndex = lngIndex
        End If
        
        lngIndex = lngIndex + 1
        Set nodChild = nodChild.Next
    Loop While Not nodChild Is Nothing
    Set ArrTemp(lngIndex) = zenDic("ENDGROUP", "True")

End Sub

Private Sub Tree_SetFocus(ByVal Index As Long)
Dim lngIndex As Long
Dim k As Long

    With tvTree
        For k = .Nodes.Count To 1 Step -1
            lngIndex = Val(Prop_Get("Index", .Nodes(k).key))
            If lngIndex = Index Then
                Set .SelectedItem = tvTree.Nodes(k)
                Exit Sub
            End If
        Next k
        
        Set .SelectedItem = .Nodes(1)
    End With

End Sub






Public Sub PrintZen(ByVal XShift As Single, ByVal YShift As Single, ByVal Col As OLE_COLOR)
'Const Top = 26, Gap = 35 '41
'Dim k As Long, strLet As String
'
'    Me.ForeColor = Col
'    For k = 1 To 6
'        strLet = Mid$("ZenKEY", k, 1)
'        'Me.CurrentX = 620 - 0.5 * Me.TextWidth(strLet) + XShift
'        Me.CurrentX = 585 'imiYY.Left + 0.5 * imiYY.Width - 0.5 * Me.TextWidth(strLet) + XShift
'
'        Me.CurrentY = Top + (k - 1) * Gap + YShift
'        Me.Print strLet
'    Next k

End Sub

Private Function Group_ItemCount(ByVal GrpIndex As Long) As Long
Dim max As Long, k As Long

    max = UBound(ZKMenu())
    k = GrpIndex + 1
    Group_ItemCount = 0
    
    Do
        If ZKMenu(k)("Class") = "Group" Then
            k = Item_GetGroupEnd(k)
        ElseIf ZKMenu(k)("EndGroup") = "True" Then
            Exit Function
        End If
        Group_ItemCount = Group_ItemCount + 1
        k = k + 1
    Loop While k < max

End Function

Private Sub Set_Init()
Dim nodX As Node
Dim k As Long
Dim strCaption As String
Dim max As Long
Dim strTemp As String
Dim lngGroups As Long, lngParent As Long
Dim strSettings() As String
    
    Rem - First, add settings that are used only by the config.
    Dim strLoad As String
    strLoad = Registry.GetRegistry(HKCU, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "ZenKEY")
    If Len(strLoad) > 0 Then settings("LOS") = "Y" Else settings("LOS") = ""
    
    Rem - Now load and build the tree.
    With tvSettings
        .Visible = False
        .Nodes.Clear
        Call LoadArray(strSettings(), App.Path & "\SetList.ini")
        
        max = UBound(strSettings())
        For k = 0 To max
            strTemp = strSettings(k)
            Select Case Prop_Get("Class", strTemp)
                Case "Group"
                    Rem - Just add to the current group
                    Set nodX = .Nodes.Add(, tvwLast, "|Index=" & CStr(k) & "|", Prop_Get("Caption", strTemp))
                    nodX.ForeColor = RGB(76, 166, 255) 'COL_Zen '&HC00000
                    nodX.Bold = True
                    
                    'nodX.Image = Choose(lngGroups Mod 7 + 1, "Appearance", "Behavior", "Window", "Quotes and messages", "Auto-window transparency", "Infinite desktop", "Desktop map")
                    nodX.Image = Choose(lngGroups Mod 7 + 1, "One", "Two", "Three", "Four", "Five", "Six", "Seven")
                    nodX.EnsureVisible
                    lngGroups = lngGroups + 1
                    lngParent = k
                Case "Setting"
                    Set nodX = .Nodes.Add("|Index=" & CStr(lngParent) & "|", tvwChild, "|Index=" & CStr(k) & "|", Prop_Get("Caption", strTemp))
                    nodX.ForeColor = vbBlack 'COL_Zen '&HC00000
                    nodX.Image = "Ying"
                Case Else
                    Set nodX = Nothing
                    lngParent = -1
            End Select
            
            If Not (nodX Is Nothing) Then
                'nodX.Bold = True
                nodX.Tag = strSettings(k)
                'nodX.ForeColor = vbBlack 'COL_Zen '&HC00000
                nodX.BackColor = vbWhite
                
            End If
        Next k
        Set tvSettings.SelectedItem = tvSettings.Nodes(1)
        .Visible = True
        
    End With
    
End Sub

Private Sub Set_Selected()
Dim strItem As String
Dim sngY As Single
Dim lngDispType As Long ' 0 - Combo box, 1 - Text box, 2 - Colour labels, 3 - Combo + Add / Remove buttons, 4 - Command button

    booLoading = True

    Rem - Clear the space
    imiSet.Visible = False
    cmbSet.Visible = False
    lblSetValue.Visible = False
    lblSetDescrip.Visible = False
    Shape1(9).Move Shape1(5).left, Shape1(5).Top, Shape1(5).Width, Shape1(5).Height
    With frmSetting
        .Move Shape1(9).left + 1, Shape1(9).Top + 1, Shape1(9).Width - 2, Shape1(9).Height - 2
    End With
    
    lblSetNotes.Visible = False
    txtSet.Visible = False
    lblTxtPostfix.Visible = False
    zbQuoteNow.Visible = False
    zbSetFileAdd.Visible = False
    zbSetFileRemove.Visible = False
    lblSetColour.Visible = False
    shpSetColour.Visible = False
    zbBackup.Visible = False

    Rem - Now get the item
    strItem = tvSettings.Nodes(SET_CurIndex).Tag
    lblZKNewSetting.Caption = Prop_Get("Caption", strItem)
    'lblSetDescrip.Width = 3165
    lblSetDescrip.Width = 3600
    'lblSetDescrip.Move frmSetting.Left + 5, frmSetting.Top + 50, frmSetting.Width - 10
    lblSetDescrip.Caption = Prop_Get("Description", strItem)
    
    If Prop_Get("Class", strItem) = "Group" Then
        Rem ----------------------------------------------------------------------------------------------
        Rem - They have selected a settings group
        Rem ----------------------------------------------------------------------------------------------
        sngY = 0.45 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * lblSetDescrip.Height
    Else
        Rem ----------------------------------------------------------------------------------------------
        Rem - They have selected a setting
        Rem ----------------------------------------------------------------------------------------------
        Rem - The string has the following properties
        Rem -     SetName - The name of the setting
        Rem -     Caption - The name to display
        Rem -     Description - A description of what the setting is and what it does
        Rem -     Default - If the new settings value is the same as this, the setting is cleared (minimizing string length)
        Rem -     Type - The type of setting. The known types are as follows, listed with their specific options
        Rem -         Dynalist / List
        Rem -             - Item1, Item2, .... - The list of item to go into the combo box
        Rem -             - Val1, Val2, .... - The list of values given when the Item is selected in thecombo box
        Rem -             - Default - The default setting if none is found.
        Rem -         Number / Integer
        Rem -             - MinValue, MaxValue - The minimum and maximum values of the number respectively
        Rem -             - Postfix - The text describing the value eg. seconds, pixels etc.
        '
        sngY = 0.2 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * lblSetDescrip.Height   '240
        'lblSetValue.Caption = "Options : "
        lblSetValue.Caption = Prop_Get("Prefix", strItem)
        If Len(lblSetValue.Caption) > 0 Then
            lblSetValue.Caption = lblSetValue.Caption & " "
        Else
            lblSetValue.Caption = vbNullString
        End If
        Rem - lngDispType = 0 - Combo box, 1 - Text box, 2 - Colour labels, 3 - Combo + Add / Remove buttons, 4- Command button
        Select Case Prop_Get("Type", strItem)
            Case "DynamicList"
                strItem = strItem & Dynalist_Load(Prop_Get("SetName", strItem))
                Call Listbox_Load(strItem)
                If Prop_Get("SetName", strItem) = "Skin" Then Call ShowSkin(cmbSet.Text)
            Case "List"
                Call Listbox_Load(strItem)
            Case "Number", "Integer"
                Call Set_NumLoad(strItem)
                lngDispType = 1
            Case "FileList", "ClassList"
                zbSetFileAdd.Tag = Prop_Get("Type", strItem)
                If Prop_Get("Loaded", strItem) <> "Y" Then
                    Call Prop_Set("Loaded", "Y", strItem)
                    strItem = strItem & Filelist_Load(strItem, zbSetFileAdd.Tag = "FileList")
                    tvSettings.Nodes(SET_CurIndex).Tag = strItem
                End If
                Call Listbox_Load(strItem)
                lngDispType = 3
            Case "Colour"
                Dim strTemp As String
                With lblSetColour
                    .Caption = Prop_Get("Caption", strItem) & Space(8)
                    .Move 0.5 * Me.ScaleX(frmSetting.Width, vbPixels, vbTwips) - 0.5 * lblSetColour.Width, 0.5 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * .Height
                    .Visible = True
                    shpSetColour.Move .left + .Width - shpSetColour.Width, .Top
                    ' Here, the setting text is actually a settings name
                    If Len(settings(Prop_Get("SetName", strItem))) > 0 Then
                        shpSetColour.FillColor = Val(settings(Prop_Get("SetName", strItem)))
                    Else
                        shpSetColour.FillColor = Val(Prop_Get("Default", strItem))
                    End If
                    shpSetColour.Visible = True
                End With
                lngDispType = 2
            Case "Joke"
                Call Listbox_Load(strItem)
                cmbSet.ListIndex = 0
            Case "Command"
                zbBackup.Caption = Prop_Get("Caption", strItem)
                zbBackup.Tag = Prop_Get("SetName", strItem) 'Backup/Restore
                lngDispType = 4
        End Select
        
        Dim sngWidth As Single
        Select Case lngDispType
            Case 0, 3 ' Combo
                With cmbSet
                    sngWidth = .Width + lblSetValue.Width
                    lblSetValue.Move 0.5 * Me.ScaleX(frmSetting.Width, vbPixels, vbTwips) - 0.5 * sngWidth, 0.5 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * .Height + 50
                    .Move lblSetValue.left + lblSetValue.Width, lblSetValue.Top - 50
                    lblSetValue.Visible = True
                    .Visible = True
                End With
                lblSetValue.Visible = True
                If lngDispType = 3 Then
                    Rem - Display add / Remove buttons
                    With zbSetFileAdd
                        '.Move 0.5 * Me.ScaleX(frmSetting.Width, vbPixels, vbTwips) - 1.1 * .Width, 0.74 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * .Height
                        .Move cmbSet.left, 0.74 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * .Height
                        .Visible = True
                    End With
                    With zbSetFileRemove
                        '.Move 0.51 * Me.ScaleX(frmSetting.Width, vbPixels, vbTwips), 0.74 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * .Height
                        .Move zbSetFileAdd.left + zbSetFileAdd.Width * 1.2, zbSetFileAdd.Top
                        .Visible = True
                    End With
                End If
            Case 1 ' Text
                lblTxtPostfix.Caption = Prop_Get("Postfix", strItem)
                If Len(lblTxtPostfix.Caption) > 0 Then
                    lblTxtPostfix.Caption = " " & lblTxtPostfix.Caption
                Else
                    lblTxtPostfix.Caption = vbNullString
                End If
                sngWidth = lblTxtPostfix.Width + lblSetValue.Width + txtSet.Width
                lblSetValue.Move 0.5 * Me.ScaleX(frmSetting.Width, vbPixels, vbTwips) - 0.5 * sngWidth, 0.5 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * txtSet.Height + 50
                txtSet.Move lblSetValue.left + lblSetValue.Width, lblSetValue.Top - 50
                lblTxtPostfix.Move txtSet.left + txtSet.Width, lblSetValue.Top
                txtSet.Visible = True
                lblSetValue.Visible = True
                lblTxtPostfix.Visible = True
            Case 2 ' Colour Label
            Case 4
                zbBackup.Move 0.5 * Me.ScaleX(frmSetting.Width, vbPixels, vbTwips) - 0.5 * zbBackup.Width, 0.5 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * zbBackup.Height
                zbBackup.Visible = True
        End Select
        
    End If
    
    lblSetDescrip.Move 250, sngY, 3600
    lblSetDescrip.Visible = True
    frmSetting.Visible = True
    
    Rem - Display buttons if required eg. 'Now', 'Add', 'Remove'
    Select Case Prop_Get("SetName", strItem)
        Case "Quotes"
            Rem - Display the extra button for showing quotes now
            With zbQuoteNow
                .Move 0.5 * Me.ScaleX(frmSetting.Width, vbPixels, vbTwips) - 0.5 * .Width, 0.74 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * .Height
                .Visible = True
            End With
        Case Else
            Rem - Display the 'Notes' field if it is there
            Dim strNotes As String
            strNotes = Prop_Get("Notes", strItem)
            If Len(strNotes) > 0 Then
                With lblSetNotes
                    .Width = 3165
                    .Caption = "Notes : " & strNotes
                    sngY = 0.74 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * .Height
                    .Move 240, sngY, 3165
                    .Visible = True
                End With
            ElseIf Prop_Get("SetName", strItem) <> "Skin" Then
                If lngDispType <> 3 Then
                    Rem - If this is no the skin setting, display a YinYang or summin'
                    Call ShowSkin(settings("Skin"))
                End If
            End If
    End Select
    
    txtSet.Enabled = True
    cmbSet.Enabled = txtSet.Enabled
    booLoading = False

End Sub


Public Function GetColour(ByRef NewCol As OLE_COLOR) As Boolean
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long

    Rem - Initialise the structure
    cc.lStructSize = Len(cc) 'set the structure size
    cc.hWndOwner = Me.hwnd
    cc.hInstance = App.hInstance
    cc.lpCustColors = StrConv(CustomColors, vbUnicode) 'set the custom colors (converted to Unicode)
    cc.flags = 0 'no extra flags

    Rem - Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
        NewCol = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
        GetColour = True
    'Else
    '    ShowColor = -1
    End If

    


End Function

Private Sub Listbox_Load(ByVal strList As String)
Dim k As Long, strItem As String
Dim strSetName As String, strCurSet As String
Dim strVal As String
Dim strMaxCaption As String ' For setting the width of the listbox

    strSetName = Prop_Get("SetName", strList) ' The setting name
    strCurSet = settings(strSetName) ' The current setting
    If Len(strCurSet) = 0 Then strCurSet = Prop_Get("Default", strList) ' Use the default if nothing
    With cmbSet
        .Clear
        k = 1
        strItem = Prop_Get("Item" & CStr(k), strList)
        While Len(strItem) > 0
            .AddItem strItem
            If Len(strItem) > Len(strMaxCaption) Then strMaxCaption = strItem
            
            Rem - Now check whether this is the current setting
            strVal = Prop_Get("Val" & CStr(k), strList)
            If Len(strVal) = 0 Then strVal = strItem
            If strCurSet = strVal Then .ListIndex = k - 1
            
            Rem - Check for the next item
            k = k + 1
            strItem = Prop_Get("Item" & CStr(k), strList)
        Wend
        
        Rem - Ensure that if no setting is picked up, the default is selected
        If .ListIndex < 0 Then
            Select Case Prop_Get("Type", strList)
                Case "FileList", "ClassList"
                    Rem - A list of files - Show the last item
                    .ListIndex = .ListCount - 1
                Case Else
                    Rem - Show the default/current setting
                    strCurSet = Prop_Get("Default", strList)
                    For k = 0 To .ListCount - 1
                        If Prop_Get("Val" & CStr(k + 1), strList) = strCurSet Then
                            .ListIndex = k
                            Exit For
                        End If
                    Next k
            End Select
        End If
        
        Set picNewSettings.Font = .Font
        .Width = picNewSettings.ScaleX(picNewSettings.TextWidth(strMaxCaption & Space(10)), picNewSettings.ScaleMode, vbTwips)
    End With
    
    
End Sub

Private Sub ShowSkin(ByVal Skin As String)
    
    If Len(Skin) = 0 Then Skin = "Default"
    'Set imiSet.Picture = LoadPicture(App.Path & "\Skins\" & cmbSet.Text & ".ico")
    Set imiSet.Picture = LoadPicture(App.Path & "\Skins\" & Skin & ".ico")
    imiSet.Move 0.5 * Me.ScaleX(frmSetting.Width, vbPixels, vbTwips) - 0.5 * imiSet.Width, 0.74 * Me.ScaleY(frmSetting.Height, vbPixels, vbTwips) - 0.5 * imiSet.Height
    'imiSkin.Left = 0.5 * Me.ScaleX(frmSetting.Width, vbPixels, vbTwips) - 0.5 * Me.ScaleX(imiSet.Width, vbTwips, vbPixels)
    
    
    imiSet.Visible = True

End Sub


Private Sub Set_NumLoad(ByVal strProp As String)
Dim strSetName As String
Dim sngVal As Single
Dim strVal As String
    
    strSetName = Prop_Get("SetName", strProp) ' The setting name
    strVal = settings(strSetName) ' The current setting
    If Len(strVal) = 0 Then strVal = Prop_Get("Default", strProp)
    sngVal = Val(strVal) ' The current setting
    If (sngVal < 1) And (sngVal <> 0) Then
        txtSet.Text = Format(sngVal, "0.#")
    Else
        txtSet.Text = Trim(str(sngVal))
    End If
    
    txtSet.Move lblSetValue.left + lblSetValue.Width + 40, lblSetValue.Top - 60
    txtSet.Tag = strProp

End Sub

Private Sub Set_Save(ByVal SaveTree As Boolean)
Rem - Save the settings
Dim k As Long

    Rem ------------------------------------------------------------------------------------------------------------------------
    Rem - First, add settings that are used only by the config.
    Rem ------------------------------------------------------------------------------------------------------------------------
    If settings("ResetCount") = "Y" Then
        If Len(Dir(settings("SavePath") & "\ProgInfo.ini")) > 0 Then Call Kill(settings("SavePath") & "\ProgInfo.ini")
    End If
    settings("ResetCount") = ""
    
    Rem ------------------------------------------------------------------------------------------------------------------------
    Rem - Now save the modified  ties
    Rem -----------------------------------------------------------------------------------------------------------------------
    Dim FNum As Long
    Dim strTemp As String, strItem As String
    
    If settings("LOS") = "Y" Then
        Call Registry.SetRegistry(HKCU, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "ZenKEY", App.Path & "\ZenKEY.exe")
    Else
        Call Registry.DelRegistry(HKCU, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "ZenKEY")
    End If
    
    If SaveTree Then
        Rem - If saving from the registry, do not process the settings tree for file lists
        With tvSettings
            For k = .Nodes.Count To 1 Step -1
                strTemp = .Nodes(k).Tag
                Select Case Prop_Get("Type", strTemp)
                    Case "FileList", "ClassList"
                        If Prop_Get("Loaded", strTemp) = "Y" Then
                            If (Len(Prop_Get("Item2", strTemp)) > 1) Or (Prop_Get("Item1", strTemp) <> List_Null) Then
                                Rem - Save the file list
                                Dim i As Long
                                
                                FNum = FreeFile
                                Open settings("SavePath") & "\" & Prop_Get("FileName", strTemp) For Output As #FNum
                                    i = 1
                                    strItem = Prop_Get("Item" & CStr(i), strTemp)
                                    While Len(strItem) > 0
                                        Write #FNum, strItem
                                        i = i + 1
                                        strItem = Prop_Get("Item" & CStr(i), strTemp)
                                    Wend
                                Close #FNum
                            Else
                                Rem - CLear the file if we have an empty list
                                If Len(Dir(settings("SavePath") & "\" & Prop_Get("FileName", strTemp))) > 0 Then Call Kill(settings("SavePath") & "\" & Prop_Get("FileName", strTemp))
                            End If
                        End If ' Prop_Get("Loaded"
                End Select ' Prop_Get("Type"
            Next k
        End With
    End If
    
    Rem ----------------------------------------------------------------------------------------------------------------------------------------
    strTemp = settings("LOS")
    settings("LOS") = ""
    Call settings.ToINI(settings("SavePath") & "\Settings.ini")
    settings("LOS") = strTemp

End Sub

Private Function Filelist_Load(ByVal strItem As String, ByVal IsExe As Boolean) As String
Dim k As Long, lngCount As Long
Dim colFiles As Collection
    
    Set colFiles = New Collection
    Call INI_LoadFiles(Prop_Get("FileName", strItem), colFiles, IsExe)
    lngCount = colFiles.Count
    
    If lngCount > 0 Then
        For k = 1 To lngCount
            Call Prop_Set("Item" & CStr(k), colFiles.Item(k), Filelist_Load)
        Next k
    Else
        Call Prop_Set("Item1", List_Null, Filelist_Load)
    End If
    
End Function

Private Sub Filelist_Update(ByVal FileName As String, ByVal booAdd As Boolean)
Dim strItem As String, lngCount  As Long
Dim booFound As Boolean, strTemp As String
Dim lngMatch As Long

    strItem = tvSettings.SelectedItem.Tag
    strTemp = Prop_Get("Item" & CStr(lngCount + 1), strItem)
    If strTemp <> List_Null Then
        While Len(strTemp) > 0
            lngCount = lngCount + 1
            If Not booAdd Then If strTemp = FileName Then lngMatch = lngCount
            strTemp = Prop_Get("Item" & CStr(lngCount + 1), strItem)
        Wend
    End If ' List_Null
    
    
    If booAdd Then
        Rem - Add the item to the end
        Call Prop_Set("Item" & CStr(lngCount + 1), FileName, strItem)
    Else
        Rem - Remove the item
        Select Case True
            Case lngCount = 1
                Rem - There is only one item
                Call Prop_Set("Item1", List_Null, strItem)
            Case lngCount = lngMatch
                Rem - Remove the last item
                Call Prop_Set("Item" & CStr(lngCount), vbNullString, strItem)
            Case Else
                Rem - Swap with the last item
                strTemp = Prop_Get("Item" & CStr(lngCount), strItem)
                Call Prop_Set("Item" & CStr(lngMatch), strTemp, strItem)
                Call Prop_Set("Item" & CStr(lngCount), vbNullString, strItem)
        End Select
    End If
    tvSettings.SelectedItem.Tag = strItem

End Sub

Private Sub Set_SetValue(ByVal strSetting As String)
Dim strNewVal As String, strSetName As String
Dim strDefault As String
        
        booSetChanged = True
        strSetName = Prop_Get("SetName", strSetting)
        Select Case Prop_Get("Type", strSetting)
            Case "FileList", "ClassList"
                Rem - The 'FileList_Update' sub should already have set the file list string
                Exit Sub
            Case "List", "DynamicList"
                strNewVal = Prop_Get("Val" & CStr(cmbSet.ListIndex + 1), strSetting)
                If Len(strNewVal) = 0 Then strNewVal = cmbSet.Text
                If strSetName = "Skin" Then Call ShowSkin(cmbSet.Text) 'Display a skin preview
            Case "Number", "Integer"
                Dim sngMin As Single, sngMax As Single
                
                Rem - Ensure the value of the text is between the specified range
                sngMin = Val(Prop_Get("MinValue", strSetting))
                Rem - Force minimum
                If Val(txtSet.Text) < sngMin Then
                    If (sngMin < 1) And (Val(txtSet.Text) <> 0) Then strNewVal = Format(sngMin, "0.#") Else strNewVal = Trim(str(sngMin))
                End If
                sngMax = Val(Prop_Get("MaxValue", strSetting))
                If sngMax > 0 Then
                    Rem - Force maximum
                    If Val(txtSet.Text) > sngMax Then strNewVal = Trim(str(sngMax))
                End If
                If Len(strNewVal) = 0 Then
                    strNewVal = txtSet.Text
                Else
                    booLoading = True
                    txtSet.Text = strNewVal
                    booLoading = False
                    Rem - Note : Modal showing seems to interfere with the firing of the node_click event
                    DoEvents
                    Call ZenMB("Sorry, but your chosen value for '" & Prop_Get("Caption", strSetting) & "' is outside of the required range.  It has been reset to " & txtSet.Text & " " & lblTxtPostfix.Caption & ".", "OK")
                End If
            Case Else
        End Select
        strDefault = Prop_Get("Default", strSetting)
        If strNewVal = strDefault Then
            settings(strSetName) = vbNullString
        Else
            settings(strSetName) = strNewVal
        End If
        
        If strSetName = "AutoTrans" Then
            Rem - Give warning about Transparency
            If strNewVal = "True" Then Call ZenMB(ZK_TransWarn, "OK")
        End If

End Sub

Public Sub ShowHotkeys()
Dim f As Long, strFName As String
Dim k As Long

    f = FreeFile
    strFName = getTemp & "ZenKeys.txt"
    Open strFName For Output As #f
        For k = 0 To UBound(ZKMenu)
            If Len(ZKMenu(k)("Hotkey")) > 0 Then
                Print #f, ZKMenu(k)("Caption") & " - " & HotKeys.GetCaption(ZKMenu(k))
            End If
        Next k
    Close #f
    Call ShellExe(strFName)

End Sub

Private Function getTemp() As String
Dim strUserName As String

    'Create a buffer
    strUserName = String(255, Chr$(0))
    'Get the username
    GetTempPath 255, strUserName
    'strip the rest of the buffer
    getTemp = left$(strUserName, InStr(strUserName, Chr$(0)) - 1)

End Function
