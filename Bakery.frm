VERSION 5.00
Begin VB.Form FrmBakery 
   BackColor       =   &H00004080&
   Caption         =   "Management_System"
   ClientHeight    =   11970
   ClientLeft      =   3810
   ClientTop       =   2160
   ClientWidth     =   21645
   LinkTopic       =   "Form1"
   ScaleHeight     =   11970
   ScaleWidth      =   21645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   50
      Top             =   120
      Width           =   21255
      Begin VB.Label Label8 
         Caption         =   "Bakery Shop Management System"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3720
         TabIndex        =   51
         Top             =   240
         Width           =   13815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9735
      Left            =   12000
      TabIndex        =   1
      Top             =   1920
      Width           =   9375
      Begin VB.TextBox txtTax 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   63
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtPaid 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Rp""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   48
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox txtService 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Rp""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   46
         Top             =   3480
         Width           =   3495
      End
      Begin VB.ComboBox cboMethod 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3480
         TabIndex        =   39
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6240
         TabIndex        =   8
         Top             =   8520
         Width           =   2535
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset "
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3240
         TabIndex        =   7
         Top             =   8520
         Width           =   2535
      End
      Begin VB.CommandButton cmdTotal 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   6
         Top             =   8520
         Width           =   2535
      End
      Begin VB.Frame Frame8 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   360
         TabIndex        =   5
         Top             =   4440
         Width           =   8415
         Begin VB.TextBox txtChange 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Rp""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14345
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            TabIndex        =   45
            Top             =   2760
            Width           =   4695
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Rp""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14345
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   3240
            TabIndex        =   44
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label Label7 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1680
            TabIndex        =   43
            Top             =   2760
            Width           =   6255
         End
         Begin VB.Label Label6 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   1680
            TabIndex        =   42
            Top             =   600
            Width           =   6135
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Tax"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   62
         Top             =   1680
         Width           =   6015
      End
      Begin VB.Label Label4 
         Caption         =   "Paid"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   47
         Top             =   2520
         Width           =   6015
      End
      Begin VB.Label Label5 
         Caption         =   "Service Charge"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   41
         Top             =   3480
         Width           =   6015
      End
      Begin VB.Label Label2 
         Caption         =   "Method"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   40
         Top             =   960
         Width           =   6015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9735
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   11655
      Begin VB.Frame Frame5 
         Caption         =   "Cake"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   5880
         TabIndex        =   52
         Top             =   480
         Width           =   5535
         Begin VB.TextBox txtCheese 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   49
            Top             =   3600
            Width           =   975
         End
         Begin VB.TextBox txtBlack 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   61
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtTiramisu 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   60
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtRed 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   59
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtSponge 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   58
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkCheese 
            Caption         =   "Cheese Cake"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   57
            Top             =   3480
            Width           =   5055
         End
         Begin VB.CheckBox chkBlack 
            Caption         =   "Black Forest"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   56
            Top             =   2760
            Width           =   5055
         End
         Begin VB.CheckBox chkTiramisu 
            Caption         =   "Tiramisu"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   55
            Top             =   2040
            Width           =   5055
         End
         Begin VB.CheckBox chkRed 
            Caption         =   "Red Velvet"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   54
            Top             =   1320
            Width           =   5055
         End
         Begin VB.CheckBox chkSponge 
            Caption         =   "Sponge Cake"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   240
            TabIndex        =   53
            Top             =   600
            Width           =   5055
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Drinks"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   5880
         TabIndex        =   4
         Top             =   5040
         Width           =   5535
         Begin VB.TextBox txtSmoothie 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   38
            Top             =   3600
            Width           =   975
         End
         Begin VB.TextBox txtIced 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   37
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtMatcha 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   36
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtEspresso 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   35
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtLatte 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   34
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkSmoothie 
            Caption         =   "Smoothie"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   23
            Top             =   3600
            Width           =   4935
         End
         Begin VB.CheckBox chkIced 
            Caption         =   "Iced Americano"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   22
            Top             =   2880
            Width           =   4935
         End
         Begin VB.CheckBox chkMatcha 
            Caption         =   "Matcha Latte"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   21
            Top             =   2160
            Width           =   4935
         End
         Begin VB.CheckBox chkEspresso 
            Caption         =   "Espresso"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   20
            Top             =   1440
            Width           =   4935
         End
         Begin VB.CheckBox chkLatte 
            Caption         =   "Latte"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   4935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Roti"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   120
         TabIndex        =   3
         Top             =   5040
         Width           =   5535
         Begin VB.TextBox txtBread 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   33
            Top             =   3600
            Width           =   975
         End
         Begin VB.TextBox txtBaguette 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   32
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtSourdough 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   31
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtGarlic 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   30
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtGandum 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   29
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkBread 
            Caption         =   "Bread Stick"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   18
            Top             =   3600
            Width           =   4935
         End
         Begin VB.CheckBox chkBaguette 
            Caption         =   "Baguette"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   17
            Top             =   2880
            Width           =   4935
         End
         Begin VB.CheckBox chkSourdough 
            Caption         =   "Sourdough"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   16
            Top             =   2160
            Width           =   4935
         End
         Begin VB.CheckBox chkGarlic 
            Caption         =   "Garlic Bread"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   15
            Top             =   1440
            Width           =   4935
         End
         Begin VB.CheckBox chkGandum 
            Caption         =   "Gandum"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   4935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Pastry"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   5535
         Begin VB.TextBox txtCromboloni 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   28
            Top             =   3600
            Width           =   975
         End
         Begin VB.TextBox txtPie 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   27
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtEclair 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   26
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtPuff 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   25
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtCroissant 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14345
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   24
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkCromboloni 
            Caption         =   "Cromboloni"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   13
            Top             =   3600
            Width           =   4935
         End
         Begin VB.CheckBox chkPie 
            Caption         =   "Pie"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   12
            Top             =   2880
            Width           =   4935
         End
         Begin VB.CheckBox chkEclair 
            Caption         =   "Eclair"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   11
            Top             =   2160
            Width           =   4935
         End
         Begin VB.CheckBox chkPuff 
            Caption         =   "Puff Pastry"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   10
            Top             =   1440
            Width           =   4935
         End
         Begin VB.CheckBox chkCroissant 
            Caption         =   "Croissant"
            BeginProperty Font 
               Name            =   "Roboto"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   4935
         End
      End
   End
End
Attribute VB_Name = "FrmBakery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Const hargaCroissant As Currency = 50000
    Private Const hargaBaguette As Currency = 28000
    Private Const hargaGandum As Currency = 15500
    Private Const hargaPuff As Currency = 25500
    Private Const hargaEclair As Currency = 27500
    Private Const hargaPie As Currency = 30000
    Private Const hargaCromboloni As Currency = 45500
    Private Const hargaSponge As Currency = 30000
    Private Const hargaCheese As Currency = 35500
    Private Const hargaTiramisu As Currency = 32000
    Private Const hargaRedVelvet As Currency = 38000
    Private Const hargaBlackForest As Currency = 37500
    Private Const hargaLatte As Currency = 25000
    Private Const hargaEspresso As Currency = 20000
    Private Const hargaMatcha As Currency = 27500
    Private Const hargaIcedAmericano As Currency = 22000
    Private Const hargaSmoothie As Currency = 23500
    Private Const hargaBreadStick As Currency = 15500
    Private Const hargaSourdough As Currency = 20000
    Private Const hargaGarlicBread As Currency = 26000
    
Private Sub cboMethod_Click()
If cboMethod.Text = "Cash" Then
        txtPaid.Enabled = True
        txtPaid.Text = ""
        txtPaid.SetFocus  ' Fokus ke input paid
    Else
        txtPaid.Enabled = False
        txtPaid.Text = ""
        txtChange.Text = ""  ' Reset change
    End If
End Sub

Private Sub chkBaguette_Click()
    If chkBaguette.Value = vbChecked Then
        txtBaguette.Enabled = True
        txtBaguette.SetFocus
    Else
        txtBaguette.Enabled = False
        txtBaguette.Text = "0"
    End If
End Sub

Private Sub chkBlack_Click()
If chkBlack.Value = vbChecked Then
        txtBlack.Enabled = True
        txtBlack.SetFocus
    Else
        txtBlack.Enabled = False
        txtBlack.Text = "0"
    End If
End Sub

Private Sub chkBread_Click()
If chkBread.Value = vbChecked Then
        txtBread.Enabled = True
        txtBread.SetFocus
    Else
        txtBread.Enabled = False
        txtBread.Text = "0"
    End If
End Sub

Private Sub chkCheese_Click()
If chkCheese.Value = vbChecked Then
        txtCheese.Enabled = True
        txtCheese.SetFocus
    Else
        txtCheese.Enabled = False
        txtCheese.Text = "0"
    End If
End Sub
    

Private Sub chkCroissant_Click()
    If chkCroissant.Value = vbChecked Then
        txtCroissant.Enabled = True
        txtCroissant.SetFocus
    Else
        txtCroissant.Enabled = False
        txtCroissant.Text = "0"
    End If
    
End Sub

Private Sub chkCromboloni_Click()
If chkCromboloni.Value = vbChecked Then
        txtCromboloni.Enabled = True
        txtCromboloni.SetFocus
    Else
        txtCromboloni.Enabled = False
        txtCromboloni.Text = "0"
    End If
End Sub

Private Sub chkEclair_Click()
 If chkEclair.Value = vbChecked Then
        txtEclair.Enabled = True
        txtEclair.SetFocus
    Else
        txtEclair.Enabled = False
        txtEclair.Text = "0"
    End If
End Sub

Private Sub chkEspresso_Click()
If chkEspresso.Value = vbChecked Then
        txtEspresso.Enabled = True
        txtEspresso.SetFocus
    Else
        txtEspresso.Enabled = False
        txtEspresso.Text = "0"
    End If
End Sub

Private Sub chkGandum_Click()
If chkGandum.Value = vbChecked Then
        txtGandum.Enabled = True
        txtGandum.SetFocus
    Else
        txtGandum.Enabled = False
        txtGandum.Text = "0"
    End If
End Sub

Private Sub chkGarlic_Click()
If chkGarlic.Value = vbChecked Then
        txtGarlic.Enabled = True
        txtGarlic.SetFocus
    Else
        txtGarlic.Enabled = False
        txtGarlic.Text = "0"
    End If
End Sub

Private Sub chkIced_Click()
If chkIced.Value = vbChecked Then
        txtIced.Enabled = True
        txtIced.SetFocus
    Else
        txtIced.Enabled = False
        txtIced.Text = "0"
    End If
End Sub

Private Sub chkLatte_Click()
If chkLatte.Value = vbChecked Then
        txtLatte.Enabled = True
        txtLatte.SetFocus
    Else
        txtLatte.Enabled = False
        txtLatte.Text = "0"
    End If
End Sub

Private Sub chkMatcha_Click()
If chkMatcha.Value = vbChecked Then
        txtMatcha.Enabled = True
        txtMatcha.SetFocus
    Else
        txtMatcha.Enabled = False
        txtMatcha.Text = "0"
    End If
End Sub

Private Sub chkPie_Click()
 If chkPie.Value = vbChecked Then
        txtPie.Enabled = True
        txtPie.SetFocus
    Else
        txtPie.Enabled = False
        txtPie.Text = "0"
    End If
End Sub

Private Sub chkPuff_Click()
 If chkPuff.Value = vbChecked Then
        txtPuff.Enabled = True
        txtPuff.SetFocus
    Else
        txtPuff.Enabled = False
        txtPuff.Text = "0"
    End If
End Sub

Private Sub chkRed_Click()
If chkRed.Value = vbChecked Then
        txtRed.Enabled = True
        txtRed.SetFocus
    Else
        txtRed.Enabled = False
        txtRed.Text = "0"
    End If
End Sub

Private Sub chkSmoothie_Click()
If chkSmoothie.Value = vbChecked Then
        txtSmoothie.Enabled = True
        txtSmoothie.SetFocus
    Else
        txtSmoothie.Enabled = False
        txtSmoothie.Text = "0"
    End If
End Sub

Private Sub chkSourdough_Click()
If chkSourdough.Value = vbChecked Then
        txtSourdough.Enabled = True
        txtSourdough.SetFocus
    Else
        txtSourdough.Enabled = False
        txtSourdough.Text = "0"
    End If
End Sub

Private Sub chkSponge_Click()
If chkSponge.Value = vbChecked Then
        txtSponge.Enabled = True
        txtSponge.SetFocus
    Else
        txtSponge.Enabled = False
        txtSponge.Text = "0"
    End If
End Sub

Private Sub chkTiramisu_Click()
If chkTiramisu.Value = vbChecked Then
        txtTiramisu.Enabled = True
        txtTiramisu.SetFocus
    Else
        txtTiramisu.Enabled = False
        txtTiramisu.Text = "0"
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdReset_Click()
' Harus menggunakan Collection atau Array
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = "0"  ' Reset semua TextBox di form
        End If
    Next ctrl
End Sub

Private Sub cmdTotal_Click()
    Dim itemTotal As Currency
    Dim tax As Currency
    Dim service As Currency
    Dim grandTotal As Currency
    Dim paid As Currency
    Dim change As Currency
    
    ' totals
    itemTotal = 0
    service = 5000
    
    ' Hitung item total
    ' PASTRY
    If chkCroissant.Value = vbChecked And IsNumeric(txtCroissant.Text) Then
        itemTotal = itemTotal + CCur(txtCroissant.Text) * hargaCroissant
    End If
    If chkPuff.Value = vbChecked And IsNumeric(txtPuff.Text) Then
        itemTotal = itemTotal + CCur(txtPuff.Text) * hargaPuff
    End If
    If chkEclair.Value = vbChecked And IsNumeric(txtEclair.Text) Then
        itemTotal = itemTotal + CCur(txtEclair.Text) * hargaEclair
    End If
    If chkPie.Value = vbChecked And IsNumeric(txtPie.Text) Then
        itemTotal = itemTotal + CCur(txtPie.Text) * hargaPie
    End If
    If chkCromboloni.Value = vbChecked And IsNumeric(txtCromboloni.Text) Then
        itemTotal = itemTotal + CCur(txtCromboloni.Text) * hargaCromboloni
    End If
    
    'ROTI
    If chkGandum.Value = vbChecked And IsNumeric(txtGandum.Text) Then
        itemTotal = itemTotal + CCur(txtGandum.Text) * hargaGandum
    End If
    If chkGarlic.Value = vbChecked And IsNumeric(txtGarlic.Text) Then
        itemTotal = itemTotal + CCur(txtGarlic.Text) * hargaGarlic
    End If
    If chkSourdough.Value = vbChecked And IsNumeric(txtSourdough.Text) Then
        itemTotal = itemTotal + CCur(txtSourdough.Text) * hargaSourdough
    End If
    If chkBaguette.Value = vbChecked And IsNumeric(txtBaguette.Text) Then
        itemTotal = itemTotal + CCur(txtBaguette.Text) * hargaBaguette
    End If
    If chkBread.Value = vbChecked And IsNumeric(txtBread.Text) Then
        itemTotal = itemTotal + CCur(txtBread.Text) * hargaBread
    End If
    
    'CAKE
     If chkSponge.Value = vbChecked And IsNumeric(txtSponge.Text) Then
        itemTotal = itemTotal + CCur(txtSponge.Text) * hargaSponge
    End If
    If chkRed.Value = vbChecked And IsNumeric(txtRed.Text) Then
        itemTotal = itemTotal + CCur(txtRed.Text) * hargaRed
    End If
    If chkTiramisu.Value = vbChecked And IsNumeric(txtTiramisu.Text) Then
        itemTotal = itemTotal + CCur(txtTiramisu.Text) * hargaTiramisu
    End If
    If chkBlack.Value = vbChecked And IsNumeric(txtBlack.Text) Then
        itemTotal = itemTotal + CCur(txt.TextBlack) * hargaBlack
    End If
    If chkCheese.Value = vbChecked And IsNumeric(txtCheese.Text) Then
        itemTotal = itemTotal + CCur(txtCheese.Text) * hargaCheese
    End If
    
    'DRINKS
     If chkLatte.Value = vbChecked And IsNumeric(txtLatte.Text) Then
        itemTotal = itemTotal + CCur(txtLatte.Text) * hargaLatte
    End If
    If chkEspresso.Value = vbChecked And IsNumeric(txtEspresso.Text) Then
        itemTotal = itemTotal + CCur(txtEspresso.Text) * hargaEspresso
    End If
    If chkMatcha.Value = vbChecked And IsNumeric(txtMatcha.Text) Then
        itemTotal = itemTotal + CCur(txtMatcha.Text) * hargaMatcha
    End If
    If chkIced.Value = vbChecked And IsNumeric(txtIced.Text) Then
        itemTotal = itemTotal + CCur(txtIced.Text) * hargaIced
    End If
    If chkSmoothie.Value = vbChecked And IsNumeric(txtSmoothie.Text) Then
        itemTotal = itemTotal + CCur(txtSmoothie.Text) * hargaSmoothie
    End If
    
    ' Hitung tax dan grand total
    tax = itemTotal * 0.1
    grandTotal = itemTotal + tax + service
    
    ' Tampilkan perhitungan nya
    txtTax.Text = Format(tax, "Rp #,##0")
    txtService.Text = Format(service, "Rp #,##0")
    txtTotal.Text = Format(grandTotal, "Rp #,##0")
    
    ' Poses pembayaran
    If cboMethod.Text = "Cash" Then
        If IsNumeric(txtPaid.Text) Then
            paid = CCur(txtPaid.Text)
            If paid >= grandTotal Then
                change = paid - grandTotal
                txtChange.Text = Format(change, "Rp #,##0")
            Else
                txtChange.Text = "Pembayaran Kurang"
            End If
        Else
            txtChange.Text = "Input nominal!"
        End If
    Else
        txtChange.Text = ""
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
 
    txtCroissant.Text = "0"
    txtPuff.Text = "0"
    txtEclair.Text = "0"
    txtPie.Text = "0"
    txtCromboloni.Text = "0"
    txtGandum.Text = "0"
    txtGarlic.Text = "0"
    txtSourdough.Text = "0"
    txtBaguette.Text = "0"
    txtBread.Text = "0"
    txtSponge.Text = "0"
    txtRed.Text = "0"
    txtTiramisu.Text = "0"
    txtBlack.Text = "0"
    txtCheese.Text = "0"
    txtLatte.Text = "0"
    txtEspresso.Text = "0"
    txtMatcha.Text = "0"
    txtIced.Text = "0"
    txtSmoothie.Text = "0"
    
    txtCroissant.Enabled = False
    txtPuff.Enabled = False
    txtEclair.Enabled = False
    txtPie.Enabled = False
    txtCromboloni.Enabled = False
    txtGandum.Enabled = False
    txtGarlic.Enabled = False
    txtSourdough.Enabled = False
    txtBaguette.Enabled = False
    txtBread.Enabled = False
    txtSponge.Enabled = False
    txtRed.Enabled = False
    txtTiramisu.Enabled = False
    txtBlack.Enabled = False
    txtCheese.Enabled = False
    txtLatte.Enabled = False
    txtEspresso.Enabled = False
    txtMatcha.Enabled = False
    txtIced.Enabled = False
    txtSmoothie.Enabled = False
   
    chkCroissant.Enabled = True
    chkPuff.Enabled = True
    chkEclair.Enabled = True
    chkPie.Enabled = True
    chkCromboloni.Enabled = True
    chkGandum.Enabled = True
    chkGarlic.Enabled = True
    chkSourdough.Enabled = True
    chkBaguette.Enabled = True
    chkBread.Enabled = True
    chkSponge.Enabled = True
    chkRed.Enabled = True
    chkTiramisu.Enabled = True
    chkBlack.Enabled = True
    chkCheese.Enabled = True
    chkLatte.Enabled = True
    chkEspresso.Enabled = True
    chkMatcha.Enabled = True
    chkIced.Enabled = True
    chkSmoothie.Enabled = True
    
    
    cboMethod.AddItem "Cash"
    cboMethod.AddItem "Credit Card"
    cboMethod.AddItem "Debit Card"
    cboMethod.AddItem "E-Wallet"
    
    txtPaid.Enabled = False
    txtPaid.Text = ""
    txtChange.Text = ""
    
End Sub


Private Sub label_Click()

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub txtBaguette_Change()
    If txtBaguette.Text = "0" Then
        txtBaguette.Text = ""
    End If
End Sub

Private Sub txtBaguette_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtBlack_Change()
If txtBlack.Text = "0" Then
        txtBlack.Text = ""
    End If
End Sub

Private Sub txtBread_Change()
If txtBread.Text = "0" Then
        txtBread.Text = ""
    End If
End Sub

Private Sub txtCheese_Change()
If txtCheese.Text = "0" Then
        txtCheese.Text = ""
    End If
End Sub

Private Sub txtCroissant_Change()
If txtCroissant.Text = "0" Then
        txtCroissant.Text = ""
    End If
End Sub

Private Sub txtCromboloni_Change()
If txtCromboloni.Text = "0" Then
        txtCromboloni.Text = ""
    End If
End Sub

Private Sub txtEclair_Change()
If txtEclair.Text = "0" Then
        txtEclair.Text = ""
    End If
End Sub

Private Sub txtEspresso_Change()
If txtEspresso.Text = "0" Then
        txtEspresso.Text = ""
    End If
End Sub

Private Sub txtGandum_Change()
If txtGandum.Text = "0" Then
        txtGandum.Text = ""
    End If
End Sub

Private Sub txtGarlic_Change()
If txtGarlic.Text = "0" Then
        txtGarlic.Text = ""
    End If
End Sub

Private Sub txtIced_Change()
If txtIced.Text = "0" Then
        txtIced.Text = ""
    End If
End Sub

Private Sub txtLatte_Change()
If txtLatte.Text = "0" Then
        txtLatte.Text = ""
    End If
End Sub

Private Sub txtMatcha_Change()
If txtMatcha.Text = "0" Then
        txtMatcha.Text = ""
    End If
End Sub

Private Sub txtPaid_Change()
    ' Hanya izinkan angka dan backspace
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPie_Change()
If txtPie.Text = "0" Then
        txtPie.Text = ""
    End If
End Sub

Private Sub txtPuff_Change()
If txtPuff.Text = "0" Then
        txtPuff.Text = ""
    End If
End Sub

Private Sub txtRed_Change()
If txtRed.Text = "0" Then
        txtRed.Text = ""
    End If
End Sub

Private Sub txtSmoothie_Change()
If txtSmoothie.Text = "0" Then
        txtSmoothie.Text = ""
    End If
End Sub

Private Sub txtSourdough_Change()
If txtSourdough.Text = "0" Then
        txtSourdough.Text = ""
    End If
End Sub

Private Sub txtSponge_Change()
If txtSponge.Text = "0" Then
        txtSponge.Text = ""
    End If
End Sub

Private Sub txtTiramisu_Change()
If txtTiramisu.Text = "0" Then
        txtTiramisu.Text = ""
    End If
End Sub
Private Sub txtCroissant_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtPuff_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtEclair_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtPie_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtCromboloni_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtGandum_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtGarlic_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtSourdough_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtBread_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtSponge_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtRed_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtTiramisu_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtBlack_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtCheese_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtLatte_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtEspresso_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtMatcha_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtIced_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtSmoothie_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
End Sub


