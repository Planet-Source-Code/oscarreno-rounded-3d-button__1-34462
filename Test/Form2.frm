VERSION 5.00
Object = "*\A..\MultiCon.vbp"
Begin VB.Form Form2 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Button"
   ClientHeight    =   6330
   ClientLeft      =   3495
   ClientTop       =   2085
   ClientWidth     =   7635
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7635
   Begin MultiCon.Boton Ir 
      Height          =   675
      Left            =   5400
      TabIndex        =   35
      Top             =   5520
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1191
      Caption         =   "Go to Page"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form2.frx":0000
      BackColor       =   8421631
      ForeColor       =   9768745
      RoundSize       =   0
      Light           =   60
   End
   Begin MultiCon.Boton BotonIzq 
      Height          =   195
      Index           =   5
      Left            =   420
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   "<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin MultiCon.Boton BotonIzq 
      Height          =   195
      Index           =   4
      Left            =   420
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4900
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   "<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin MultiCon.Boton BotonIzq 
      Height          =   195
      Index           =   3
      Left            =   420
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   "<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin MultiCon.Boton BotonIzq 
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3195
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   "<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin MultiCon.Boton BotonDer 
      Height          =   195
      Index           =   1
      Left            =   3180
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1640
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   ">"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin MultiCon.Boton BotonIzq 
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1635
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   "<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin VB.HScrollBar HScroll 
      Height          =   100
      Index           =   1
      Left            =   420
      Max             =   10
      TabIndex        =   1
      Top             =   1680
      Value           =   4
      Width           =   3000
   End
   Begin MultiCon.Boton BotonIzq 
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   315
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   "<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin MultiCon.Boton BotonDer 
      Height          =   195
      Index           =   5
      Left            =   3180
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   ">"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin MultiCon.Boton BotonDer 
      Height          =   195
      Index           =   4
      Left            =   3180
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4900
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   ">"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin MultiCon.Boton BotonDer 
      Height          =   195
      Index           =   3
      Left            =   3180
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   ">"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin MultiCon.Boton BotonDer 
      Height          =   195
      Index           =   2
      Left            =   3180
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   ">"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin MultiCon.Boton BotonDer 
      Height          =   195
      Index           =   0
      Left            =   3180
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Caption         =   ">"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9768745
      RoundSize       =   5
      Light           =   80
      Roundness3D     =   50
   End
   Begin VB.HScrollBar HScroll 
      Height          =   100
      Index           =   3
      Left            =   420
      Max             =   255
      TabIndex        =   21
      Top             =   4620
      Value           =   170
      Width           =   3000
   End
   Begin VB.CheckBox En 
      BackColor       =   &H00800000&
      Caption         =   "Buttons Enabled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   5760
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.HScrollBar HScroll 
      Height          =   100
      Index           =   5
      Left            =   420
      Max             =   255
      TabIndex        =   4
      Top             =   5220
      Value           =   170
      Width           =   3000
   End
   Begin VB.HScrollBar HScroll 
      Height          =   100
      Index           =   4
      Left            =   420
      Max             =   255
      TabIndex        =   3
      Top             =   4920
      Value           =   170
      Width           =   3000
   End
   Begin VB.HScrollBar HScroll 
      Height          =   100
      Index           =   2
      Left            =   420
      Max             =   10
      TabIndex        =   2
      Top             =   3240
      Value           =   7
      Width           =   3000
   End
   Begin VB.HScrollBar HScroll 
      Height          =   100
      Index           =   0
      Left            =   420
      Max             =   10
      TabIndex        =   0
      Top             =   360
      Value           =   2
      Width           =   3000
   End
   Begin MultiCon.Boton Boton 
      Height          =   1035
      Index           =   0
      Left            =   3960
      TabIndex        =   6
      Top             =   1020
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1826
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   14737632
   End
   Begin MultiCon.Boton Boton 
      Height          =   1035
      Index           =   1
      Left            =   5040
      TabIndex        =   7
      Top             =   1020
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1826
      Caption         =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   14737632
   End
   Begin MultiCon.Boton Boton 
      Height          =   1035
      Index           =   2
      Left            =   6120
      TabIndex        =   8
      Top             =   1020
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1826
      Caption         =   "3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   14737632
   End
   Begin MultiCon.Boton Boton 
      Height          =   1035
      Index           =   3
      Left            =   3960
      TabIndex        =   9
      Top             =   2100
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1826
      Caption         =   "4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   14737632
   End
   Begin MultiCon.Boton Boton 
      Height          =   1035
      Index           =   4
      Left            =   5040
      TabIndex        =   10
      Top             =   2100
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1826
      Caption         =   "5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   14737632
   End
   Begin MultiCon.Boton Boton 
      Height          =   1035
      Index           =   5
      Left            =   6120
      TabIndex        =   11
      Top             =   2100
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1826
      Caption         =   "6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   14737632
   End
   Begin MultiCon.Boton Boton 
      Height          =   1035
      Index           =   6
      Left            =   3960
      TabIndex        =   12
      Top             =   3180
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1826
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   14737632
   End
   Begin MultiCon.Boton Boton 
      Height          =   1035
      Index           =   7
      Left            =   5040
      TabIndex        =   13
      Top             =   3180
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1826
      Caption         =   "8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   14737632
   End
   Begin MultiCon.Boton Boton 
      Height          =   1035
      Index           =   8
      Left            =   6120
      TabIndex        =   14
      Top             =   3180
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1826
      Caption         =   "9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   14737632
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form2.frx":08DA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1020
      Index           =   3
      Left            =   3780
      TabIndex        =   39
      Top             =   0
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C0C0&
      Index           =   1
      X1              =   240
      X2              =   3780
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C0C0&
      Index           =   0
      X1              =   240
      X2              =   3780
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Defines how strong the 3D effect is lighted"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   540
      Index           =   2
      Left            =   540
      TabIndex        =   38
      Top             =   3780
      Width           =   2955
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Defines the 3D effect of the button. How much it's raised"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   540
      Index           =   1
      Left            =   480
      TabIndex        =   37
      Top             =   2220
      Width           =   2955
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Defines the rounded corners of the button"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   540
      Index           =   0
      Left            =   480
      TabIndex        =   36
      Top             =   900
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   3
      Left            =   3480
      TabIndex        =   22
      Top             =   4620
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   5
      Left            =   3480
      TabIndex        =   20
      Top             =   5220
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   3480
      TabIndex        =   19
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Light"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   2
      Left            =   540
      TabIndex        =   18
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3D Roundness"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   17
      Top             =   1920
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2D Roundness"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   16
      Top             =   600
      Width           =   1530
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   1020
      Left            =   5340
      TabIndex        =   15
      Top             =   4380
      Width           =   255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Boton_Click(Index As Integer)
    Label.Caption = Index + 1
End Sub

Private Sub BotonDer_Click(Index As Integer)
On Error Resume Next
    HScroll(Index).Value = HScroll(Index).Value + 1
End Sub

Private Sub BotonIzq_Click(Index As Integer)
On Error Resume Next
    HScroll(Index).Value = HScroll(Index).Value - 1
End Sub

Private Sub En_Click()
    Boton(0).Enabled = En.Value = 1
    Boton(2).Enabled = En.Value = 1
    Boton(4).Enabled = En.Value = 1
    Boton(6).Enabled = En.Value = 1
    Boton(8).Enabled = En.Value = 1
End Sub

Private Sub HScroll_Change(Index As Integer)
Dim x As Integer
    Select Case Index
    Case 0
        For x = 0 To 8
            Boton(x).RoundSize = HScroll(Index).Value * 10 - 1
        Next x
    Case 1
        For x = 0 To 8
            Boton(x).Roundness3D = HScroll(Index).Value * 5
        Next x
    Case 2
        For x = 0 To 8
            Boton(x).light = HScroll(Index).Value * 10 - 1
        Next x
    Case Else
        For x = 0 To 8
            Boton(x).BackColor = RGB(HScroll(3), HScroll(4), HScroll(5))
            Boton(x).ForeColor = RGB(HScroll(3) + 80, HScroll(4) + 80, HScroll(5) + 80)
        Next x
        Label.ForeColor = Boton(0).BackColor
    End Select

End Sub

Private Sub Ir_Click()
    Dim id As Long
    id = 34462
    Call ShellExecute(0&, vbNullString, "http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=" & id & "&lngWId=1", vbNullString, vbNullString, vbNormalFocus)
End Sub
