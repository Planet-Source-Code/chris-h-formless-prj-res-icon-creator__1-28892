VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formless Prj *.RES Icon Creator"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12210
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Height          =   5655
      Left            =   143
      TabIndex        =   2
      Top             =   120
      Width           =   11535
      Begin VB.Frame fraWiz 
         Caption         =   "0"
         Height          =   3600
         Index           =   0
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   5295
         Begin VB.Label Label1 
            Caption         =   $"frmMain.frx":1D42
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2760
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   720
            Width           =   4635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "INFORMATION:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1725
         End
      End
      Begin VB.Frame fraWiz 
         Caption         =   "4"
         Height          =   3650
         Index           =   4
         Left            =   840
         TabIndex        =   21
         Top             =   1440
         Width           =   5295
         Begin prjFormlessWizard.xpcmdbutton cmdPaths 
            Height          =   375
            Index           =   4
            Left            =   4080
            TabIndex        =   31
            Top             =   3240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Caption         =   "Create"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "tech@ets4u.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   13
            Left            =   1065
            TabIndex        =   33
            Top             =   3105
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "tech@ets4u.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   18
            Left            =   1080
            TabIndex        =   37
            Top             =   3120
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chris Hoffman"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   17
            Left            =   1080
            TabIndex        =   36
            Top             =   2865
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Chris Hoffman"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   16
            Left            =   1095
            TabIndex        =   35
            Top             =   2880
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contact:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Index           =   15
            Left            =   720
            TabIndex        =   34
            Top             =   2475
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Questions? Comments?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   11
            Left            =   720
            TabIndex        =   32
            Top             =   2040
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   $"frmMain.frx":1F4A
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Index           =   14
            Left            =   360
            TabIndex        =   30
            Top             =   840
            Width           =   4695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "STEP 4 (CREATION):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   12
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   2310
         End
      End
      Begin VB.Frame fraWiz 
         Caption         =   "3"
         Height          =   3650
         Index           =   3
         Left            =   720
         TabIndex        =   7
         Top             =   1200
         Width           =   5295
         Begin VB.CheckBox chkPrj 
            Caption         =   "Insert New *.RES into existing VB project?"
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   840
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin prjFormlessWizard.xpcmdbutton cmdPaths 
            Height          =   255
            Index           =   2
            Left            =   4920
            TabIndex        =   19
            Top             =   2160
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   960
            Index           =   10
            Left            =   120
            TabIndex        =   27
            Top             =   2520
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.Label Label1 
            Caption         =   "   The new resource file will be located in the same path as the project file."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   600
            Index           =   9
            Left            =   360
            TabIndex        =   25
            Top             =   2880
            Width           =   4575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "FYI:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Index           =   8
            Left            =   360
            TabIndex        =   24
            Top             =   2640
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "STEP 3 (VB PROJECT):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   7
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   2565
         End
         Begin VB.Label Label1 
            Caption         =   "   Now you need to pick the project that the new *.RES file will be inserted into."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   6
            Left            =   360
            TabIndex        =   22
            Top             =   1320
            Width           =   4695
         End
         Begin VB.Label lblPrjPath 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   2160
            Width           =   4575
         End
      End
      Begin VB.Frame fraWiz 
         Caption         =   "2"
         Height          =   3495
         Index           =   2
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   5295
         Begin VB.PictureBox picIcon 
            Height          =   765
            Left            =   2160
            Picture         =   "frmMain.frx":2012
            ScaleHeight     =   705
            ScaleWidth      =   705
            TabIndex        =   18
            Top             =   1680
            Width           =   765
         End
         Begin prjFormlessWizard.xpcmdbutton cmdPaths 
            Height          =   255
            Index           =   1
            Left            =   4920
            TabIndex        =   14
            Top             =   2760
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "STEP 2 (ICON):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   5
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label Label1 
            Caption         =   "   Now you need to pick the icon that you would like to use for the resource file / project."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   4
            Left            =   360
            TabIndex        =   16
            Top             =   840
            Width           =   4695
         End
         Begin VB.Label lblIcon 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   2760
            Width           =   4575
         End
      End
      Begin VB.Frame fraWiz 
         Caption         =   "1"
         Height          =   3495
         Index           =   1
         Left            =   480
         TabIndex        =   5
         Top             =   720
         Width           =   5295
         Begin prjFormlessWizard.xpcmdbutton cmdPaths 
            Height          =   255
            Index           =   0
            Left            =   4920
            TabIndex        =   13
            Top             =   2640
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin prjFormlessWizard.xpcmdbutton cmdPaths 
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   28
            Top             =   3100
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   661
            Caption         =   "Extract RC.EXE and RCDLL.DLL To Apps Path."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblRC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   2640
            Width           =   4575
         End
         Begin VB.Label Label1 
            Caption         =   $"frmMain.frx":3D54
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1680
            Index           =   3
            Left            =   360
            TabIndex        =   11
            Top             =   840
            Width           =   4695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "STEP 1 (COMPILER):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   2325
         End
      End
   End
   Begin prjFormlessWizard.xpcmdbutton cmdBut 
      Height          =   495
      Index           =   0
      Left            =   3900
      TabIndex        =   0
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      Caption         =   "<<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjFormlessWizard.xpcmdbutton cmdBut 
      Height          =   495
      Index           =   1
      Left            =   4800
      TabIndex        =   1
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      Caption         =   ">>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjFormlessWizard.xpcmdbutton cmdBut 
      Height          =   495
      Index           =   2
      Left            =   165
      TabIndex        =   3
      Top             =   5520
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   873
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'    cSoft Terms of Usage:
'
'    By using this code, you agree to the following terms...
'
'    1) You may use this code in your own programs
'    (and may compile it into a program and distribute it
'    in compiled format for langauges that allow it) freely and with no charge.
'
'    2) If you do use this code for profit, an mention of the Author and
'    Company name would be more than appreciated.
'
'    3) You MAY NOT redistribute this code without written
'    permission from the original author. Failure to do so is a violation of copyright laws.
'
'    4) In Otherwords, Don't Screw ME! It isn't necessary,
'    Im just looking for a LIL Recognition, Wouldn't you?
'
'    Copyright:        © 2000 cSoft.
'    AUTHOR:           Chris Hoffman, cSoft
'    AUTHORS EMAIL:    tech@ets4u.com
'    AUTHORS WEBSITE:  http://www.ets4u.com
'
'        Project Type:  Programmers Utility
'
'        What does this Prj Do?
'                       Real simple, ever notice that after you have created
'                       that aswesome EXE or ActiveX EXE or any kind of FORMLESS prj,
'                       that when compiled it uses the DEFAULT VB ICON!!
'                       How shitty!, This app encomposes the RC.EXE file that
'                       is included with Vis Studio Enterprise edition
'                       (Although Im sure its a free D/L womewhere :O)
'                       It Creates, Inserts an RES file to your specified
'                       Prj, so that when you COMPILE the prj it will use the
'                       ICON that you choose! No more Defaults!
'
'
'        Dependents  :  You need RC.EXE, however I have
'                       included them in the RES file, which can
'                       be extracted automatically by the wizard.
'
'        References  :  na
'
'        Thanx Too   :  VB ACCELERATOR, This is where i figured
'                       this out, then i built this app to make it a lil easier.
'
'        FUTURE PLANS:  none


'
'   Wizard Cls
Private WithEvents cWiz As CWizard
Attribute cWiz.VB_VarHelpID = -1
Private cFile As clsFileAPI
Private m_Last As Boolean
Private m_UsePrj As Boolean



Private Sub Form_Load()
'
'   Frm Specific stuff
    Call InitGUI
    Call InitWiz(True)
    m_UsePrj = True
    Call FindRCEXE
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
'   CleanUp
'
    Call InitWiz(False)

End Sub



Private Sub cmdBut_Click(Index As Integer)

    Select Case Index
        '
        '   Back
        Case 0:  cWiz.MoveBack: If Not m_Last Then cmdBut(1).Enabled = True
        '
        '   Forward
        Case 1: cWiz.MoveNext: If m_Last Then cmdBut(1).Enabled = False
        '
        '   Cancel and Unload
        Case 2: Unload Me
    End Select
    
End Sub



Private Sub cmdPaths_Click(Index As Integer)
'
'   Paths and such
'
    '   Proc Vars
    Dim sFile As String
    
    Select Case Index
    
        Case 0
            '
            '   Location of RC.EXE
            If VBGetOpenFileName( _
                sFile, , , , , , _
                "Resource Compiler (*.EXE)|*.EXE|Executables (*.EXE)|*.EXE)", _
                1, , "Choose Your RC.EXE File", "EXE", Me.hWnd _
                ) Then
                lblRC.Tag = sFile
                Debug.Print lblRC.Tag
                lblRC.ToolTipText = sFile
                lblRC.Caption = sFile
                lblRC.Caption = CompactedPathSh( _
                    sFile, lblRC.Width \ Screen.TwipsPerPixelX - 20, Me.hDC)

            End If

        Case 1
            '
            '   Specified Icon
            If VBGetOpenFileName( _
                sFile, , , , , , _
                "Res ICON (*.ICO)|*.ICO|Icon Files", _
                1, , "Choose Your ICON File", "ICO", Me.hWnd _
                ) Then
                lblIcon.Tag = sFile
                Debug.Print lblIcon.Tag
                lblIcon.ToolTipText = sFile
                lblIcon.Caption = CompactedPathSh( _
                    sFile, lblIcon.Width \ Screen.TwipsPerPixelX - 20, Me.hDC)
                Set picIcon.Picture = LoadPicture(sFile)
            End If
        
        Case 2
            '
            '   Specified Prj File
            If m_UsePrj Then
                If VBGetOpenFileName( _
                    sFile, , , , , , _
                    "VB Prj File (*.VBP)|*.VBP|Visual Basic Project Files", _
                    1, , "Choose your Project File", "VBP", Me.hWnd _
                    ) Then
                    lblPrjPath.Tag = sFile
                    Debug.Print lblPrjPath.Tag
                    lblPrjPath.ToolTipText = sFile
                    lblPrjPath.Caption = CompactedPathSh( _
                        sFile, lblPrjPath.Width \ Screen.TwipsPerPixelX - 20, Me.hDC)
                End If

            Else
                '
                '   Saving to Res instead
                If VBGetSaveFileName(sFile, , , "VB RES File(*.RES|*.RES|Visual Basic Resource File", , , "Save to *.RES", "RES", Me.hWnd) Then
                    lblPrjPath.Tag = sFile
                    Debug.Print lblPrjPath.Tag
                    lblPrjPath.ToolTipText = sFile
                    lblPrjPath.Caption = CompactedPathSh( _
                        sFile, lblPrjPath.Width \ Screen.TwipsPerPixelX - 20, Me.hDC)
                End If
            End If
            
        Case 3
            '
            '   Extract the Included Rc files to apps path
            '   This is an excellent (albeit dangerous piece of
            '   Code, meaning there are assholes out there that
            '   Would abuse this, (Extract and run a virus for example,
            '   Dont be a dick, dont abuse it))
            sFile = QS(App.Path)
            If Not LoadDataIntoFile(101, sFile & "RC.EXE") And _
                Not LoadDataIntoFile(102, sFile & "RCDLL.DLL") Then
                
                MsgBox "Couldnt extract, Try finding the files at Microsoft"
            Else
                cmdPaths(3).Enabled = False
                lblRC.Tag = sFile & "RC.EXE"
                Debug.Print lblRC.Tag
                lblRC.ToolTipText = sFile & "RC.EXE"
                lblRC.Caption = CompactedPathSh( _
                        sFile & "RC.EXE", lblRC.Width \ Screen.TwipsPerPixelX - 20, Me.hDC)
            End If
        
        Case 4: Call CreatRC
        
    End Select

End Sub



Private Sub chkPrj_Click()
'
'   If user wants to use an exisitng Prj
'   Or maybe just create a res for later use.
'
    m_UsePrj = chkPrj.Value
    Label1(10).Visible = Not m_UsePrj
    
    If m_UsePrj Then
        Label1(6).Caption = "   Now you need to pick the project that the new *.RES file will be inserted into."
    Else
        Label1(6).Caption = " Ok, Since not using an Project, then you need to specify where you want to save the New *.Res file."
    End If

End Sub



Private Sub InitWiz( _
    bLoad As Boolean _
)
'
'   cWiz Stuff
'
    If bLoad Then
        '
        '   Create
        Set cWiz = New CWizard
        '
        '   Frame count
        cWiz.StepCount = fraWiz.Count
        cWiz.Step = 0
    Else
        '
        '   Clean Up
        Set cWiz = Nothing
    End If
    
End Sub

Private Sub cWiz_DisplayStep()
    '
    '   Placement, Zorder
    fraWiz(cWiz.Step).ZOrder 0
    If cWiz.Step = fraWiz.Count - 1 Then _
        m_Last = True Else m_Last = False
    
End Sub



Private Sub InitGUI( _
)
'
'   Placement of frames,Buttons etc.
'
    '   Proc Vars
    Dim i As Byte
    '
    '   Frm
    Me.Width = 5910
    Me.Height = 5200
    '
    '   Main Frame (Container)
    fraMain.left = 143
    fraMain.top = 120
    fraMain.Width = 5535
    fraMain.Height = 3855
    '
    '   fraWiz Frames
    For i = 0 To fraWiz.Count - 1
        '
        '   Fewer Commands so...
        On Error Resume Next

        fraWiz(i).BorderStyle = 0
        fraWiz(i).left = 120
        fraWiz(i).top = 120
        fraWiz(i).Height = 3650
        fraWiz(i).Width = 5295
        '
        '   cmd's
        cmdBut(i).top = 4080
    Next i

End Sub


Private Sub FindRCEXE( _
)
'
'   If its not in the default path
'   Then we need to enable the browse button
'
    If QualifyPath("C:\Program Files\Microsoft Visual Studio\VB98\Wizards\rc.exe") Then
    
        Label1(3).Caption = "Your RC.EXE has been found. No need to locate it."
        lblRC.Tag = "C:\Program Files\Microsoft Visual Studio\VB98\Wizards\rc.exe"
        lblRC.ToolTipText = "C:\Program Files\Microsoft Visual Studio\VB98\Wizards\rc.exe"
        lblRC.Caption = CompactedPathSh( _
                "C:\Program Files\Microsoft Visual Studio\VB98\Wizards\rc.exe", _
                lblRC.Width \ Screen.TwipsPerPixelX - 20, Me.hDC)
        cmdPaths(3).Enabled = False
        cmdPaths(0).Enabled = False
        
    Else
        cmdPaths(0).Enabled = True
        cmdPaths(3).Enabled = True
        Label1(3).Caption = "   We need to locate your resource compiler, " & _
            " Usually this is found in the Wizards folder of your of your" & _
            " VB98 folder. The file we are looking for is, ""RC.EXE"". It would" & _
            " seem that RC.EXE only comes with the Enterprise edition of VB," & _
            " Ive packed it into this program, If you dont have it, click the" & _
            " Extract button to save the required files to this apps path."
        lblRC.Caption = vbNullString
    End If
    
End Sub




Private Sub CreatRC( _
)
'
'   Create the *.RC File for *.RES
'
    '
    '   Note, I opted to copy Each file to the Root of
    '   The C drive... Why? rather than deal with DOS
    '   Paths, RC.EXE only works with DOS paths, This seems a bit easier)
    '
    '   Proc Vars
    Dim lRet As Double
    Dim lCount As Double
    Dim lRetry As Long
    Dim sTmp As String
    Dim tmpA() As String
    Dim lB As Long
    Dim UB As Long
    Dim i As Long
    Dim bFound As Boolean

    On Error GoTo ER
    '
    '   Show Busy, althought if all goes well
    '   This shouldnt take more than a few MS
    Me.MousePointer = vbHourglass
    '
    '   Init cls
    Set cFile = New clsFileAPI
    cFile.OpenAPI "c:\tmp.rc"
    '
    '   ZERO tells VB that this ICON will be the APPS Icon
    cFile.WriteAPI "0   ICON   MOVEABLE   PRELOAD   c:\tmp.ico"
    Set cFile = Nothing
    '
    '   Copy Icon
    FileCopy lblIcon.Tag, "c:\tmp.ico"
    '
    '   Make the Actual RES
    lRet = Shell(lblRC.Tag & " /r /fo c:\tmp.res c:\tmp.rc", vbNormalFocus)
    '
    '   Since the Process of making seems to take a bit
    '   (Sometimes, im just putting in a loop so to WAIT)
    Do Until lRet <> 0
        '
        '   So not too loop forever
        lCount = lCount + 1
        DoEvents
    Loop
    '
    '   If user just wants to MAKE the file
    '   And NOT insert it into there project.
    If Not m_UsePrj Then
        FileCopy "c:\tmp.res", lblPrjPath.Tag
    Else
        '
        '   Path Only
        sTmp = ParsePath(lblPrjPath.Tag)
        FileCopy "c:\tmp.res", QS(sTmp) & "PRJ.RES"
        '
        '   Init New Cls
        Set cFile = New clsFileAPI
        sTmp = ParseFileName(lblPrjPath.Tag)
        cFile.OpenAPI lblPrjPath.Tag
        '
        '   Read the Contents and fill our Var
        cFile.ReadAPI sTmp
        '
        '   VB6+ Only
        tmpA = Split(sTmp, vbCrLf)
        lB = LBound(tmpA)
        UB = UBound(tmpA)
        '
        '   Loop and fill
        For i = lB To UB - 1
            '
            '   If the user has specified an RES already
            '   Then we will replace the Entry, but NOT
            '   Delete there Original RES file.
            If InStr(1, LCase(tmpA(i)), "resfile32") <> 0 Then
                tmpA(i) = "ResFile32=""PRJ.RES"""
                '
                '   Set Flag
                bFound = True
            End If
        Next i
        '
        '   Close is important, although the class Terminate
        '   Does this, its just a good habbit to get into.
        cFile.CloseAPI
        If bFound Then
            '
            '   Join the array element to element
            sTmp = Join(tmpA, vbCrLf)
        Else
            '
            '   Didnt find an existing so we will
            '   Just put the new entry at top of file.
            sTmp = "ResFile32=""PRJ.RES""" & vbCrLf & Join(tmpA, vbCrLf)
        End If
        '
        '   Kill the Orig Prj File
        Kill lblPrjPath.Tag
        '
        '   Create the New
        cFile.OpenAPI lblPrjPath.Tag
        '
        '   Fill it
        cFile.WriteAPI sTmp
        '
        '   Again, close
        cFile.CloseAPI
    End If
    '
    '   Clean Up class
    If Not cFile Is Nothing Then Set cFile = Nothing
    '
    '   Clean up our Tmp files
    Kill "c:\tmp.rc"
    Kill "c:\tmp.ico"
    Kill "c:\tmp.res"
    '
    '   Reset Mouse
    Me.MousePointer = vbNormal
    '
    '   Jet if all is well!
    Exit Sub
ER:
    '
    '   1000? excesive? seems ok to me
    lRetry = lRetry + 1
    If lRetry > 1000 Then
        MsgBox "Error has occured, please check your paths?"
    Else
        Resume
    End If
    Me.MousePointer = vbNormal

End Sub

