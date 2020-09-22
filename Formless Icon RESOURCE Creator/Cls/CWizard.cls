VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================================================================
' Name:     CWizard
' Purpose:  Encapsulate the general functionality of a wizard.
' Programmer: Andrew Davey
' Email:    andrewdavey@hotmail.com
'======================================================================================
' Usage:    (See frmTest for example)
'       Create a new form.
'       Add Back and Next Command Buttons.
'       Add Frame controls for each step. (Call them all fraStep, and make them a control array.)
'       Add other controls for each step.

'       Add the line, "Private WithEvents m_Wizard as CWizard",
'       to the declarations section of code.

'       Instance m_Wizard in Form_Load. (Set m_Wizard = New CWizard)
'       Set m_Wizard.StepCount = number of steps you want.
'       Set m_Wizard.Step = 0 (This will call the display step event).

'       Destroy m_Wizard in Form_Unload. (Set m_Wizard = Nothing)

'       Add "m_Wizard.MoveNext", in Next button click event handler.
'       Add "m_Wizard.MoveBack", in Back button click event handler.

'       Add code to show fraStep(m_Wizard.Step) and change buttons etc
'       in m_Wizard_DisplayStep event handler.

'       Add validation code in "m_Wizard_Validate" to check user input.
'       Set CancelMove = True to stop move if input is bad.
'       Also, add code to do 'stuff' here. Open files, run functions, etc.

'       Add code to the "m_Wizard_Finish" event handler, eg close form.
'       That's it, RUN!

Option Explicit

' The raisable events.
Public Event Validate(ByRef CancelMove As Boolean)
Public Event DisplayStep()
Public Event Finish()

' Private property holding variables.
Private m_Step As Integer   ' Range: 0 To StepCount - 1
Private m_StepCount As Integer

' Public Property access methods.

' --- Step ---
Public Property Get Step() As Integer
    Step = m_Step
End Property

Public Property Let Step(nStep As Integer)
    If nStep < 0 Or nStep >= m_StepCount Then
        Exit Property
    Else
        m_Step = nStep
        RaiseEvent DisplayStep
    End If
End Property
' --- End Step ---

' --- StepCount ---
Public Property Get StepCount() As Integer
    StepCount = m_StepCount
End Property

Public Property Let StepCount(nStepCount As Integer)
    If nStepCount < 0 Then
        Exit Property
    Else
        m_StepCount = nStepCount
    End If
End Property
' --- End StepCount ---

' Public Methods.

Public Sub MoveNext()
    Dim bCancel As Boolean
    
    RaiseEvent Validate(bCancel) ' bCancel might be changed by event_proc.
    If bCancel Then ' If set to true then cancel the move.
        Exit Sub
    Else
        ' Next step
        m_Step = m_Step + 1
        If m_Step >= m_StepCount Then ' - 1 Then
            RaiseEvent Finish
            m_Step = m_StepCount - 1 ' Don't move on.
            Exit Sub
        End If
        
        RaiseEvent DisplayStep ' Let owner update display for new step.
    End If
End Sub

Public Sub MoveBack()
    m_Step = m_Step - 1
    If m_Step < 0 Then m_Step = 0 ' Can't go back any more!
    RaiseEvent DisplayStep
End Sub

