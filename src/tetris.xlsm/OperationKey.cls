VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OperationKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Integer

Option Explicit

Enum KEY_STATE
    RELEASED = 0
    PRESSED = 1
    LONG_PRESSED = 2
End Enum

Private m_key               As Integer
Private m_pressStartTime    As Double

'Constructor
Private Sub Class_Initialize()
    m_pressStartTime = 0
    m_key = vbKeySpace
End Sub

Public Sub SetKey(ByVal key As Integer)
    m_key = key
End Sub

Property Get State() As Integer
    'key is being pressed
    If GetAsyncKeyState(m_key) <> 0 Then
        'first pressed
        If m_pressStartTime = 0 Then
            m_pressStartTime = timer
            State = PRESSED

        'delay decision of long press for 0.2 seconds
        ElseIf timer <= m_pressStartTime + 0.2 Then
            State = RELEASED

        'long pressed
        Else
            State = LONG_PRESSED
        End If

    'key is not being pressed
    Else
        m_pressStartTime = 0
        State = RELEASED
    End If
End Property
