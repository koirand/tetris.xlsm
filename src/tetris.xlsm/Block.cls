VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Block"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_distY     As Integer
Private m_distX     As Integer

'Constructor
Private Sub Class_Initialize()

End Sub

Public Sub SetDist(Y As Integer, X As Integer)
    m_distY = Y
    m_distX = X
End Sub

Property Let X(arg As Integer)
     m_positionX = arg
End Property

Property Let Y(arg As Integer)
     m_positionY = arg
End Property

Property Get X() As Integer
     X = m_distX
End Property

Property Get Y() As Integer
     Y = m_distY
End Property

Public Sub replaceYX()
    Dim tmpDist As Integer
    tmpDist = m_distY
    m_distY = m_distX
    m_distX = tmpDist

End Sub

Public Sub InvertY()
    m_distY = -m_distY
End Sub

Public Sub InvertX()
    m_distX = -m_distX
End Sub

