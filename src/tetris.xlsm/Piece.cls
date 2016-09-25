VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Piece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum PIECE_TYPE
    TYPE_I = 1
    TYPE_O = 2
    TYPE_S = 3
    TYPE_Z = 4
    TYPE_J = 5
    TYPE_L = 6
    TYPE_T = 7

End Enum

Private m_positionY         As Integer
Private m_positionX         As Integer
Private m_pieceType         As Long
Private m_blocks(4)         As Block

'Constructor
Private Sub Class_Initialize()

    m_positionY = 0
    m_positionX = 0


    Dim i As Integer
    For i = 0 To 4
        Set m_blocks(i) = New Block
    Next

End Sub

Public Sub InitBlocks(Optional pieceType As Integer = 0)

    If pieceType = 0 Then
        m_pieceType = Int(6 * Rnd + 1)
    
    Else
        m_pieceType = pieceType
        
    End If
    
    
    '■■
    '■■
    If m_pieceType = PIECE_TYPE.TYPE_O Then
        Call m_blocks(0).SetDist(0, -1)
        Call m_blocks(1).SetDist(0, 0)
        Call m_blocks(2).SetDist(1, -1)
        Call m_blocks(3).SetDist(1, 0)
        
    '■
    '■
    '■
    '■
    ElseIf m_pieceType = PIECE_TYPE.TYPE_I Then
        Call m_blocks(0).SetDist(0, -2)
        Call m_blocks(1).SetDist(0, -1)
        Call m_blocks(2).SetDist(0, 0)
        Call m_blocks(3).SetDist(0, 1)
        
    '　■■
    '■■
    ElseIf m_pieceType = PIECE_TYPE.TYPE_S Then
        Call m_blocks(0).SetDist(1, -1)
        Call m_blocks(1).SetDist(1, 0)
        Call m_blocks(2).SetDist(0, 0)
        Call m_blocks(3).SetDist(0, 1)

    '■■
    '　■■
    ElseIf m_pieceType = PIECE_TYPE.TYPE_Z Then
        Call m_blocks(0).SetDist(0, -1)
        Call m_blocks(1).SetDist(0, 0)
        Call m_blocks(2).SetDist(1, 0)
        Call m_blocks(3).SetDist(1, 1)

    '■■■
    '　　■
    ElseIf m_pieceType = PIECE_TYPE.TYPE_J Then
        Call m_blocks(0).SetDist(0, -1)
        Call m_blocks(1).SetDist(0, 0)
        Call m_blocks(2).SetDist(0, 1)
        Call m_blocks(3).SetDist(1, 1)

    '■■■
    '■
    ElseIf m_pieceType = PIECE_TYPE.TYPE_L Then
        Call m_blocks(0).SetDist(1, -1)
        Call m_blocks(1).SetDist(0, -1)
        Call m_blocks(2).SetDist(0, 0)
        Call m_blocks(3).SetDist(0, 1)

    '■■■
    '　■
    ElseIf m_pieceType = PIECE_TYPE.TYPE_T Then
        Call m_blocks(0).SetDist(0, -1)
        Call m_blocks(1).SetDist(0, 0)
        Call m_blocks(2).SetDist(1, 0)
        Call m_blocks(3).SetDist(0, 1)

    End If
    
End Sub

Public Sub SetPosition(ByVal Y As Integer, ByVal X As Integer)
    m_positionY = Y
    m_positionX = X
End Sub

Public Sub MoveDown()
    m_positionY = m_positionY + 1
End Sub

Public Sub MoveUp()
    m_positionY = m_positionY - 1
End Sub

Public Sub MoveLeft()
    m_positionX = m_positionX - 1
End Sub

Public Sub MoveRight()
    m_positionX = m_positionX + 1
End Sub

Public Sub SpinRight()

    If m_pieceType <> PIECE_TYPE.TYPE_O Then
        Call replaceYX
        Call InvertX
    End If
    
End Sub

Public Sub SpinLeft()

    If m_pieceType <> PIECE_TYPE.TYPE_O Then
        Call InvertX
        Call replaceYX
    End If
    
End Sub

Property Get Y() As Integer
     Y = m_positionY
End Property

Property Get X() As Integer
     X = m_positionX
End Property

Property Get Blocks(idx As Integer) As Block
     Set Blocks = m_blocks(idx)
End Property

Property Get pieceType() As Long
     pieceType = m_pieceType
End Property

Property Get BlockCount() As Integer
     BlockCount = UBound(m_blocks)
End Property

Private Sub replaceYX()
    Dim i As Integer
    For i = 0 To UBound(m_blocks)
        Call m_blocks(i).replaceYX
    Next
End Sub

Private Sub InvertY()
    Dim i As Integer
    For i = 0 To UBound(m_blocks)
        Call m_blocks(i).InvertY
    Next
End Sub

Private Sub InvertX()
    Dim i As Integer
    For i = 0 To UBound(m_blocks)
        Call m_blocks(i).InvertX
    Next
End Sub



