VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Scene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const LEVEL_RANGE       As String = "D15"
Const LINES_RANGE       As String = "D16"
Const SCORE_RANGE       As String = "D17"
Const HIGH_SCORE_RANGE  As String = "D19"

Private m_top           As Integer
Private m_left          As Integer
Private m_width         As Integer
Private m_height        As Integer

Private m_isLanding     As Boolean
Private m_values()      As Long
Private m_piece         As Piece
Private m_score         As Integer
Private m_lines         As Integer
Private m_level         As Integer


Private Sub Class_Initialize()
    Call Reset
End Sub

Public Sub InitScene( _
    ByVal sceneTop As Integer, ByVal sceneLeft As Integer, _
    ByVal sceneHeight As Integer, ByVal sceneWidth As Integer)

    m_top = sceneTop
    m_left = sceneLeft
    m_width = sceneWidth
    m_height = sceneHeight

    ReDim m_values(m_height, m_width) As Long

End Sub

Public Sub Reset()

    Dim i As Integer
    Dim j As Integer

    m_isLanding = True
    m_score = 0
    m_lines = 0
    m_level = 0
    
    For i = 0 To m_height - 1
        For j = 0 To m_width - 1
            m_values(i, j) = 0
        Next
    Next
    
End Sub

Public Sub HiddenPiece()

    Dim i As Integer

    If TypeName(m_piece) <> "Nothing" Then
        For i = 0 To m_piece.BlockCount - 1
            Cells( _
                m_top + m_piece.Y + m_piece.Blocks(i).Y, _
                m_left + m_piece.X + m_piece.Blocks(i).X _
            ).Value = 0
        Next
    End If
    
End Sub

Public Sub PaintPiece()

    Dim i As Integer

    If TypeName(m_piece) <> "Nothing" Then
    
        For i = 0 To m_piece.BlockCount - 1
            Cells( _
                m_top + m_piece.Y + m_piece.Blocks(i).Y, _
                m_left + m_piece.X + m_piece.Blocks(i).X _
            ).Value = m_piece.pieceType
        Next
    End If
    
End Sub

Public Sub PaintFixBlocks()

    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To m_height - 1
        For j = 0 To m_width - 1
            Cells(m_top + i, m_left + j).Value = m_values(i, j)
        Next
    Next
    
    'score update
    Range(SCORE_RANGE).Value = m_score
    Range(LINES_RANGE).Value = m_lines
    Range(LEVEL_RANGE).Value = m_level
    
    If Range(HIGH_SCORE_RANGE).Value <= m_score Then
        Range(HIGH_SCORE_RANGE).Value = m_score
    End If
    
End Sub

Public Function MoveDown() As Boolean

    Call m_piece.MoveDown
    
    If IsInvalid() Then
        Call m_piece.MoveUp
        m_isLanding = True
        MoveDown = False
        Exit Function
    End If
    
    MoveDown = True

End Function

Public Sub Fall()

    Do While True
        If MoveDown() = False Then Exit Do
    Loop
    
End Sub

Public Sub MoveLeft()

    Call m_piece.MoveLeft

    If IsInvalid() Then
        Call m_piece.MoveRight
        Exit Sub
    End If

End Sub

Public Sub MoveRight()

    Call m_piece.MoveRight

    If IsInvalid() Then
        Call m_piece.MoveLeft
        Exit Sub
    End If
    
End Sub

Public Sub Spin()

    Call m_piece.SpinRight

    If IsInvalid() Then
    
        'try only one block shift left
        Call m_piece.MoveLeft
        If IsInvalid() Then
            Call m_piece.MoveRight
        Else
            Exit Sub
        End If
        
        'try only one block shift right
        Call m_piece.MoveRight
        If IsInvalid() Then
            Call m_piece.MoveLeft
        Else
            Exit Sub
        End If
        
        'try only one block shift down
        Call m_piece.MoveDown
        If IsInvalid() Then
            Call m_piece.MoveUp
        Else
            Exit Sub
        End If
        
        Call m_piece.SpinLeft
        Exit Sub
    End If

End Sub

Public Function IsInvalid() As Boolean

    Dim i As Integer

    'check whether out of scene range
    For i = 0 To m_piece.BlockCount - 1
        If m_piece.Y + m_piece.Blocks(i).Y < 0 Or _
            m_piece.Y + m_piece.Blocks(i).Y >= m_height Or _
            m_piece.X + m_piece.Blocks(i).X < 0 Or _
            m_piece.X + m_piece.Blocks(i).X >= m_width Then
            
            IsInvalid = True
            Exit Function
        End If
    Next

    'check whether overlapping other blocks
    For i = 0 To m_piece.BlockCount - 1
        If m_values(m_piece.Y + m_piece.Blocks(i).Y, m_piece.X + m_piece.Blocks(i).X) <> 0 Then
            
            IsInvalid = True
            Exit Function
        End If
    Next
    
    IsInvalid = False

End Function

Public Sub FixPiece()

    Dim i As Integer
    Dim j As Integer

    Dim eraseFlg As Boolean
    Dim eraseCount As Integer
    
    m_isLanding = True

    For i = 0 To m_piece.BlockCount - 1
        m_values(m_piece.Y + m_piece.Blocks(i).Y, m_piece.X + m_piece.Blocks(i).X) = m_piece.pieceType
    Next
    
    'Erase
    eraseCount = 0
    For i = m_height To 0 Step -1

        eraseFlg = True
        Do While eraseFlg

            For j = 0 To m_width - 1
                If m_values(i, j) = 0 Then
                    eraseFlg = False
                    Exit For
                End If
            Next
            If eraseFlg Then
                Call EraseLine(i)
                eraseCount = eraseCount + 1
            End If
        Loop
            
    Next
    
    'Score Update
    m_lines = m_lines + eraseCount
    If eraseCount = 1 Then
        m_score = m_score + 100
        
    ElseIf eraseCount = 2 Then
        m_score = m_score + 250
        
    ElseIf eraseCount = 3 Then
        m_score = m_score + 400
        
    ElseIf eraseCount = 4 Then
        m_score = m_score + 600
        
    End If
    
    Call PaintFixBlocks
    
End Sub

Public Sub EraseLine(idx As Integer)

    Dim i As Integer
    Dim j As Integer

    For i = idx To 1 Step -1
        For j = 0 To m_width - 1
            m_values(i, j) = m_values(i - 1, j)
        Next
    Next
    
    'Clear Top Line
    For j = 0 To m_width - 1
        m_values(0, j) = 0
    Next

End Sub

Public Function SetNewBlock(Optional pieceType As Integer = 0) As Boolean

    Dim i As Integer

    Set m_piece = New Piece
    Call m_piece.InitBlocks(pieceType)
    Call m_piece.SetPosition(0, m_width / 2)
    
    m_isLanding = False
    
    'Check dead
    For i = 0 To m_piece.BlockCount - 1
        If m_values(m_piece.Y + m_piece.Blocks(i).Y, m_piece.X + m_piece.Blocks(i).X) <> 0 Then
            SetNewBlock = False
            Exit Function
        End If
    Next
        
    SetNewBlock = True

End Function

Public Sub LevelUp()

    m_level = m_level + 1
    
End Sub

Public Function IsLanding()

    IsLanding = m_isLanding

End Function

Property Get Piece() As Piece
    Set Piece = m_piece
End Property

Property Get Score() As Integer
     Score = m_score
End Property

Property Get Lines() As Integer
     Lines = m_lines
End Property

Property Get Level() As Integer
     Level = m_level
End Property


