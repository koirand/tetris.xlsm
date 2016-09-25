Attribute VB_Name = "Main"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Option Explicit

Const BUTTON1_RANGE As String = "C20"

Enum STATUS
    STOPPING = 0
    RUNNING = 1
    PAUSING = 2
End Enum

Dim mainScene       As Scene
Dim nextBlockScene  As Scene
Dim leftKey         As OperationKey
Dim rightKey        As OperationKey
Dim downKey         As OperationKey
Dim SpinKey         As OperationKey
Dim FallKey         As OperationKey

Dim fallSpeed       As Single
Dim nextFallTime    As Single
Dim levelUpSpeed    As Single
Dim nextLevelUpTime As Single

Dim gameStatus As Integer

Sub StartButtonClicked()

    If gameStatus = STOPPING Then
    
        gameStatus = RUNNING
        Range(BUTTON1_RANGE) = "Pause"

        Call fixCusols

        fallSpeed = 1
        nextFallTime = timer() + fallSpeed

        levelUpSpeed = 60
        nextLevelUpTime = timer() + levelUpSpeed

        'create main scene
        Set mainScene = New Scene
        Call mainScene.InitScene(1, 7, 22, 10)

        'create next block scene
        Set nextBlockScene = New Scene
        Call nextBlockScene.InitScene(5, 2, 4, 4)

        'create OperationKey objects
        Set leftKey = New OperationKey
        Set rightKey = New OperationKey
        Set downKey = New OperationKey
        Set SpinKey = New OperationKey
        Set FallKey = New OperationKey

        'set operation key
        Call leftKey.SetKey(vbKeyLeft)
        Call rightKey.SetKey(vbKeyRight)
        Call downKey.SetKey(vbKeyDown)
        Call SpinKey.SetKey(vbKeyUp)
        Call FallKey.SetKey(vbKeySpace)
        
        'set new block
        Call mainScene.SetNewBlock
        Call mainScene.PaintPiece
        Call nextBlockScene.SetNewBlock
        Call nextBlockScene.PaintPiece


    ElseIf gameStatus = PAUSING Then
        gameStatus = RUNNING
        Range(BUTTON1_RANGE) = "Pause"
        Call fixCusols
        
    ElseIf gameStatus = RUNNING Then
        gameStatus = PAUSING
        Range(BUTTON1_RANGE) = "Rasume"
        Call unfixCusols
    
    End If

    Do While gameStatus = RUNNING

        'disable screen updating
        Application.ScreenUpdating = False
    
        'when piece has landing
        If mainScene.IsLanding Then
        
            Call mainScene.FixPiece
            
            'create new block in main scene
            If mainScene.SetNewBlock(nextBlockScene.Piece.pieceType) = False Then
                Call MsgBox("Game Over")
                Exit Do
            End If
            
            'create new block in next block scene
            Call nextBlockScene.HiddenPiece
            Call nextBlockScene.SetNewBlock
            Call nextBlockScene.PaintPiece
            
            '0.5 second interval
            Call Sleep(500)
        
        End If
            
        'hidden piece before move
        mainScene.HiddenPiece
        
        'key operation
        If leftKey.State > 0 Then Call mainScene.MoveLeft
        If rightKey.State > 0 Then Call mainScene.MoveRight
        If downKey.State > 0 Then Call mainScene.MoveDown
        If SpinKey.State > 0 Then Call mainScene.Spin
        If FallKey.State > 0 Then Call mainScene.Fall
        
        'move down at regular time intervals
        If timer() >= nextFallTime Then
            mainScene.MoveDown
            nextFallTime = timer() + fallSpeed
        End If
                
        'paint piece
        mainScene.PaintPiece
            
        
        'level up at regular time intervals
        If timer() >= nextLevelUpTime Then
            fallSpeed = fallSpeed * 0.8
            Call mainScene.LevelUp
            'reset timer
            nextLevelUpTime = timer() + levelUpSpeed
        End If
        
        'enable screen updating
        Application.ScreenUpdating = True
        
        DoEvents
        Call Sleep(1)
    
    Loop

End Sub

Sub ResetButtonClicked()

    'disable screen updating
    Application.ScreenUpdating = False
    
    gameStatus = STOPPING
    Range(BUTTON1_RANGE) = "Start"

    If TypeName(nextBlockScene) <> "Nothing" Then
        Call nextBlockScene.Reset
        Call nextBlockScene.PaintFixBlocks
    End If
    
    If TypeName(mainScene) <> "Nothing" Then
        Call mainScene.Reset
        Call mainScene.PaintFixBlocks
    End If
    
    Call unfixCusols
    
    'enable screen updating
    Application.ScreenUpdating = True

End Sub

Sub fixCusols()
    Application.OnKey "{UP}", "DoNothing"
    Application.OnKey "{DOWN}", "DoNothing"
    Application.OnKey "{LEFT}", "DoNothing"
    Application.OnKey "{RIGHT}", "DoNothing"
    Application.OnKey " ", "DoNothing"
End Sub

Sub unfixCusols()
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey " "
End Sub

Sub DoNothing()

    'do nothing
    
End Sub

