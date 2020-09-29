Public Class Form1
    'Robert Sirota
    ' 0,0  0,1   0,2
    ' 1,0  1,1   1,2
    ' 2,0  2,1   2,2
    Private Sub btnNoveMade_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn00.Click, btn01.Click, btn02.Click, btn10.Click, _
    btn11.Click, btn12.Click, btn22.Click, btn20.Click, btn21.Click
        Dim btnSquareClicked As Button = sender
        Static chrTTT(2, 2) As Char  'Store player moves
        Static player As Char = "X"  'X goes first
        'Check for existing X or O
        If btnSquareClicked.Text <> Nothing Then
            MessageBox.Show("Invalid Move")
        Else
            'show move
            btnSquareClicked.Text = player
            'store move in chrTTT
            Dim index As String
            index = btnSquareClicked.Tag
            Dim x As Integer = Val(index.Chars(0))
            Dim y As Integer = Val(index.Chars(2))
            Call storemove(x, y, player, chrTTT)
            'check for winner
            If IsWinner(chrTTT) Then
                MessageBox.Show("Game Over!")
            Else
                If player = "X" Then
                    player = "O"
                Else
                    player = ("X")
                End If
            End If
        End If
    End Sub

    Sub StoreMove(ByVal x As Integer, ByVal y As Integer, ByVal player As Char, ByRef TTT(,) As Char)
        TTT(x, y) = player
    End Sub

    Function IsWinner(ByRef TTT(,) As Char) As Boolean
        For row As Integer = 0 To 2
            If TTT(row, 0) = TTT(row, 1) And TTT(row, 1) = TTT(row, 2) And (TTT(row, 0) = "X" Or TTT(row, 0) = "O") Then
                Return True
            End If
        Next row
        For col As Integer = 0 To 2
            If TTT(0, col) = TTT(1, col) And TTT(1, col) = TTT(2, col) And (TTT(0, col) = "X" Or TTT(0, col) = "O") Then
                Return True
            End If
        Next col
        If TTT(0, 0) = TTT(1, 1) And TTT(1, 1) = TTT(2, 2) And (TTT(0, 0) = "X" Or TTT(0, 0) = "O") Then
            Return True
        End If
        If TTT(0, 2) = TTT(1, 1) And TTT(1, 1) = TTT(2, 0) And (TTT(0, 2) = "X" Or TTT(0, 2) = "O") Then
            Return True
        End If
        Dim moveLeft As Boolean = False
        For row As Integer = 0 To 2
            For col As Integer = 0 To 2
                If TTT(row, col) = Nothing Then
                    moveLeft = True
                End If
            Next col
        Next row
        If Not moveLeft Then
            Return True
        End If
        Return False
    End Function
End Class
