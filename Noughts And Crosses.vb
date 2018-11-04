Public Class FrmNoughtsAndCrosses
    Private ReadOnly _playGame(3, 3) As Integer    'Creates the array for the game board to be stored in
    Private _playerGo, _counter, _playerNumber, _computerDifficulty As Integer
    'Defines these variables as type "Integer"

    Private Class MiniMaxReturn    'Creates a class called "MiniMaxReturn"
        Public X As Integer        'Sets "X", "Y" and "Score" as integer properties of the class
        Public Y As Integer
        Public Score As Integer
    End Class

    Private Sub FrmNoughtsAndCrosses_Load() Handles MyBase.Load 'Procedure executed when form loaded
        FormLoad()    'Executes procedure "FormLoad"
    End Sub

    Private Sub FormLoad()            'Procedure "FormLoad"
        Dim modeOutput, themeOutput As String    'Defines these as strings
        Dim btn As Control            'Defines "btn" as a control
        modeOutput = "Normal"         'Sets "modeOutput" to "Normal"
        themeOutput = "Light"         'Sets "themeOutput" to "Light"
        For Each btn In Controls            'Loops through all the control elements on the form
            If TypeOf btn Is Button Then    'If the element is a button then it is disabled
                btn.Enabled = False
            End If
        Next
        btnStart.Enabled = True       'Enables the button "btnStart"
        trbSelectDifficulty.Hide()    'Hides the elements used in one player mode
        lblEasy.Hide()
        lblHard.Hide()
        btnTwoPlayer.Checked = True   'Checks the radio button "btnTwoPlayer"
        _playerNumber = 2             'Sets "_playerNumber" to "2"
        _computerDifficulty = 1       'Sets "_computerDifficulty" to "1"
        cmbModeSelect.Text = modeOutput    'Sets the text in "cmbModeSelect" to "modeOutput"
        cmbThemeSelect.Text = themeOutput  'Sets the text in "cmbThemeSelect" to "themeOutput"
    End Sub

    Private Sub btnOnePlayer_Checked(sender As Object, e As EventArgs) Handles btnOnePlayer.CheckedChanged
        'Procedure executed when the radio button "btnOnePlayer" is checked
        Dim change As String          'Defines "change" as a "string"
        change = "Computer Colour"    'Sets "changes" to "Computer Colour"
        lblPlayer2Name.Hide()         'Hides the elements used in two player mode
        tbxPlayer2.Hide()
        lblColour2.Text = change      'Sets the value of "lblColour2.Text" to "change"
        lblEasy.Show()                'Shows the element used in one player mode
        lblHard.Show()
        trbSelectDifficulty.Show()
        _playerNumber = 1             'Sets "_playerNumber" to "1"
    End Sub

    Private Sub btnTwoPlayer_Checked(sender As Object, e As EventArgs) Handles btnTwoPlayer.CheckedChanged
        'Procedure executed when the radio button "btnTwoPlayer" is checked
        Dim change As String        'Defines "change" as a "string"
        change = "Colour"           'Sets "changes" to "Computer Colour"
        lblEasy.Hide()              'Hides the elements used in one player mode
        lblHard.Hide()
        trbSelectDifficulty.Hide()
        lblColour2.Text = change    'Sets the value of "lblColour2.Text" to "change"
        lblPlayer2Name.Show()       'Shows the element used in two player mode
        tbxPlayer2.Show()
        _playerNumber = 2           'Sets "_playerNumber" to "2"
    End Sub

    Private Sub trbSelectDifficulty_Change(sender As Object, e As EventArgs) Handles trbSelectDifficulty.Scroll
        'Procedure executed when the difficulty slider is changed
        If trbSelectDifficulty.Value = 0 Then 'If set to position "0" then "_computerDifficulty" set to "1"
            _computerDifficulty = 1
        ElseIf trbSelectDifficulty.Value = 1 Then 'Else if set to position "1" then "_computerDifficulty" set to "2"
            _computerDifficulty = 2
        Else 'Else runs procedure "ErrorHandler"                  
            ErrorHandler()
        End If
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        'Procedure executed when the start button is clicked
        StartRun()    'Executes procedure "StartRun"
    End Sub

    Private Sub StartRun() 'Executes procedure "StartRun"

        If _playerNumber = 2 Then 'If "_playerNumber" is 2 then:
            If tbxPlayer1.Text = "" Or tbxPlayer2.Text = "" Then
                MsgBox("Please enter a name each", MsgBoxStyle.Exclamation, "Choose A Name")
                Exit Sub
                'If either name boxes are blank then generates a message box and exits the procedure
            ElseIf tbxPlayer1.Text = tbxPlayer2.Text Then
                MsgBox("You have both chosen the same name. Please choose another", MsgBoxStyle.Exclamation,
                       "Choose A Name")
                Exit Sub
                'Else if either colour boxes are blank then generates a message box and exits the procedure
            End If
        ElseIf _playerNumber = 1 Then 'Else if "_playerNumber" is 1 then:
            If tbxPlayer1.Text = "" Then
                MsgBox("Please enter a name", MsgBoxStyle.Exclamation, "Choose A Name")
                Exit Sub
                'If the name box is blank then generates a message box and exits the procedure
            End If
        End If

        If cmbColoursPlayer1.Text = "" Or cmbColourPlayer2.Text = "" Then
            MsgBox("Please choose a colour each", MsgBoxStyle.Exclamation, "Choose A Colour")
            Exit Sub
            'If either name boxes have the same contents then generates a message box and exits the procedure
        ElseIf cmbColoursPlayer1.Text = cmbColourPlayer2.Text Then
            MsgBox("You have both chosen the same colour. Please choose another", MsgBoxStyle.Exclamation,
                   "Choose A Colour")
            Exit Sub
            'Else if either colour boxes have the same contents then generates a message box and exits the procedure
        End If

       Dim starter As String                          'Defines "starter" as type "string"
        Dim ctrl As Control                           'Defines "ctrl" as type "control"
        GamePlay_Clear()                              'Executes the procedure "GamePlay_Clear"
        For Each ctrl In Controls                     'Loops through all the control elements on the form
            If TypeOf ctrl Is Button Then             'If the element is a button then it is enabled
                ctrl.Enabled = True
            ElseIf TypeOf ctrl Is TextBox Then        'Else if the element is a text box it is disabled
                ctrl.Enabled = False
            ElseIf TypeOf ctrl Is ComboBox Then       'Else if the element is a combo box then it is disabled
                ctrl.Enabled = False
            ElseIf TypeOf ctrl Is RadioButton Then    'Else if the element is a radio button then it is disabled
                ctrl.Enabled = False
            ElseIf TypeOf ctrl Is TrackBar Then       'Else if the element is a track bar then it is disabled
                ctrl.Enabled = False
            End If
        Next
        btnStart.Enabled = False               'Disables the button "btnStart"
        starter = ChoseStarter()               'Sets the value of "starter" to the result of the function "ChoseStater"
        MsgBox("The starter is: " & starter, MsgBoxStyle.Information, "Starter")
        'Prints a message box telling who the starter is
        _playerGo = StartPlayer(starter)
        'Sets the value of "_playerGo" to the result of function "StartPlayer", passing in "player"
        _counter = 0                           'Sets the value of "_counter" to 0

        If _playerNumber = 1 And _playerGo = 10 Then
            'If there is one player and player two starts then the "ComputerGo" procedure is run
            ComputerGo()
        End If
    End Sub

    Private Sub GamePlay_Clear() 'Procedure "GamePlay_Clear"
        For i = 0 To 2 'Loops from 0 to 2 twice
            For a = 0 To 2
                _playGame(i, a) = 0     'Sets values of "_playGame" to "0"
            Next
        Next
    End Sub

    Private Function ChoseStarter()                    'Function "ChoseStarter"
        Dim answer As String                           'Defines "answer" as a "string"
        Dim randomSelected As Integer                  'Defines "randomSelected" as an "integer"
        Dim randomNumber = New Random                  'Defines "randomNumber" as a new "random"
        answer = ""                                    'Sets the value of "answer" to nothing

        randomSelected = randomNumber.Next(1, 50)
        'Generates a random number between 1 and 50 and stores it in the variable "randomNumber"

        Select Case (randomSelected Mod 2)
            'Matches the result of the remainder of "randomSelected" divided by 2 against values
            Case 0 'If the result is 0, then "answer" is the text in "tbxPlayer1"
                answer = tbxPlayer1.Text
            Case 1 'If the result is 1, then this selection statement is executed
                If _playerNumber = 1 Then 'If "_playerNumber" is 1 then "answer" is "The Computer"
                    answer = "The Computer"
                ElseIf _playerNumber = 2 Then 'Else if "_playerNumber" is 2 then "answer" is the text in "tbxPlayer2"
                    answer = tbxPlayer2.Text
                Else 'Else runs procedure "ErrorHandler"
                    ErrorHandler()
                End If
        End Select

        Return answer                          'Returns the string "answer"
    End Function

    Private Function StartPlayer(player) 'Function "StartPlayer" with "player" passed in
        Dim answer As Integer                  'Defines "answer" as "Integer"

        Select Case player 'Matches the variable "player" against certain values
            Case tbxPlayer1.Text 'If "player" is the text in "tbxPlayer1" then "answer" is 1
                answer = 1
            Case tbxPlayer2.Text 'If "player" is the text in "txbPlayer2" then "answer" is 10 
                answer = 10
            Case "The Computer" 'If "player" is "The Computer" then "answer" is 10
                answer = 10
        End Select

        Return answer                       'Returns "answer"
    End Function

    Private Sub GamePlay(index1, index2)            'Procedure "GamePlay", with "index1" and "index2" passed in
        Dim userChoice, winner As String            'Defines these as "strings"
        Dim win As Boolean                          'Defines "win" as "boolean"
        Dim playerChange As Integer                    'Defines "playerChange" as "Integer"

        If _playerGo = 1 Then                       'If "playerGo" is 1 then:
            _playGame(index1, index2) = _playerGo   'The relevant array value is set to 1
            win = CheckNought()
            'Function "CheckNought" is executed, passing in "_playGame" and its results stored in "win"
            GraphicsDraw(index1, index2)            'Procedure "GraphicsDraw" is run, passing in "index1" and "index2"
            playerChange = 10                       'The value of "playerChange" is set to 10
        ElseIf _playerGo = 10 Then                  'Else if "_playerGo" is 10 then:
            _playGame(index1, index2) = _playerGo   'The relevant array value is set to 10
            playerChange = 1                        'The value of "_playerChange" is set to 1
            win = CheckCross()
            'Function "CheckCross" is executed, passing in "_playGame" and its results stored in "win"
            GraphicsDraw(index1, index2)
            'Procedure "GraphicsDraw" executed with "index1" and "index2" passed in
        Else             'Else executes procedure "ErrorHandler"
            ErrorHandler()
        End If

        If win = True Then 'If "win" is true then execute the function "FindWinner" and store the result in "winner"
            winner = FindWinner()
            userChoice = MsgBox("The winner is: " & winner & vbCr & "Would you like to play again?",
                                MsgBoxStyle.YesNo, "End Of Game")
            'Generates a message box with the winner in it and stores their answer in "userChoice"
            RefreshForm(userChoice)    'Executes the procedure "RefreshForm", passing in "userChoice"
            Exit Sub                   'Exits the procedure
        End If

        If playerChange = 10 And _playerNumber = 1 Then    'If "playerChange" is 10 and "_playerNumber" is 1 then:
            _playerGo = 10                                 '"_playerGo" is set to 10
            ComputerGo()                                   'The procedure "ComputerGo" is executed
        ElseIf playerChange = 1 Then                       'Else if "playerChange" is 1 then "_playerGo" is set to 1
            _playerGo = 1
        ElseIf playerChange = 10 Then                      'Else if "_playerChange" is 10 then "_playerGo" is set to 10
            _playerGo = 10
        Else                                               'Else executes procedure "ErrorHandler"
            ErrorHandler()
        End If

        _counter += 1        'Adds 1 to the counter 
        If _counter = 9 Then 'If the counter is greater than 9 then generates a message box telling them its a draw
            userChoice = MsgBox("That game was a draw. Would you like to play again?", MsgBoxStyle.YesNo,
                                "End Of Game")
            'Stores the player's option in "userChoice"
            RefreshForm(userChoice)    'Executes the procedure "RefreshForm", passing in "userChoice"
            Exit Sub                   'Exits the procedure
        End If
    End Sub

    Private Sub ComputerGo()                  'Procedure "ComputerGo"
        Dim index1, index2 As Integer         'Defines these as "Integer"
        Dim index = New MiniMaxReturn()       'Defines "index" as a "MiniMaxReturn"
        If _computerDifficulty = 1 Then       'If "_computerDifficulty" is 1 then:
            index = EasyComputer()            'Sets the value of "index" to the result from the function "EasyComputer"
        ElseIf _computerDifficulty = 2 Then   'Else if "_computerDifficulty" is 2 then:
            index = HardComputer()            'Sets the value of "index" to the result from the function "HardComputer"
        Else                                  'Else executes the procedure "ErrorHandler"
            ErrorHandler()
        End If

        index1 = index.X                      'Sets the value of "index1" to the "X" property of "index"
        index2 = index.Y                      'Sets the value of "index2" to the "Y" property of "index"
        GamePlay(index1, index2)              'Executes the procedure "GamePlay", passing in "index1" and "index2"
    End Sub

    Private Function FindWinner()                            'Function "FindWinner"
        Dim answer As String                                 'Defines "answer" as a string
        answer = ""                                          'Sets the contents of "answer" to ""
        If _playerGo = 1 Then                                'If "_playerGo" is 1 then:
            answer = tbxPlayer1.Text                         'Contents of "answer" is the text in "tbxPlayer1"
        ElseIf _playerGo = 10 And _playerNumber = 1 Then     'Else if "_playerGo" is 10 and "_playerNumber" is 1 then:
            answer = "The Computer"                          'Contents of "answer" is "The Computer"
        ElseIf _playerGo = 10 Then                           'Else if "_playerGo" is 10 then:
            answer = tbxPlayer2.Text                         'Contents of "answer" is the text in "tbxPlayerTwo"
        Else                                                 'Else executes the procedure "ErrorHandler"
            ErrorHandler()
        End If

        Return answer                                        'Returns "answer"
    End Function

    Private Function CheckNought()                      'Function "CheckNought"
        Dim result As Integer                           'Defines "result" as Integer
        result = 0                                      'Sets the value of "result" to 0

        For x = 0 To 2     'Loops through the columns in the array and adds the values to "result"
            For y = 0 To 2
                result += _playGame(x, y)
            Next
            If result = 3 Then     'If "result" is 3 then returns true
                Return True
            End If
            result = 0             'Sets "result" to 0
        Next

        For y = 0 To 2     'Loops through the rows in the array and adds the values to "result"
            For x = 0 To 2
                result += _playGame(x, y)
            Next
            If result = 3 Then     'If "result" is 3 then returns true
                Return True
            End If
            result = 0             'Sets "result" to 0
        Next

        For index = 0 To 2     'Loops through one diagonal in the array and adds the values to "result"
            result += _playGame(index, index)
        Next
        If result = 3 Then     'If "result" is 3 then returns true
            Return True
        End If
        result = 0             'Sets "result" to 0

        For index = 0 To 2     'Loops through the other diagonal in the array and adds the values to "result"
            result += _playGame(index, index)
        Next
        If result = 3 Then     'If "result" is 3 then returns true
            Return True
        End If

        Return False           'Returns false
    End Function

    Private Function CheckCross()                        'Function "CheckCross"
        Dim result As Integer                            'Defines "result" as Integer
        result = 0                                       'Sets the value of "result" to 0

        For x = 0 To 2     'Loops through the columns in the array and adds the values to "result"
            For y = 0 To 2
                result += _playGame(x, y)
            Next
            If result = 30 Then     'If "result" is 30 then returns true
                Return True
            End If
            result = 0              'Sets "result" to 0
        Next

        For y = 0 To 2     'Loops through the rows in the array and adds the values to "result"
            For x = 0 To 2
                result += _playGame(x, y)
            Next
            If result = 30 Then     'If "result" is 30 then returns true
                Return True
            End If
            result = 0              'Sets "result" to 0
        Next

        For index = 0 To 2     'Loops through one diagonal in the array and adds the values to "result"
            result += _playGame(index, index)
        Next
        If result = 30 Then     'If "result" is 30 then returns true
            Return True
        End If
        result = 0              'Sets "result" to 0

        For index = 0 To 2     'Loops through the other diagonal in the array and adds the values to "result"
            result += _playGame(index, index)
        Next
        If result = 30 Then     'If "result" is 30 then returns true
            Return True
        End If

        Return False            'Returns false
    End Function

    Private Function DrawGraphics()                'Function "DrawGraphics"
        Dim imageReturn As New Bitmap(120, 120)    'Defines "imageReturn" as a bitmap 120px by 120px
        Dim colour As String                       'Defines "colour" as string
        If _playerGo = 1 Then                      'If "_playerGo" is 1 then:
            colour = cmbColoursPlayer1.Text        '"colour" is the text in "cmbColoursPlayer1"
            imageReturn = DrawNought(colour)
            '"imageReturn" is set to the result of the function "DrawNought", passing in "colour"
        ElseIf _playerGo = 10 Then 'Else if "_playerGo" is 10 then:
            colour = cmbColourPlayer2.Text         '"colour" is the text in "cmbColourPlayer2"
            imageReturn = DrawCross(colour)
            '"imageReturn" is set to the result of the function "DrawNought", passing in "colour""
        Else                                       'Else executes the procedure "ErrorHandler"
            ErrorHandler()
        End If

        Return imageReturn                         'Returns "imageReturn"
    End Function

    Private Shared Function DrawCross(colour) As Bitmap
        'Function "DrawCross" with "colour" passed in and result returned as a bitmap
        Dim crossPermanent As New Bitmap(120, 120)              'Defines "crossPermanent" as a bitmap 120px by 120px
        Dim colourUser As Color                                 'Defines "colourUser" as a colour
        Dim cross As Graphics                                   'Defines "cross" as graphics
        cross = Graphics.FromImage(crossPermanent)
        'Defines "cross" as a drawing surface for the bitmap "crossPermanent"
        colourUser = Color.FromName(colour)                     'Sets "colourUser" to the colour named in "colour"
        Dim colourPen As Pen                                    'Defines "colourPen" as a pen
        colourPen = New Pen(colourUser, 4)
        'Sets "colourPen" as a pen with the colour "colourUser" with a width of 4px
        Dim penPoints1, penPoints2 As Point()                   'Defines these as lists of points
        penPoints1 = {New Point(10, 10), New Point(110, 110)}
        'Sets the value of "penPoints1" to the points (10, 10) and (110, 110)
        penPoints2 = {New Point(10, 110), New Point(110, 10)}
        'Sets the value of "penPoints2" to the points (10, 110) and (110, 10)

        cross.DrawLines(colourPen, penPoints1)        'Draws a line using "colourPen" from the points in "penPoints1"
        cross.DrawLines(colourPen, penPoints2)        'Draws a line using "colourPen" from the points in "penPoints2"
        Return crossPermanent                         'Returns "crossPermanent"
    End Function

    Private Shared Function DrawNought(colour) As Bitmap
        'Function "DrawNought" with "colour" passed in and result returned as a bitmap
        Dim noughtPermanent As New Bitmap(120, 120)              'Defines "noughtPermanent" as a bitmap 120px by 120px
        Dim colourUser As Color                                  'Defines "colourUser" as a colour
        Dim nought As Graphics                                   'Defines "nought" as graphics
        nought = Graphics.FromImage(noughtPermanent)
        'Defines "cross" as a drawing surface for the bitmap "crossPermanent"
        colourUser = Color.FromName(colour)                      'Sets "colourUser" to the colour named in "colour"
        Dim colourPen As Pen                                     'Defines "colourPen" as a pen
        colourPen = New Pen(colourUser, 4)
        'Sets "colourPen" as a pen with the colour "colourUser" with a width of 4px

        nought.DrawEllipse(colourPen, 10, 10, 100, 100)
        'Draws a circle using "colourPen" from the point (10, 10) with width 100px and height 100 px
        Return noughtPermanent                                  'Returns "noughtPermanent"
    End Function

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        GamePlay(0, 0)    'Runs procedure "GamePlay" passing in 0 and 0 when "btn1" is clicked
    End Sub

    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
        GamePlay(0, 1)    'Runs procedure "GamePlay" passing in 0 and 1 when "btn2" is clicked
    End Sub

    Private Sub btn3_Click(sender As Object, e As EventArgs) Handles btn3.Click
        GamePlay(0, 2)    'Runs procedure "GamePlay" passing in 0 and 2 when "btn3" is clicked
    End Sub

    Private Sub btn4_Click(sender As Object, e As EventArgs) Handles btn4.Click
        GamePlay(1, 0)    'Runs procedure "GamePlay" passing in 1 and 0 when "btn4" is clicked
    End Sub

    Private Sub btn5_Click(sender As Object, e As EventArgs) Handles btn5.Click
        GamePlay(1, 1)    'Runs procedure "GamePlay" passing in 1 and 1 when "btn5" is clicked
    End Sub

    Private Sub btn6_Click(sender As Object, e As EventArgs) Handles btn6.Click
        GamePlay(1, 2)    'Runs procedure "GamePlay" passing in 1 and 2 when "btn6" is clicked
    End Sub

    Private Sub btn7_Click(sender As Object, e As EventArgs) Handles btn7.Click
        GamePlay(2, 0)    'Runs procedure "GamePlay" passing in 2 and 0 when "btn7" is clicked
    End Sub

    Private Sub btn8_Click(sender As Object, e As EventArgs) Handles btn8.Click
        GamePlay(2, 1)    'Runs procedure "GamePlay" passing in 2 and 1 when "btn8" is clicked
    End Sub

    Private Sub btn9_Click(sender As Object, e As EventArgs) Handles btn9.Click
        GamePlay(2, 2)    'Runs procedure "GamePlay" passing in 2 and 2 when "btn9" is clicked
    End Sub

    Private Sub ErrorHandler() 'Procedure "errorHandler"
        Dim userChoice As String      'Defines "userChoice" as string
        userChoice = MsgBox("An error has occured: " & Err.Description & "Would you like to try again?",
                            MsgBoxStyle.YesNo, "Error")
        'Sets "userChoice" to the option the user inputs
        RefreshForm(userChoice)       'Executes the procedure "RefreshForm", passing in "userChoice"
    End Sub

    Private Sub RefreshForm(userChoice)             'Procedure "RefreshForm", with "userChoice" being passed in
        Dim sameUser As String                      'Defines "sameUser" as a string
        If userChoice = DialogResult.Yes Then       'If the user clicked yes then:
            sameUser = MsgBox("Would you like to play with the same players?", MsgBoxStyle.YesNo,
                              "Play Again")
            'Sets "sameUser" to the option the user chooses
            If sameUser = DialogResult.Yes Then     'If the user chooses yes then execute the procedure "KeepUser"
                KeepUser()
            ElseIf sameUser = DialogResult.No Then  'Else if the user chooses no refreshes the form
                Controls.Clear()
                InitializeComponent()
                Show()
                FormLoad()
            Else                                    'Else execute the procedure "ErrorHandler"
                ErrorHandler()
            End If
        ElseIf userChoice = DialogResult.No Then    'Else if the user chooses no to the original box then close the form
            Close()
        Else                                        'Else execute the procedure "ErrorHandler"
            ErrorHandler()
        End If
    End Sub

    Private Sub KeepUser()         'Procedure "KeepUser"
        Dim playerOne, playerTwo, playerOneColour, playerTwoColour As String    'Defines these variables as "string"
        Dim modeSelected, difficultySelected As Integer    'Defines these variables as "Integer"
        playerOne = tbxPlayer1.Text                        'Sets the value of "playerOne" to the text in "tbxPlayer1"
        playerTwo = tbxPlayer2.Text                        'Sets the value of "playerTwo" to the text in "tbxPlayer2"
        playerOneColour = cmbColoursPlayer1.Text
        'Sets the value of "playerOneColour" to the text in "cmbColoursPlayer1"
        playerTwoColour = cmbColourPlayer2.Text
        'Sets the value of "playerTwoColour" to the text in "cmbColourPlayer2"

        If btnOnePlayer.Checked = True Then          'If "btnOnePlayer" is checked then "modeSelected" = 1
            modeSelected = 1
        ElseIf btnTwoPlayer.Checked = True Then      'Else if "btnTwoPlayer" is checked then "modeSelected" = 2
            modeSelected = 2
        Else                                         'Else execute procedure "ErrorHandler"
            ErrorHandler()
        End If

        If trbSelectDifficulty.Value = 0 Then        'If the value of "trbSelectDifficulty" is 0 then:
            difficultySelected = 1                   '"difficultySelected" = 1
        ElseIf trbSelectDifficulty.Value = 1 Then    'Else if the value of "trbSelectDifficulty" is 1 then:
            difficultySelected = 2                   '"difficultySelected" = 2
        Else                                         'Else execute the procedure "ErrorHandler"
            ErrorHandler()
        End If

        Controls.Clear()                             'Refreshes the form
        InitializeComponent()
        Show()
        FormLoad()

        tbxPlayer1.Text = playerOne                'Sets the text in "tbxPlayer1" to "playerOne"
        tbxPlayer2.Text = playerTwo                'Sets the text in "tbxPlayer2" to "playerTwo"
        cmbColoursPlayer1.Text = playerOneColour   'Sets the text in "cmbColoursPlayer1" to "playerOneColour"
        cmbColourPlayer2.Text = playerTwoColour    'Sets the text in "cmbColourPlayer2" to "playerTwoColour"
        If modeSelected = 1 Then                   'If "modeSelected" is 1 then "btnOnePlayer" is checked
            btnOnePlayer.Checked = True
            If difficultySelected = 1 Then         'If "difficultySelected" is 1 then:
                trbSelectDifficulty.Value = 0      'Sets the track bar "trbSelectDifficulty" to position 0
                _computerDifficulty = 1            'Sets "_computerDifficulty" to 1
            ElseIf difficultySelected = 2 Then     'Else if "difficultySelected" is 2 then:
                trbSelectDifficulty.Value = 1      'Sets the track bar "trbSelectDifficulty" to position 1
                _computerDifficulty = 2            'Sets "_computerDifficulty" to 2
            Else                                   'Else executes the procedure "ErrorHandler"
                ErrorHandler()
            End If
        ElseIf modeSelected = 2 Then               'Else if "modeSelected" is 2 then "btnTwoPlayer" is checked
            btnTwoPlayer.Checked = True
        Else                                       'Else executes the procedure "ErrorHandler"
            ErrorHandler()
        End If
        StartRun()                                 'Executes the procedure "StartRun"
    End Sub

    Private Function EasyComputer()                'Function "EasyComputer"
        Dim randomNumber As New Random             'Defines "randomNumber" as a "random"
        Dim newBoard(2, 2) As Integer              'Defines "newBoard" as a 2 by 2 array of "integers"
        Dim x, y As Integer                        'Defines these as "integer"
        Dim freeSpaces As New List(Of Integer)     'Defines "freeSpaces" as a list of "integer"
        Dim result = New MiniMaxReturn()           'Defines "result" as a member of the class "MiniMaxReturn"
        newBoard = _playGame                       'Sets the array "newBoard" to the main board "_playGame"
        freeSpaces = FindFreeSpaces(newBoard)
        'Sets the value of "freeSpaces" to the result of the function "FindFreeSpaces", passing in "newBoard"

        x = randomNumber.Next(0, (freeSpaces.Count \ 2))
        'Sets "x" to a random number between 0 and the length of "freeSpaces" integer divided by 2
        x *= 2        'Multiplies "x" by 2
        y = x + 1     'Sets "y" to the value of "x" + 1
        result.X = freeSpaces(x)    'Sets the "X" property of "result" to the value in "freeSpaces" at index "x"
        result.Y = freeSpaces(y)    'Sets the "Y" property of "result" to the value in "freeSpaces" at index "y"

        Return result        'Returns "result"
    End Function

    Private Shared Function FindFreeSpaces(board)        'Function "FindFreeSpaces" with "board" passed in
        Dim result As New List(Of Integer)

        For a = 0 To 2
            For b = 0 To 2
                If board(a, b) = 0 Then
                    result.Add(a)
                    result.Add(b)
                End If
            Next
        Next

        Return result

    End Function

    Private Function HardComputer()
        Dim answer = New MiniMaxReturn
        Dim newBoard(2, 2) As Integer
        Dim counter, playerGo As Integer
        counter = 0
        playerGo = 10
        answer.Score = 0
        newBoard = _playGame
        If _counter = 8 Then
            answer = EasyComputer()
        Else
            answer = MiniMax(newBoard, answer, playerGo, counter)
        End If
        Return answer

    End Function

    Private Function MiniMax(newBoard, result, player, counter)
        Dim emptySpaces As New List(Of Integer)
        Dim answers As New List(Of MiniMaxReturn)
        Dim answer = New MiniMaxReturn()
        Dim answerReturn = New MiniMaxReturn()
        Dim bestAnswer = New MiniMaxReturn()
        Dim bestScore, bestMove, x, y As Integer

        If CheckNought() = True Then
            result.Score = -10
            Return result
        ElseIf CheckCross() = True Then
            result.Score = 10
            Return result
        ElseIf counter = 9 Then
            result.Score = 0
            Return result
        End If

        emptySpaces = FindFreeSpaces(newBoard)

        For i = 0 To (emptySpaces.Count - 1) Step 2
            x = emptySpaces(i)
            y = emptySpaces(i + 1)
            newBoard(x, y) = player
            answer.X = x
            answer.Y = y
            result = MiniMax(newBoard, result, 10, (counter + 1))
            answer = result
            counter = 0
            newBoard(emptySpaces(i), emptySpaces(i + 1)) = 0
            answers.Add(answer)
        Next


        bestScore = -10000
        For i = 0 To (answers.Count - 1)
            bestAnswer = answers(i)
            If bestAnswer.Score < bestScore Then
                bestMove = i
            End If
        Next

        If answers.Count <> 0 Then
            answerReturn = answers(bestMove)
        Else
            answerReturn = EasyComputer()
        End If

        Return answerReturn

    End Function

    Private Sub GraphicsDraw(x, y) 'Procedure "GraphicsDraw" with "x" and "y" passed in
        Dim image As Bitmap                   'Defines "image" as a "bitmap"
        image = DrawGraphics()

        If x = 0 And y = 0 Then 'If "x" is 0 and "y" is 0 then "btn1" is disabled and background set to "image"
            btn1.Enabled = False
            btn1.BackgroundImage = image
        ElseIf x = 0 And y = 1 Then 'If "x" is 0 and "y" is 1 then "btn2" is disabled and background set to "image"
            btn2.Enabled = False
            btn2.BackgroundImage = image
        ElseIf x = 0 And y = 2 Then 'If "x" is 0 and "y" is 2 then "btn3" is disabled and background set to "image"
            btn3.Enabled = False
            btn3.BackgroundImage = image
        ElseIf x = 1 And y = 0 Then 'If "x" is 1 and "y" is 0 then "btn4" is disabled and background set to "image"
            btn4.Enabled = False
            btn4.BackgroundImage = image
        ElseIf x = 1 And y = 1 Then 'If "x" is 1 and "y" is 1 then "btn5" is disabled and background set to "image"
            btn5.Enabled = False
            btn5.BackgroundImage = image
        ElseIf x = 1 And y = 2 Then 'If "x" is 1 and "y" is 2 then "btn6" is disabled and background set to "image"
            btn6.Enabled = False
            btn6.BackgroundImage = image
        ElseIf x = 2 And y = 0 Then 'If "x" is 0 and "y" is 0 then "btn7" is disabled and background set to "image"
            btn7.Enabled = False
            btn7.BackgroundImage = image
        ElseIf x = 2 And y = 1 Then 'If "x" is 2 and "y" is 1 then "btn8" is disabled and background set to "image"
            btn8.Enabled = False
            btn8.BackgroundImage = image
        ElseIf x = 2 And y = 2 Then 'If "x" is 2 and "y" is 2 then "btn9" is disabled and background set to "image"
            btn9.Enabled = False
            btn9.BackgroundImage = image
        Else 'Else procedure "ErrorHandler" is executed
            ErrorHandler()
        End If
    End Sub

    Private Sub cmbModeSelect_Click(system As Object, e As EventArgs) Handles cmbModeSelect.SelectedIndexChanged
        Dim super As String
        super = "Super"
        If cmbModeSelect.Text = super Then
            FrmSuperNoughtsAndCrosses.Show()
            Hide()
        End If
    End Sub

    Private Sub cmbThemeSelect_Click(system As Object, e As EventArgs) Handles cmbThemeSelect.SelectedIndexChanged
        Dim light, dark, mac As String
        light = "Light"
        dark = "Goth"
        mac = "MacOS"

        If cmbThemeSelect.Text = light Then
            LightMode()
        ElseIf cmbThemeSelect.Text = dark Then
            DarkMode()
        Else If cmbThemeSelect.Text = mac Then
            MacChoice()
        Else
            ErrorHandler()
        End If

    End Sub

    Private Sub LightMode()
        Dim backColour As Color
        Dim foreColour As Color
        backColour = Color.White
        foreColour = Color.Black

        BackColor = backColour
        lblMode.ForeColor = foreColour
        lblTheme.ForeColor = foreColour
        btnOnePlayer.ForeColor = foreColour
        btnTwoPlayer.ForeColor = foreColour
        lblColour1.ForeColor = foreColour
        lblColour2.ForeColor = foreColour
        lblEasy.ForeColor = foreColour
        lblHard.ForeColor = foreColour
        btnStart.ForeColor = foreColour
        trbSelectDifficulty.ForeColor = foreColour
        trbSelectDifficulty.BackColor = backColour
        lblPlayer1Name.ForeColor = foreColour
        lblPlayer2Name.ForeColor = foreColour
        btnStart.ForeColor = foreColour
        btnStart.BackColor = backColour
    End Sub

    Private Sub DarkMode()
        Dim backColour As Color
        Dim foreColour As Color
        backColour = Color.Black
        foreColour = Color.White

        BackColor = backColour
        lblMode.ForeColor = foreColour
        lblTheme.ForeColor = foreColour
        btnOnePlayer.ForeColor = foreColour
        btnTwoPlayer.ForeColor = foreColour
        lblColour1.ForeColor = foreColour
        lblColour2.ForeColor = foreColour
        lblEasy.ForeColor = foreColour
        lblHard.ForeColor = foreColour
        btnStart.ForeColor = foreColour
        trbSelectDifficulty.ForeColor = foreColour
        trbSelectDifficulty.BackColor = backColour
        lblPlayer1Name.ForeColor = foreColour
        lblPlayer2Name.ForeColor = foreColour
        btnStart.ForeColor = foreColour
        btnStart.BackColor = backColour
    End Sub
    
    Private Sub MacChoice()
        Dim userChoice, output As String
        output = "Light"
        userChoice = MsgBox("Are you sure you want to do this?", MsgBoxStyle.YesNo, "Do You Want To Continue?")
        If userChoice = DialogResult.Yes Then
            MsgBox("Your choice of operating system has caused this program to suffer a fatal error", _
                    MsgBoxStyle.Critical, "Fatal Error")
            Close()
        ElseIf userChoice = DialogResult.No Then
            MsgBox("That was a close shave", MsgBoxStyle.Exclamation, "Close Shave")
            cmbThemeSelect.Text = output
        Else 
            ErrorHandler()
        End If
    End Sub

End Class