Public Class FrmSuperNoughtsAndCrosses
    Private ReadOnly _board1(2, 2), _board2(2, 2), _board3(2, 2), _board4(2, 2), _board5(2, 2), _
                     _board6(2, 2), _board7(2, 2), _board8(2, 2), _board9(2, 2), _board(2, 2) As Integer
    Private _playerGo As Integer
    Private _board1Win, _board2Win, _board3Win, _board4Win, _board5Win, _board6Win, _board7Win, _board8Win, _
        _board9Win As Boolean
    
    Private Sub FrmNoughtsAndCrosses_Load() Handles MyBase.Load     'Procedure executed when form loaded
        FormLoad()    'Executes procedure "FormLoad"
    End Sub
    
    Private Sub FormLoad()            'Procedure "FormLoad"
        Dim modeOutput, themeOutput As String    'Defines these as strings
        Dim btn As Control            'Defines "btn" as a control
        modeOutput = "Super"          'Sets "modeOutput" to "Normal"
        themeOutput = "Light"         'Sets "themeOutput" to "Light"
        For Each btn In Controls            'Loops through all the control elements on the form
            If TypeOf btn Is Button Then    'If the element is a button then it is disabled
                btn.Enabled = False
            End If
        Next
        btnBoard1.Hide()                   'Hides all the big buttons
        btnBoard2.Hide()
        btnBoard3.Hide()
        btnBoard4.Hide()
        btnBoard5.Hide()
        btnBoard6.Hide()
        btnBoard7.Hide()
        btnBoard8.Hide()
        btnBoard9.Hide()
        btnStart.Enabled = True            'Enables the button "btnStart"
        cmbModeSelect.Text = modeOutput    'Sets the text in "cmbModeSelect" to "modeOutput"
        cmbThemeSelect.Text = themeOutput  'Sets the text in "cmbThemeSelect" to "themeOutput"
        _board1Win = False
        _board2Win = False
        _board3Win = False
        _board4Win = False
        _board5Win = False
        _board6Win = False
        _board7Win = False
        _board8Win = False
        _board9Win = False
    End Sub
    
    Private Sub cmbModeSelect_Click(system As Object, e As EventArgs) Handles cmbModeSelect.SelectedIndexChanged
        Dim normal As String
        normal = "Normal"
        If cmbModeSelect.Text = normal Then
            FrmNoughtsAndCrosses.Show()
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
        lblColour1.ForeColor = foreColour
        lblColour2.ForeColor = foreColour
        btnStart.ForeColor = foreColour
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
        lblColour1.ForeColor = foreColour
        lblColour2.ForeColor = foreColour
        btnStart.ForeColor = foreColour
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
    
    Private Sub ErrorHandler() 'Procedure "errorHandler"
        Dim userChoice As String      'Defines "userChoice" as string
        userChoice = MsgBox("An error has occured: " & Err.Description & "Would you like to try again?",
                            MsgBoxStyle.YesNo, "Error")
        'Sets "userChoice" to the option the user inputs
        RefreshForm(userChoice)       'Executes the procedure "RefreshForm", passing in "userChoice"
    End Sub

    Private Sub RefreshForm(userChoice) 'Procedure "RefreshForm", with "userChoice" being passed in
        Dim sameUser As String             'Defines "sameUser" as a string
        If userChoice = DialogResult.Yes Then 'If the user clicked yes then:
            sameUser = MsgBox("Would you like to play with the same players?", MsgBoxStyle.YesNo,
                              "Play Again")
            'Sets "sameUser" to the option the user chooses
            If sameUser = DialogResult.Yes Then 'If the user chooses yes then execute the procedure "KeepUser"
                KeepUser()
            ElseIf sameUser = DialogResult.No Then 'Else if the user chooses no refreshes the form
                Controls.Clear()
                InitializeComponent()
                Show()
                FormLoad()
            Else 'Else execute the procedure "ErrorHandler"
                ErrorHandler()
            End If
        ElseIf userChoice = DialogResult.No Then 'Else if the user chooses no to the original box then close the form
            Close()
        Else 'Else execute the procedure "ErrorHandler"
            ErrorHandler()
        End If
    End Sub

    Private Sub KeepUser() 'Procedure "KeepUser"
        Dim playerOne, playerTwo, playerOneColour, playerTwoColour As String    'Defines these variables as "string"
        playerOne = tbxPlayer1.Text                'Sets the value of "playerOne" to the text in "tbxPlayer1"
        playerTwo = tbxPlayer2.Text                'Sets the value of "playerTwo" to the text in "tbxPlayer2"
        playerOneColour = cmbColoursPlayer1.Text
        'Sets the value of "playerOneColour" to the text in "cmbColoursPlayer1"
        playerTwoColour = cmbColourPlayer2.Text
        'Sets the value of "playerTwoColour" to the text in "cmbColourPlayer2"

        Controls.Clear()                             'Refreshes the form
        InitializeComponent()
        Show()
        FormLoad()

        tbxPlayer1.Text = playerOne                'Sets the text in "tbxPlayer1" to "playerOne"
        tbxPlayer2.Text = playerTwo                'Sets the text in "tbxPlayer2" to "playerTwo"
        cmbColoursPlayer1.Text = playerOneColour   'Sets the text in "cmbColoursPlayer1" to "playerOneColour"
        cmbColourPlayer2.Text = playerTwoColour    'Sets the text in "cmbColourPlayer2" to "playerTwoColour"
        
        StartRun()                                 'Executes the procedure "StartRun"
    End Sub
    
    Private Sub btnStart_Click(system As Object, e As EventArgs) Handles btnStart.Click
        StartRun()
    End Sub
    
    Private Sub StartRun()     'Executes procedure "StartRun"
        Dim penPoints As Point()
        Dim background As New Bitmap(656, 656)
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

        Dim starter As String                    'Defines "starter" as type "string"
        Dim ctrl As Control                      'Defines "ctrl" as type "control"
        btnStart.Enabled = False                 'Disables all the input controls
        tbxPlayer1.Enabled = False
        tbxPlayer2.Enabled = False
        cmbColoursPlayer1.Enabled = False
        cmbColourPlayer2.Enabled = False
        GamePlay_Clear()                         'Executes the procedure "GamePlay_Clear"
        For Each ctrl In Controls                'Loops through all the control elements on the form
            If TypeOf ctrl Is Button Then        'If the element is a button then it is enabled
                ctrl.Enabled = True
            ElseIf TypeOf ctrl Is TextBox Then   'Else if the element is a text box it is disabled
                ctrl.Enabled = False
            ElseIf TypeOf ctrl Is ComboBox Then  'Else if the element is a combo box then it is disabled
                ctrl.Enabled = False
            End If
        Next
        btnStart.Enabled = False                 'Disables the button "btnStart"
        starter = ChoseStarter()               'Sets the value of "starter" to the result of the function "ChoseStater"
        MsgBox("The starter is: " & starter, MsgBoxStyle.Information, "Starter")
        'Prints a message box telling who the starter is
        _playerGo = StartPlayer(starter)
        'Sets the value of "_playerGo" to the result of function "StartPlayer", passing in "player"
        penPoints = {New Point(10, 10), New Point(646, 10), New Point(646, 646), New Point(10, 646), _ 
                     New Point(10, 10)}
        'Sets the value of "penPoints" to the points (10, 10), (646, 10), (646, 646), (10, 646) and (10, 10)
        background = BackgroundDraw(background, penPoints)
        pnlBackground.BackgroundImage = background
    End Sub

    Private Shared Sub GamePlay_Clear() 'Procedure "GamePlay_Clear"
        For i = 0 To 2 'Loops from 0 to 2 twice
            For a = 0 To 2
               ' _playGame(i, a) = 0     'Sets values of "_playGame" to "0"
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
            Case 0         'If the result is 0, then "answer" is the text in "tbxPlayer1"
                answer = tbxPlayer1.Text
            Case 1         'If the result is 1, then "answer" is the text in "tbxPlayer2"
                    answer = tbxPlayer2.Text
        End Select

        Return answer                          'Returns the string "answer"
    End Function

    Private Function StartPlayer(player)       'Function "StartPlayer" with "player" passed in
        Dim answer As Integer                  'Defines "answer" as "Integer"

        Select Case player                     'Matches the variable "player" against certain values
            Case tbxPlayer1.Text               'If "player" is the text in "tbxPlayer1" then "answer" is 1
                answer = 1
            Case tbxPlayer2.Text               'If "player" is the text in "txbPlayer2" then "answer" is 10 
                answer = 10
        End Select

        Return answer                          'Returns "answer"
    End Function
    
    Private Function DrawGraphics(type, player)       'Function "DrawGraphics" with "type" and "player" passed in
        Dim imageReturnSmall As New Bitmap(60, 60)    'Defines "imageReturn" as a bitmap 60px by 60px
        Dim imageReturnBig As New Bitmap(192, 192)
        Dim colour As String                          'Defines "colour" as string
        Select Case player
            Case 1
                Select Case type
                    Case 1
                        colour = cmbColoursPlayer1.Text
                        imageReturnSmall = DrawNought(colour, type, imageReturnSmall)
                        Return imageReturnSmall
                    Case 2
                        colour = cmbColoursPlayer1.Text
                        imageReturnBig = DrawNought(colour, type, imageReturnBig)
                        Return imageReturnBig
                End Select
            Case 10
                Select Case type
                    Case 1
                        colour = cmbColourPlayer2.Text
                        imageReturnSmall = DrawCross(colour, type, imageReturnSmall)
                        Return imageReturnSmall
                    Case 2
                        colour = cmbColourPlayer2.Text
                        imageReturnBig = DrawCross(colour, type, imageReturnBig)
                        Return imageReturnBig
                End Select
            Case 0
                imageReturnBig = DrawNo(imageReturnBig)
                Return imageReturnBig
        End Select
        
        Return Nothing
        
    End Function
    
    Private Shared Function DrawNo(noPermanent) As Bitmap
        'Function "DrawCross" with "colour" passed in and result returned as a bitmap
        Dim colourUser As Color                                 'Defines "colourUser" as a colour
        Dim no As Graphics                                      'Defines "no" as graphics
        no = Graphics.FromImage(noPermanent)
        'Defines "no" as a drawing surface for the bitmap "crossPermanent"
        colourUser = Color.DimGray                              'Sets "colourUser" to dim gray
        Dim colourPen As Pen                                    'Defines "colourPen" as a pen
        colourPen = New Pen(colourUser, 10)
        'Sets "colourPen" as a pen with the colour "colourUser" with a width of 4px
        Dim penPoints1, penPoints2 As Point()                   'Defines these as lists of points
        penPoints1 = {New Point(40, 40), New Point(152, 152)}
        'Sets the value of "penPoints1" to the points (40, 40) and (152, 152)
        penPoints2 = {New Point(40, 152), New Point(152, 40)}
        'Sets the value of "penPoints1" to the points (40, 152) and (152, 40)
        
        no.DrawLines(colourPen, penPoints1)        'Draws a line using "colourPen" from the points in "penPoints1"
        no.DrawLines(colourPen, penPoints2)        'Draws a line using "colourPen" from the points in "penPoints2"
        no.DrawEllipse(colourPen, 20, 20, 152, 152)
        'Draws a circle using "colourPen" from the point (20, 20) with width 152px and height 152px
        Return noPermanent
    End Function

    Private Function DrawCross(colour, type, crossPermanent) As Bitmap
        'Function "DrawCross" with "colour" passed in and result returned as a bitmap
        Dim colourUser As Color                                 'Defines "colourUser" as a colour
        Dim cross As Graphics                                   'Defines "cross" as graphics
        cross = Graphics.FromImage(crossPermanent)
        'Defines "cross" as a drawing surface for the bitmap "crossPermanent"
        colourUser = Color.FromName(colour)                     'Sets "colourUser" to the colour named in "colour"
        Dim colourPen As Pen                                    'Defines "colourPen" as a pen
        colourPen = New Pen(colourUser, 4)
        'Sets "colourPen" as a pen with the colour "colourUser" with a width of 4px
        Dim penPoints1, penPoints2 As Point()                   'Defines these as lists of points
        If type = 1 Then
            penPoints1 = {New Point(10, 10), New Point(50, 50)}
            'Sets the value of "penPoints1" to the points (10, 10) and (50, 50)
            penPoints2 = {New Point(10, 50), New Point(50, 10)}
            'Sets the value of "penPoints2" to the points (10, 50) and (50, 10)
        Else If type = 2 Then
            penPoints1 = {New Point(10, 10), New Point(182, 182)}
            'Sets the value of "penPoints1" to the points (10, 10) and (182, 182)
            penPoints2 = {New Point(10, 182), New Point(182, 10)}
            'Sets the value of "penPoints2" to the points (10, 182) and (182, 10)
        Else 
            ErrorHandler()
        End If

        cross.DrawLines(colourPen, penPoints1)        'Draws a line using "colourPen" from the points in "penPoints1"
        cross.DrawLines(colourPen, penPoints2)        'Draws a line using "colourPen" from the points in "penPoints2"
        Return crossPermanent                         'Returns "crossPermanent"
    End Function

    Private Function DrawNought(colour, type, noughtPermanent) As Bitmap
        'Function "DrawNought" with "colour" passed in and result returned as a bitmap
        Dim colourUser As Color                                  'Defines "colourUser" as a colour
        Dim nought As Graphics                                   'Defines "nought" as graphics
        nought = Graphics.FromImage(noughtPermanent)
        'Defines "cross" as a drawing surface for the bitmap "crossPermanent"
        colourUser = Color.FromName(colour)                      'Sets "colourUser" to the colour named in "colour"
        Dim colourPen As Pen                                     'Defines "colourPen" as a pen
        colourPen = New Pen(colourUser, 4)
        'Sets "colourPen" as a pen with the colour "colourUser" with a width of 4px

        If type = 1 Then
            nought.DrawEllipse(colourPen, 10, 10, 40, 40)
            'Draws a circle using "colourPen" from the point (10, 10) with width 40px and height 40px
        ElseIf type = 2 Then
            nought.DrawEllipse(colourPen, 10, 10, 172, 172)
            'Draws a circle using "colourPen" from the point (10, 10) with width 172px and height 172px
        Else 
            ErrorHandler()
        End If
        Return noughtPermanent                                  'Returns "noughtPermanent"
    End Function
    
    Private Sub btnBoard1No1_Click(sender As Object, e As EventArgs) Handles btnBoard1No1.Click
        'Procedure run when "btnBoard1No1" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard1No1.BackgroundImage = image   'Sets the background of "btnBoard1No1" to image
        _board1(0, 0) = GamePlay(0, 0, _board1, 1)
        'Sets the value indexed at (0, 0) in _board1 to the result of the function "GamePlay", passing in 0, 0,
        '"_board1" and 1
        FindNextBoard(1)              'Runs subroutines "FindNextBoard", passing in 1 and "_board1"
    End Sub
    
    Private Sub btnBoard1No2_Click(sender As Object, e As EventArgs) Handles btnBoard1No2.Click
        'Procedure run when "btnBoard1No2" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard1No2.BackgroundImage = image   'Sets the background of "btnBoard1No2" to image
        _board1(0, 1) = GamePlay(0, 1, _board1, 1)
        'Sets the value indexed at (0, 1) in _board1 to the result of the function "GamePlay", passing in 0, 1,
        '"_board1" and 1
        FindNextBoard(2)              'Runs subroutines "FindNextBoard", passing in 2 and "_board1"
    End Sub
    
    Private Sub btnBoard1No3_Click(sender As Object, e As EventArgs) Handles btnBoard1No3.Click
        'Procedure run when "btnBoard1No3" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard1No3.BackgroundImage = image   'Sets the background of "btnBoard1No3" to image
        _board1(0, 2) = GamePlay(0, 2, _board1, 1)
        'Sets the value indexed at (0, 0) in _board1 to the result of the function "GamePlay", passing in 0, 2,
        '"_board1" and 1
        FindNextBoard(3)              'Runs subroutines "FindNextBoard", passing in 3 and "_board1"
    End Sub
    
    Private Sub btnBoard1No4_Click(sender As Object, e As EventArgs) Handles btnBoard1No4.Click
        'Procedure run when "btnBoard1No4" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard1No4.BackgroundImage = image   'Sets the background of "btnBoard1No4" to image
        _board1(1, 0) = GamePlay(1, 0, _board1, 1)
        'Sets the value indexed at (0, 0) in _board1 to the result of the function "GamePlay", passing in 1, 0,
        '"_board1" and 1
        FindNextBoard(4)              'Runs subroutines "FindNextBoard", passing in 4 and "_board1"
    End Sub
    
    Private Sub btnBoard1No5_Click(sender As Object, e As EventArgs) Handles btnBoard1No5.Click
        'Procedure run when "btnBoard1No5" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard1No5.BackgroundImage = image   'Sets the background of "btnBoard1No5" to image
        _board1(1, 1) = GamePlay(1, 1, _board1, 1)
        'Sets the value indexed at (0, 0) in _board1 to the result of the function "GamePlay", passing in 1, 1,
        '"_board1" and 1
        FindNextBoard(5)              'Runs subroutines "FindNextBoard", passing in 5 and "_board1"
    End Sub
    
    Private Sub btnBoard1No6_Click(sender As Object, e As EventArgs) Handles btnBoard1No6.Click
        'Procedure run when "btnBoard1No6" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard1No6.BackgroundImage = image   'Sets the background of "btnBoard1No6" to image
        _board1(1, 2) = GamePlay(1, 2, _board1, 1)
        'Sets the value indexed at (0, 0) in _board1 to the result of the function "GamePlay", passing in 1, 2,
        '"_board1" and 1
        FindNextBoard(6)              'Runs subroutines "FindNextBoard", passing in 6 and "_board1"
    End Sub

    Private Sub btnBoard1No7_Click(sender As Object, e As EventArgs) Handles btnBoard1No7.Click
        'Procedure run when "btnBoard1No7" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard1No7.BackgroundImage = image   'Sets the background of "btnBoard1No7" to image
        _board1(2, 0) = GamePlay(2, 0, _board1, 1)
        'Sets the value indexed at (0, 0) in _board1 to the result of the function "GamePlay", passing in 2, 0,
        '"_board1" and 1
        FindNextBoard(7)              'Runs subroutines "FindNextBoard", passing in 7 and "_board1"
    End Sub

    Private Sub btnBoard1No8_Click(sender As Object, e As EventArgs) Handles btnBoard1No8.Click
        'Procedure run when "btnBoard1No8" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard1No8.BackgroundImage = image   'Sets the background of "btnBoard1No8" to image
        _board1(2, 1) = GamePlay(2, 1, _board1, 1)
        'Sets the value indexed at (0, 0) in _board1 to the result of the function "GamePlay", passing in 2, 1,
        '"_board1" and 1
        FindNextBoard(8)              'Runs subroutines "FindNextBoard", passing in 8 and "_board1"
    End Sub
    
    Private Sub btnBoard1No9_Click(sender As Object, e As EventArgs) Handles btnBoard1No9.Click
        'Procedure run when "btnBoard1No9" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard1No9.BackgroundImage = image   'Sets the background of "btnBoard1No9" to image
        _board1(2, 2) = GamePlay(2, 2, _board1, 1)
        'Sets the value indexed at (2, 2) in _board1 to the result of the function "GamePlay", passing in 2, 2,
        '"_board1" and 1
        FindNextBoard(9)              'Runs subroutines "FindNextBoard", passing in 9 and "_board1"
    End Sub
    
    Private Sub btnBoard2No1_Click(sender As Object, e As EventArgs) Handles btnBoard2No1.Click
        'Procedure run when "btnBoard2No1" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard2No1.BackgroundImage = image   'Sets the background of "btnBoard2No1" to image
        _board2(0, 0) = GamePlay(0, 0, _board2, 2)
        'Sets the value indexed at (0, 0) in _board2 to the result of the function "GamePlay", passing in 0, 0,
        '"_board2" and 2
        FindNextBoard(1)              'Runs subroutines "FindNextBoard", passing in 1 and "_board1"
    End Sub
    
    Private Sub btnBoard2No2_Click(sender As Object, e As EventArgs) Handles btnBoard2No2.Click
        'Procedure run when "btnBoard2No2" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard2No2.BackgroundImage = image   'Sets the background of "btnBoard2No2" to image
        _board2(0, 1) = GamePlay(0, 1, _board2, 2)
        'Sets the value indexed at (0, 1) in _board2 to the result of the function "GamePlay", passing in 0, 1,
        '"_board2" and 2
        FindNextBoard(2)              'Runs subroutines "FindNextBoard", passing in 2 and "_board2"
    End Sub
    
    Private Sub btnBoard2No3_Click(sender As Object, e As EventArgs) Handles btnBoard2No3.Click
        'Procedure run when "btnBoard2No3" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard2No3.BackgroundImage = image   'Sets the background of "btnBoard2No3" to image
        _board2(0, 2) = GamePlay(0, 2, _board2, 2)
        'Sets the value indexed at (0, 0) in _board2 to the result of the function "GamePlay", passing in 0, 2,
        '"_board2" and 2
        FindNextBoard(3)              'Runs subroutines "FindNextBoard", passing in 3 and "_board2"
    End Sub
    
    Private Sub btnBoard2No4_Click(sender As Object, e As EventArgs) Handles btnBoard2No4.Click
        'Procedure run when "btnBoard2No4" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard2No4.BackgroundImage = image   'Sets the background of "btnBoard2No4" to image
        _board2(1, 0) = GamePlay(1, 0, _board2, 2)
        'Sets the value indexed at (0, 0) in _board2 to the result of the function "GamePlay", passing in 1, 0,
        '"_board2" and 2
        FindNextBoard(4)              'Runs subroutines "FindNextBoard", passing in 4 and "_board2"
    End Sub
    
    Private Sub btnBoard2No5_Click(sender As Object, e As EventArgs) Handles btnBoard2No5.Click
        'Procedure run when "btnBoard2No5" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard2No5.BackgroundImage = image   'Sets the background of "btnBoard2No5" to image
        _board1(1, 1) = GamePlay(1, 1, _board2, 2)
        'Sets the value indexed at (0, 0) in _board2 to the result of the function "GamePlay", passing in 1, 1,
        '"_board2" and 2
        FindNextBoard(5)              'Runs subroutines "FindNextBoard", passing in 5 and "_board2"
    End Sub
    
    Private Sub btnBoard2No6_Click(sender As Object, e As EventArgs) Handles btnBoard2No6.Click
        'Procedure run when "btnBoard2No6" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard2No6.BackgroundImage = image   'Sets the background of "btnBoard2No6" to image
        _board2(1, 2) = GamePlay(1, 2, _board2, 2)
        'Sets the value indexed at (0, 0) in _board2 to the result of the function "GamePlay", passing in 1, 2,
        '"_board2" and 2
        FindNextBoard(6)              'Runs subroutines "FindNextBoard", passing in 6 and "_board2"
    End Sub

    Private Sub btnBoard2No7_Click(sender As Object, e As EventArgs) Handles btnBoard2No7.Click
        'Procedure run when "btnBoard2No7" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard2No7.BackgroundImage = image   'Sets the background of "btnBoard2No7" to image
        _board2(2, 0) = GamePlay(2, 0, _board2, 2)
        'Sets the value indexed at (0, 0) in _board2 to the result of the function "GamePlay", passing in 2, 0,
        '"_board2" and 2
        FindNextBoard(7)              'Runs subroutines "FindNextBoard", passing in 7 and "_board2"
    End Sub

    Private Sub btnBoard2No8_Click(sender As Object, e As EventArgs) Handles btnBoard2No8.Click
        'Procedure run when "btnBoard2No8" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard2No8.BackgroundImage = image   'Sets the background of "btnBoard2No8" to image
        _board2(2, 1) = GamePlay(2, 1, _board2, 2)
        'Sets the value indexed at (0, 0) in _board2 to the result of the function "GamePlay", passing in 2, 1,
        '"_board2" and 2
        FindNextBoard(8)              'Runs subroutines "FindNextBoard", passing in 8 and "_board2"
    End Sub
    
    Private Sub btnBoard2No9_Click(sender As Object, e As EventArgs) Handles btnBoard2No9.Click
        'Procedure run when "btnBoard2No9" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard2No9.BackgroundImage = image   'Sets the background of "btnBoard2No9" to image
        _board2(2, 2) = GamePlay(2, 2, _board2, 2)
        'Sets the value indexed at (2, 2) in _board2 to the result of the function "GamePlay", passing in 2, 2,
        '"_board2" and 2
        FindNextBoard(9)              'Runs subroutines "FindNextBoard", passing in 9 and "_board2"
    End Sub
    
    Private Sub btnBoard3No1_Click(sender As Object, e As EventArgs) Handles btnBoard3No1.Click
        'Procedure run when "btnBoard3No1" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard3No1.BackgroundImage = image   'Sets the background of "btnBoard3No1" to image
        _board3(0, 0) = GamePlay(0, 0, _board3, 3)
        'Sets the value indexed at (0, 0) in _board3 to the result of the function "GamePlay", passing in 0, 0,
        '"_board3" and 3
        FindNextBoard(1)              'Runs subroutines "FindNextBoard", passing in 1 and "_board3"
    End Sub
    
    Private Sub btnBoard3No2_Click(sender As Object, e As EventArgs) Handles btnBoard3No2.Click
        'Procedure run when "btnBoard3No2" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard3No2.BackgroundImage = image   'Sets the background of "btnBoard3No2" to image
        _board3(0, 1) = GamePlay(0, 1, _board3, 3)
        'Sets the value indexed at (0, 1) in _board3 to the result of the function "GamePlay", passing in 0, 1,
        '"_board3" and 3
        FindNextBoard(2)              'Runs subroutines "FindNextBoard", passing in 2 and "_board3"
    End Sub
    
    Private Sub btnBoard3No3_Click(sender As Object, e As EventArgs) Handles btnBoard3No3.Click
        'Procedure run when "btnBoard3No3" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard3No3.BackgroundImage = image   'Sets the background of "btnBoard3No3" to image
        _board3(0, 2) = GamePlay(0, 2, _board3, 3)
        'Sets the value indexed at (0, 0) in _board3 to the result of the function "GamePlay", passing in 0, 2,
        '"_board3" and 3
        FindNextBoard(3)              'Runs subroutines "FindNextBoard", passing in 3 and "_board3"
    End Sub
    
    Private Sub btnBoard3No4_Click(sender As Object, e As EventArgs) Handles btnBoard3No4.Click
        'Procedure run when "btnBoard3No4" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard3No4.BackgroundImage = image   'Sets the background of "btnBoard3No4" to image
        _board3(1, 0) = GamePlay(1, 0, _board3, 3)
        'Sets the value indexed at (0, 0) in _board3 to the result of the function "GamePlay", passing in 1, 0,
        '"_board3" and 3
        FindNextBoard(4)              'Runs subroutines "FindNextBoard", passing in 4 and "_board3"
    End Sub
    
    Private Sub btnBoard3No5_Click(sender As Object, e As EventArgs) Handles btnBoard3No5.Click
        'Procedure run when "btnBoard3No5" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard3No5.BackgroundImage = image   'Sets the background of "btnBoard3No5" to image
        _board3(1, 1) = GamePlay(1, 1, _board3, 3)
        'Sets the value indexed at (0, 0) in _board3 to the result of the function "GamePlay", passing in 1, 1,
        '"_board1" and 3
        FindNextBoard(5)              'Runs subroutines "FindNextBoard", passing in 5 and "_board3"
    End Sub
    
    Private Sub btnBoard3No6_Click(sender As Object, e As EventArgs) Handles btnBoard3No6.Click
        'Procedure run when "btnBoard3No6" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard3No6.BackgroundImage = image   'Sets the background of "btnBoard3No6" to image
        _board3(1, 2) = GamePlay(1, 2, _board3, 3)
        'Sets the value indexed at (0, 0) in _board3 to the result of the function "GamePlay", passing in 1, 2,
        '"_board3" and 3
        FindNextBoard(6)              'Runs subroutines "FindNextBoard", passing in 6 and "_board3"
    End Sub

    Private Sub btnBoard3No7_Click(sender As Object, e As EventArgs) Handles btnBoard3No7.Click
        'Procedure run when "btnBoard3No7" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard3No7.BackgroundImage = image   'Sets the background of "btnBoard3No7" to image
        _board3(2, 0) = GamePlay(2, 0, _board3, 3)
        'Sets the value indexed at (0, 0) in _board3 to the result of the function "GamePlay", passing in 2, 0,
        '"_board1" and 3
        FindNextBoard(7)              'Runs subroutines "FindNextBoard", passing in 7 and "_board3"
    End Sub

    Private Sub btnBoard3No8_Click(sender As Object, e As EventArgs) Handles btnBoard3No8.Click
        'Procedure run when "btnBoard3No8" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard3No8.BackgroundImage = image   'Sets the background of "btnBoard3No8" to image
        _board3(2, 1) = GamePlay(2, 1, _board3, 3)
        'Sets the value indexed at (0, 0) in _board3 to the result of the function "GamePlay", passing in 2, 1,
        '"_board3" and 3
        FindNextBoard(8)              'Runs subroutines "FindNextBoard", passing in 8 and "_board3"
    End Sub
    
    Private Sub btnBoard3No9_Click(sender As Object, e As EventArgs) Handles btnBoard3No9.Click
        'Procedure run when "btnBoard3No9" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard3No9.BackgroundImage = image   'Sets the background of "btnBoard3No9" to image
        _board3(2, 2) = GamePlay(2, 2, _board3, 3)
        'Sets the value indexed at (2, 2) in _board3 to the result of the function "GamePlay", passing in 2, 2,
        '"_board3" and 3
        FindNextBoard(9)              'Runs subroutines "FindNextBoard", passing in 9 and "_board3"
    End Sub
    
    Private Sub btnBoard4No1_Click(sender As Object, e As EventArgs) Handles btnBoard4No1.Click
        'Procedure run when "btnBoard4No1" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard4No1.BackgroundImage = image   'Sets the background of "btnBoard4No1" to image
        _board4(0, 0) = GamePlay(0, 0, _board4, 4)
        'Sets the value indexed at (0, 0) in _board4 to the result of the function "GamePlay", passing in 0, 0,
        '"_board4" and 4
        FindNextBoard(1)              'Runs subroutines "FindNextBoard", passing in 1 and "_board4"
    End Sub
    
    Private Sub btnBoard4No2_Click(sender As Object, e As EventArgs) Handles btnBoard4No2.Click
        'Procedure run when "btnBoard4No2" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard4No2.BackgroundImage = image   'Sets the background of "btnBoard4No2" to image
        _board4(0, 1) = GamePlay(0, 1, _board4, 4)
        'Sets the value indexed at (0, 1) in _board4 to the result of the function "GamePlay", passing in 0, 1,
        '"_board4" and 4
        FindNextBoard(2)              'Runs subroutines "FindNextBoard", passing in 2 and "_board4"
    End Sub
    
    Private Sub btnBoard4No3_Click(sender As Object, e As EventArgs) Handles btnBoard4No3.Click
        'Procedure run when "btnBoard4No3" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard4No3.BackgroundImage = image   'Sets the background of "btnBoard4No3" to image
        _board4(0, 2) = GamePlay(0, 2, _board4, 4)
        'Sets the value indexed at (0, 0) in _board4 to the result of the function "GamePlay", passing in 0, 2,
        '"_board4" and 4
        FindNextBoard(3)              'Runs subroutines "FindNextBoard", passing in 3 and "_board4"
    End Sub
    
    Private Sub btnBoard4No4_Click(sender As Object, e As EventArgs) Handles btnBoard4No4.Click
        'Procedure run when "btnBoard4No4" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard4No4.BackgroundImage = image   'Sets the background of "btnBoard4No4" to image
        _board4(1, 0) = GamePlay(1, 0, _board4, 4)
        'Sets the value indexed at (0, 0) in _board4 to the result of the function "GamePlay", passing in 1, 0,
        '"_board4" and 4
        FindNextBoard(4)              'Runs subroutines "FindNextBoard", passing in 4 and "_board4"
    End Sub
    
    Private Sub btnBoard4No5_Click(sender As Object, e As EventArgs) Handles btnBoard4No5.Click
        'Procedure run when "btnBoard4No5" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard4No5.BackgroundImage = image   'Sets the background of "btnBoard1No5" to image
        _board4(1, 1) = GamePlay(1, 1, _board4, 4)
        'Sets the value indexed at (0, 0) in _board4 to the result of the function "GamePlay", passing in 1, 1,
        '"_board4" and 4
        FindNextBoard(5)              'Runs subroutines "FindNextBoard", passing in 5 and "_board4"
    End Sub
    
    Private Sub btnBoard4No6_Click(sender As Object, e As EventArgs) Handles btnBoard4No6.Click
        'Procedure run when "btnBoard4No6" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard4No6.BackgroundImage = image   'Sets the background of "btnBoard4No6" to image
        _board4(1, 2) = GamePlay(1, 2, _board4, 4)
        'Sets the value indexed at (0, 0) in _board4 to the result of the function "GamePlay", passing in 1, 2,
        '"_board4" and 4
        FindNextBoard(6)              'Runs subroutines "FindNextBoard", passing in 6 and "_board4"
    End Sub

    Private Sub btnBoard4No7_Click(sender As Object, e As EventArgs) Handles btnBoard4No7.Click
        'Procedure run when "btnBoard4No7" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard4No7.BackgroundImage = image   'Sets the background of "btnBoard4No7" to image
        _board4(2, 0) = GamePlay(2, 0, _board4, 4)
        'Sets the value indexed at (0, 0) in _board4 to the result of the function "GamePlay", passing in 2, 0,
        '"_board4" and 4
        FindNextBoard(7)              'Runs subroutines "FindNextBoard", passing in 7 and "_board4"
    End Sub

    Private Sub btnBoard4No8_Click(sender As Object, e As EventArgs) Handles btnBoard4No8.Click
        'Procedure run when "btnBoard4No8" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard4No8.BackgroundImage = image   'Sets the background of "btnBoard4No8" to image
        _board4(2, 1) = GamePlay(2, 1, _board4, 4)
        'Sets the value indexed at (0, 0) in _board1 to the result of the function "GamePlay", passing in 2, 1,
        '"_board4" and 4
        FindNextBoard(8)              'Runs subroutines "FindNextBoard", passing in 8 and "_board4"
    End Sub
    
    Private Sub btnBoard4No9_Click(sender As Object, e As EventArgs) Handles btnBoard4No9.Click
        'Procedure run when "btnBoard4No9" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard4No9.BackgroundImage = image   'Sets the background of "btnBoard4No9" to image
        _board4(2, 2) = GamePlay(2, 2, _board4, 4)
        'Sets the value indexed at (2, 2) in _board1 to the result of the function "GamePlay", passing in 2, 2,
        '"_board4" and 4
        FindNextBoard(9)              'Runs subroutines "FindNextBoard", passing in 9 and "_board4"
    End Sub
    
    Private Sub btnBoard5No1_Click(sender As Object, e As EventArgs) Handles btnBoard5No1.Click
        'Procedure run when "btnBoard5No1" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard5No1.BackgroundImage = image   'Sets the background of "btnBoard5No1" to image
        _board5(0, 0) = GamePlay(0, 0, _board5, 5)
        'Sets the value indexed at (0, 0) in _board5 to the result of the function "GamePlay", passing in 0, 0,
        '"_board5" and 5
        FindNextBoard(1)              'Runs subroutines "FindNextBoard", passing in 1 and "_board5"
    End Sub
    
    Private Sub btnBoard5No2_Click(sender As Object, e As EventArgs) Handles btnBoard5No2.Click
        'Procedure run when "btnBoard5No2" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard5No2.BackgroundImage = image   'Sets the background of "btnBoard5No2" to image
        _board5(0, 1) = GamePlay(0, 1, _board5, 5)
        'Sets the value indexed at (0, 1) in _board5 to the result of the function "GamePlay", passing in 0, 1,
        '"_board1" and 5
        FindNextBoard(2)              'Runs subroutines "FindNextBoard", passing in 2 and "_board5"
    End Sub
    
    Private Sub btnBoard5No3_Click(sender As Object, e As EventArgs) Handles btnBoard5No3.Click
        'Procedure run when "btnBoard5No3" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard5No3.BackgroundImage = image   'Sets the background of "btnBoard5No3" to image
        _board5(0, 2) = GamePlay(0, 2, _board5, 5)
        'Sets the value indexed at (0, 0) in _board5 to the result of the function "GamePlay", passing in 0, 2,
        '"_board5" and 5
        FindNextBoard(3)              'Runs subroutines "FindNextBoard", passing in 3 and "_board5"
    End Sub
    
    Private Sub btnBoard5No4_Click(sender As Object, e As EventArgs) Handles btnBoard5No4.Click
        'Procedure run when "btnBoard5No4" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard5No4.BackgroundImage = image   'Sets the background of "btnBoard5No4" to image
        _board5(1, 0) = GamePlay(1, 0, _board5, 5)
        'Sets the value indexed at (0, 0) in _board1 to the result of the function "GamePlay", passing in 1, 0,
        '"_board5" and 5
        FindNextBoard(4)              'Runs subroutines "FindNextBoard", passing in 4 and "_board5"
    End Sub
    
    Private Sub btnBoard5No5_Click(sender As Object, e As EventArgs) Handles btnBoard5No5.Click
        'Procedure run when "btnBoard5No5" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard5No5.BackgroundImage = image   'Sets the background of "btnBoard5No5" to image
        _board5(1, 1) = GamePlay(1, 1, _board5, 5)
        'Sets the value indexed at (0, 0) in _board5 to the result of the function "GamePlay", passing in 1, 1,
        '"_board5" and 5
        FindNextBoard(5)              'Runs subroutines "FindNextBoard", passing in 5 and "_board5"
    End Sub
    
    Private Sub btnBoard5No6_Click(sender As Object, e As EventArgs) Handles btnBoard5No6.Click
        'Procedure run when "btnBoard1No6" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard5No6.BackgroundImage = image   'Sets the background of "btnBoard5No6" to image
        _board5(1, 2) = GamePlay(1, 2, _board5, 5)
        'Sets the value indexed at (0, 0) in _board5 to the result of the function "GamePlay", passing in 1, 2,
        '"_board5" and 5
        FindNextBoard(6)              'Runs subroutines "FindNextBoard", passing in 6 and "_board5"
    End Sub

    Private Sub btnBoard5No7_Click(sender As Object, e As EventArgs) Handles btnBoard5No7.Click
        'Procedure run when "btnBoard5No7" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard5No7.BackgroundImage = image   'Sets the background of "btnBoard5No7" to image
        _board5(2, 0) = GamePlay(2, 0, _board5, 5)
        'Sets the value indexed at (0, 0) in _board5 to the result of the function "GamePlay", passing in 2, 0,
        '"_board5" and 5
        FindNextBoard(7)              'Runs subroutines "FindNextBoard", passing in 7 and "_board5"
    End Sub

    Private Sub btnBoard5No8_Click(sender As Object, e As EventArgs) Handles btnBoard5No8.Click
        'Procedure run when "btnBoard5No8" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard5No8.BackgroundImage = image   'Sets the background of "btnBoard5No8" to image
        _board5(2, 1) = GamePlay(2, 1, _board5, 5)
        'Sets the value indexed at (0, 0) in _board5 to the result of the function "GamePlay", passing in 2, 1,
        '"_board5" and 5
        FindNextBoard(8)              'Runs subroutines "FindNextBoard", passing in 8 and "_board1"
    End Sub
    
    Private Sub btnBoard5No9_Click(sender As Object, e As EventArgs) Handles btnBoard5No9.Click
        'Procedure run when "btnBoard5No9" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard5No9.BackgroundImage = image   'Sets the background of "btnBoard5No9" to image
        _board5(2, 2) = GamePlay(2, 2, _board5, 5)
        'Sets the value indexed at (2, 2) in _board5 to the result of the function "GamePlay", passing in 2, 2,
        '"_board5" and 5
        FindNextBoard(9)              'Runs subroutines "FindNextBoard", passing in 9 and "_board5"
    End Sub
    
    Private Sub btnBoard6No1_Click(sender As Object, e As EventArgs) Handles btnBoard6No1.Click
        'Procedure run when "btnBoard6No1" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard6No1.BackgroundImage = image   'Sets the background of "btnBoard6No1" to image
        _board6(0, 0) = GamePlay(0, 0, _board6, 6)
        'Sets the value indexed at (0, 0) in _board6 to the result of the function "GamePlay", passing in 0, 0,
        '"_board6" and 6
        FindNextBoard(1)              'Runs subroutines "FindNextBoard", passing in 1 and "_board6"
    End Sub
    
    Private Sub btnBoar6No2_Click(sender As Object, e As EventArgs) Handles btnBoard6No2.Click
        'Procedure run when "btnBoard6No2" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard6No2.BackgroundImage = image   'Sets the background of "btnBoard6No2" to image
        _board6(0, 1) = GamePlay(0, 1, _board6, 6)
        'Sets the value indexed at (0, 1) in _board6 to the result of the function "GamePlay", passing in 0, 1,
        '"_board6" and 6
        FindNextBoard(2)              'Runs subroutines "FindNextBoard", passing in 2 and "_board6"
    End Sub
    
    Private Sub btnBoard6No3_Click(sender As Object, e As EventArgs) Handles btnBoard6No3.Click
        'Procedure run when "btnBoard6No3" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard6No3.BackgroundImage = image   'Sets the background of "btnBoard6No3" to image
        _board6(0, 2) = GamePlay(0, 2, _board6, 6)
        'Sets the value indexed at (0, 0) in _board6 to the result of the function "GamePlay", passing in 0, 2,
        '"_board6" and 6
        FindNextBoard(3)              'Runs subroutines "FindNextBoard", passing in 3 and "_board6"
    End Sub
    
    Private Sub btnBoard6No4_Click(sender As Object, e As EventArgs) Handles btnBoard6No4.Click
        'Procedure run when "btnBoard1No4" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard6No4.BackgroundImage = image   'Sets the background of "btnBoard6No4" to image
        _board6(1, 0) = GamePlay(1, 0, _board6, 6)
        'Sets the value indexed at (0, 0) in _board6 to the result of the function "GamePlay", passing in 1, 0,
        '"_board6" and 6
        FindNextBoard(4)              'Runs subroutines "FindNextBoard", passing in 4 and "_board6"
    End Sub
    
    Private Sub btnBoard6No5_Click(sender As Object, e As EventArgs) Handles btnBoard6No5.Click
        'Procedure run when "btnBoard6No5" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard6No5.BackgroundImage = image   'Sets the background of "btnBoard6No5" to image
        _board6(1, 1) = GamePlay(1, 1, _board6, 6)
        'Sets the value indexed at (0, 0) in _board6 to the result of the function "GamePlay", passing in 1, 1,
        '"_board6" and 6
        FindNextBoard(5)              'Runs subroutines "FindNextBoard", passing in 5 and "_board6"
    End Sub
    
    Private Sub btnBoard6No6_Click(sender As Object, e As EventArgs) Handles btnBoard6No6.Click
        'Procedure run when "btnBoard6No6" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard6No6.BackgroundImage = image   'Sets the background of "btnBoard6No6" to image
        _board6(1, 2) = GamePlay(1, 2, _board6, 6)
        'Sets the value indexed at (0, 0) in _board6 to the result of the function "GamePlay", passing in 1, 2,
        '"_board1" and 6
        FindNextBoard(6)              'Runs subroutines "FindNextBoard", passing in 6 and "_board6"
    End Sub

    Private Sub btnBoard6No7_Click(sender As Object, e As EventArgs) Handles btnBoard6No7.Click
        'Procedure run when "btnBoard6No7" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard6No7.BackgroundImage = image   'Sets the background of "btnBoard6No7" to image
        _board6(2, 0) = GamePlay(2, 0, _board6, 6)
        'Sets the value indexed at (0, 0) in _board6 to the result of the function "GamePlay", passing in 2, 0,
        '"_board1" and 6
        FindNextBoard(7)              'Runs subroutines "FindNextBoard", passing in 7 and "_board6"
    End Sub

    Private Sub btnBoard6No8_Click(sender As Object, e As EventArgs) Handles btnBoard6No8.Click
        'Procedure run when "btnBoard6No8" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard6No8.BackgroundImage = image   'Sets the background of "btnBoard6No8" to image
        _board6(2, 1) = GamePlay(2, 1, _board6, 6)
        'Sets the value indexed at (0, 0) in _board6 to the result of the function "GamePlay", passing in 2, 1,
        '"_board6" and 6
        FindNextBoard(8)              'Runs subroutines "FindNextBoard", passing in 8 and "_board6"
    End Sub
    
    Private Sub btnBoard6No9_Click(sender As Object, e As EventArgs) Handles btnBoard6No9.Click
        'Procedure run when "btnBoard6No9" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard6No9.BackgroundImage = image   'Sets the background of "btnBoard6No9" to image
        _board6(2, 2) = GamePlay(2, 2, _board6, 6)
        'Sets the value indexed at (2, 2) in _board6 to the result of the function "GamePlay", passing in 2, 2,
        '"_board6" and 6
        FindNextBoard(9)              'Runs subroutines "FindNextBoard", passing in 9 and "_board6"
    End Sub
    
    Private Sub btnBoard7No1_Click(sender As Object, e As EventArgs) Handles btnBoard7No1.Click
        'Procedure run when "btnBoard7No1" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard7No1.BackgroundImage = image   'Sets the background of "btnBoard7No1" to image
        _board7(0, 0) = GamePlay(0, 0, _board7, 7)
        'Sets the value indexed at (0, 0) in _board7 to the result of the function "GamePlay", passing in 0, 0,
        '"_board1" and 7
        FindNextBoard(1)              'Runs subroutines "FindNextBoard", passing in 1 and "_board7"
    End Sub
    
    Private Sub btnBoard7No2_Click(sender As Object, e As EventArgs) Handles btnBoard7No2.Click
        'Procedure run when "btnBoard7No2" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard7No2.BackgroundImage = image   'Sets the background of "btnBoard7No2" to image
        _board7(0, 1) = GamePlay(0, 1, _board7, 7)
        'Sets the value indexed at (0, 1) in _board7 to the result of the function "GamePlay", passing in 0, 1,
        '"_board7" and 7
        FindNextBoard(2)              'Runs subroutines "FindNextBoard", passing in 2 and "_board7"
    End Sub
    
    Private Sub btnBoard7No3_Click(sender As Object, e As EventArgs) Handles btnBoard7No3.Click
        'Procedure run when "btnBoard7No3" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard7No3.BackgroundImage = image   'Sets the background of "btnBoard7No3" to image
        _board7(0, 2) = GamePlay(0, 2, _board7, 7)
        'Sets the value indexed at (0, 0) in _board7 to the result of the function "GamePlay", passing in 0, 2,
        '"_board7" and 7
        FindNextBoard(3)              'Runs subroutines "FindNextBoard", passing in 3 and "_board7"
    End Sub
    
    Private Sub btnBoard7No4_Click(sender As Object, e As EventArgs) Handles btnBoard7No4.Click
        'Procedure run when "btnBoard7No4" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard7No4.BackgroundImage = image   'Sets the background of "btnBoard7No4" to image
        _board7(1, 0) = GamePlay(1, 0, _board7, 7)
        'Sets the value indexed at (0, 0) in _board7 to the result of the function "GamePlay", passing in 1, 0,
        '"_board7" and 7
        FindNextBoard(4)              'Runs subroutines "FindNextBoard", passing in 4 and "_board7"
    End Sub
    
    Private Sub btnBoard7No5_Click(sender As Object, e As EventArgs) Handles btnBoard7No5.Click
        'Procedure run when "btnBoard7No5" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard7No5.BackgroundImage = image   'Sets the background of "btnBoard7No5" to image
        _board7(1, 1) = GamePlay(1, 1, _board7, 7)
        'Sets the value indexed at (0, 0) in _board7 to the result of the function "GamePlay", passing in 1, 1,
        '"_board7" and 7
        FindNextBoard(5)              'Runs subroutines "FindNextBoard", passing in 5 and "_board7"
    End Sub
    
    Private Sub btnBoard7No6_Click(sender As Object, e As EventArgs) Handles btnBoard7No6.Click
        'Procedure run when "btnBoard7No6" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard7No6.BackgroundImage = image   'Sets the background of "btnBoard7No6" to image
        _board7(1, 2) = GamePlay(1, 2, _board7, 7)
        'Sets the value indexed at (0, 0) in _board7 to the result of the function "GamePlay", passing in 1, 2,
        '"_board7" and 7
        FindNextBoard(6)              'Runs subroutines "FindNextBoard", passing in 6 and "_board7"
    End Sub

    Private Sub btnBoard7No7_Click(sender As Object, e As EventArgs) Handles btnBoard7No7.Click
        'Procedure run when "btnBoard1No7" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard7No7.BackgroundImage = image   'Sets the background of "btnBoard7No7" to image
        _board7(2, 0) = GamePlay(2, 0, _board7, 7)
        'Sets the value indexed at (0, 0) in _board7 to the result of the function "GamePlay", passing in 2, 0,
        '"_board7" and 7
        FindNextBoard(7)              'Runs subroutines "FindNextBoard", passing in 7 and "_board7"
    End Sub

    Private Sub btnBoard7No8_Click(sender As Object, e As EventArgs) Handles btnBoard7No8.Click
        'Procedure run when "btnBoard7No8" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard7No8.BackgroundImage = image   'Sets the background of "btnBoard7No8" to image
        _board7(2, 1) = GamePlay(2, 1, _board7, 7)
        'Sets the value indexed at (0, 0) in _board7 to the result of the function "GamePlay", passing in 2, 1,
        '"_board7" and 7
        FindNextBoard(8)              'Runs subroutines "FindNextBoard", passing in 8 and "_board7"
    End Sub
    
    Private Sub btnBoard7No9_Click(sender As Object, e As EventArgs) Handles btnBoard7No9.Click
        'Procedure run when "btnBoard7No9" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard7No9.BackgroundImage = image   'Sets the background of "btnBoard7No9" to image
        _board7(2, 2) = GamePlay(2, 2, _board7, 7)
        'Sets the value indexed at (2, 2) in _board7 to the result of the function "GamePlay", passing in 2, 2,
        '"_board7" and 7
        FindNextBoard(9)              'Runs subroutines "FindNextBoard", passing in 9 and "_board7"
    End Sub
    
    Private Sub btnBoard8No1_Click(sender As Object, e As EventArgs) Handles btnBoard8No1.Click
        'Procedure run when "btnBoard8No1" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard8No1.BackgroundImage = image   'Sets the background of "btnBoard8No1" to image
        _board8(0, 0) = GamePlay(0, 0, _board8, 8)
        'Sets the value indexed at (0, 0) in _board8 to the result of the function "GamePlay", passing in 0, 0,
        '"_board8" and 8
        FindNextBoard(1)              'Runs subroutines "FindNextBoard", passing in 1 and "_board8"
    End Sub
    
    Private Sub btnBoard8No2_Click(sender As Object, e As EventArgs) Handles btnBoard8No2.Click
        'Procedure run when "btnBoard8No2" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard8No2.BackgroundImage = image   'Sets the background of "btnBoard8No2" to image
        _board8(0, 1) = GamePlay(0, 1, _board8, 8)
        'Sets the value indexed at (0, 1) in _board8 to the result of the function "GamePlay", passing in 0, 1,
        '"_board8" and 8
        FindNextBoard(2)              'Runs subroutines "FindNextBoard", passing in 2 and "_board8"
    End Sub
    
    Private Sub btnBoard8No3_Click(sender As Object, e As EventArgs) Handles btnBoard8No3.Click
        'Procedure run when "btnBoard8No3" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard8No3.BackgroundImage = image   'Sets the background of "btnBoard8No3" to image
        _board8(0, 2) = GamePlay(0, 2, _board8, 8)
        'Sets the value indexed at (0, 0) in _board8 to the result of the function "GamePlay", passing in 0, 2,
        '"_board8" and 8
        FindNextBoard(3)              'Runs subroutines "FindNextBoard", passing in 3 and "_board8"
    End Sub
    
    Private Sub btnBoard8No4_Click(sender As Object, e As EventArgs) Handles btnBoard8No4.Click
        'Procedure run when "btnBoard8No4" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard8No4.BackgroundImage = image   'Sets the background of "btnBoard8No4" to image
        _board8(1, 0) = GamePlay(1, 0, _board8, 8)
        'Sets the value indexed at (0, 0) in _board8 to the result of the function "GamePlay", passing in 1, 0,
        '"_board1" and 8
        FindNextBoard(4)              'Runs subroutines "FindNextBoard", passing in 4 and "_board8"
    End Sub
    
    Private Sub btnBoard8No5_Click(sender As Object, e As EventArgs) Handles btnBoard8No5.Click
        'Procedure run when "btnBoard8No5" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard8No5.BackgroundImage = image   'Sets the background of "btnBoard8No5" to image
        _board8(1, 1) = GamePlay(1, 1, _board8, 8)
        'Sets the value indexed at (0, 0) in _board8 to the result of the function "GamePlay", passing in 1, 1,
        '"_board8" and 1
        FindNextBoard(5)              'Runs subroutines "FindNextBoard", passing in 5 and "_board8"
    End Sub
    
    Private Sub btnBoard8No6_Click(sender As Object, e As EventArgs) Handles btnBoard8No6.Click
        'Procedure run when "btnBoard8No6" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard8No6.BackgroundImage = image   'Sets the background of "btnBoard8No6" to image
        _board8(1, 2) = GamePlay(1, 2, _board8, 8)
        'Sets the value indexed at (0, 0) in _board8 to the result of the function "GamePlay", passing in 1, 2,
        '"_board8" and 1
        FindNextBoard(6)              'Runs subroutines "FindNextBoard", passing in 6 and "_board8"
    End Sub

    Private Sub btnBoard8No7_Click(sender As Object, e As EventArgs) Handles btnBoard8No7.Click
        'Procedure run when "btnBoard8No7" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard8No7.BackgroundImage = image   'Sets the background of "btnBoard8No7" to image
        _board8(2, 0) = GamePlay(2, 0, _board8, 8)
        'Sets the value indexed at (0, 0) in _board8 to the result of the function "GamePlay", passing in 2, 0,
        '"_board8" and 8
        FindNextBoard(7)              'Runs subroutines "FindNextBoard", passing in 7 and "_board8"
    End Sub

    Private Sub btnBoard8No8_Click(sender As Object, e As EventArgs) Handles btnBoard8No8.Click
        'Procedure run when "btnBoard8No8" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard8No8.BackgroundImage = image   'Sets the background of "btnBoard8No8" to image
        _board8(2, 1) = GamePlay(2, 1, _board8, 8)
        'Sets the value indexed at (0, 0) in _board8 to the result of the function "GamePlay", passing in 2, 1,
        '"_board8" and 8
        FindNextBoard(8)              'Runs subroutines "FindNextBoard", passing in 8 and "_board8"
    End Sub
    
    Private Sub btnBoard8No9_Click(sender As Object, e As EventArgs) Handles btnBoard8No9.Click
        'Procedure run when "btnBoard8No9" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard8No9.BackgroundImage = image   'Sets the background of "btnBoard8No9" to image
        _board8(2, 2) = GamePlay(2, 2, _board8, 8)
        'Sets the value indexed at (2, 2) in _board8 to the result of the function "GamePlay", passing in 2, 2,
        '"_board8" and 8
        FindNextBoard(9)              'Runs subroutines "FindNextBoard", passing in 9 and "_board8"
    End Sub
    
    Private Sub btnBoard9No1_Click(sender As Object, e As EventArgs) Handles btnBoard9No1.Click
        'Procedure run when "btnBoard9No1" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard9No1.BackgroundImage = image   'Sets the background of "btnBoard9No1" to image
        _board9(0, 0) = GamePlay(0, 0, _board9, 9)
        'Sets the value indexed at (0, 0) in _board1 to the result of the function "GamePlay", passing in 0, 0,
        '"_board9" and 9
        FindNextBoard(1)              'Runs subroutines "FindNextBoard", passing in 1 and "_board9"
    End Sub
    
    Private Sub btnBoard9No2_Click(sender As Object, e As EventArgs) Handles btnBoard9No2.Click
        'Procedure run when "btnBoard9No2" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard9No2.BackgroundImage = image   'Sets the background of "btnBoard9No2" to image
        _board9(0, 1) = GamePlay(0, 1, _board9, 9)
        'Sets the value indexed at (0, 1) in _board1 to the result of the function "GamePlay", passing in 0, 1,
        '"_board1" and 9
        FindNextBoard(2)              'Runs subroutines "FindNextBoard", passing in 2 and "_board9"
    End Sub
    
    Private Sub btnBoard9No3_Click(sender As Object, e As EventArgs) Handles btnBoard9No3.Click
        'Procedure run when "btnBoard9No3" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard9No3.BackgroundImage = image   'Sets the background of "btnBoard9No3" to image
        _board9(0, 2) = GamePlay(0, 2, _board9, 9)
        'Sets the value indexed at (0, 0) in _board9 to the result of the function "GamePlay", passing in 0, 2,
        '"_board9" and 9
        FindNextBoard(3)              'Runs subroutines "FindNextBoard", passing in 3 and "_board9"
    End Sub
    
    Private Sub btnBoard9No4_Click(sender As Object, e As EventArgs) Handles btnBoard9No4.Click
        'Procedure run when "btnBoard9No4" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard9No4.BackgroundImage = image   'Sets the background of "btnBoard9No4" to image
        _board9(1, 0) = GamePlay(1, 0, _board9, 9)
        'Sets the value indexed at (0, 0) in _board9 to the result of the function "GamePlay", passing in 1, 0,
        '"_board9" and 9
        FindNextBoard(4)              'Runs subroutines "FindNextBoard", passing in 4 and "_board9"
    End Sub
    
    Private Sub btnBoard9No5_Click(sender As Object, e As EventArgs) Handles btnBoard9No5.Click
        'Procedure run when "btnBoard9No5" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard9No5.BackgroundImage = image   'Sets the background of "btnBoard9No5" to image
        _board9(1, 1) = GamePlay(1, 1, _board9, 9)
        'Sets the value indexed at (0, 0) in _board9 to the result of the function "GamePlay", passing in 1, 1,
        '"_board9" and 9
        FindNextBoard(5)              'Runs subroutines "FindNextBoard", passing in 5 and "_board9"
    End Sub
    
    Private Sub btnBoard9No6_Click(sender As Object, e As EventArgs) Handles btnBoard9No6.Click
        'Procedure run when "btnBoard9No6" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard9No6.BackgroundImage = image   'Sets the background of "btnBoard9No6" to image
        _board9(1, 2) = GamePlay(1, 2, _board9, 9)
        'Sets the value indexed at (0, 0) in _board9 to the result of the function "GamePlay", passing in 1, 2,
        '"_board9" and 9
        FindNextBoard(6)              'Runs subroutines "FindNextBoard", passing in 6 and "_board9"
    End Sub

    Private Sub btnBoard9No7_Click(sender As Object, e As EventArgs) Handles btnBoard9No7.Click
        'Procedure run when "btnBoard9No7" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard9No7.BackgroundImage = image   'Sets the background of "btnBoard9No7" to image
        _board9(2, 0) = GamePlay(2, 0, _board9, 9)
        'Sets the value indexed at (0, 0) in _board9 to the result of the function "GamePlay", passing in 2, 0,
        '"_board9" and 9
        FindNextBoard(7)              'Runs subroutines "FindNextBoard", passing in 7 and "_board9"
    End Sub

    Private Sub btnBoard9No8_Click(sender As Object, e As EventArgs) Handles btnBoard9No8.Click
        'Procedure run when "btnBoard9No8" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard9No8.BackgroundImage = image   'Sets the background of "btnBoard9No8" to image
        _board9(2, 1) = GamePlay(2, 1, _board9, 9)
        'Sets the value indexed at (0, 0) in _board9 to the result of the function "GamePlay", passing in 2, 1,
        '"_board9" and 9
        FindNextBoard(8)              'Runs subroutines "FindNextBoard", passing in 8 and "_board9"
    End Sub
    
    Private Sub btnBoard9No9_Click(sender As Object, e As EventArgs) Handles btnBoard9No9.Click
        'Procedure run when "btnBoard9No9" is clicked
        Dim image As New Bitmap(60, 60)        'Defines "image" as bitmap 60px by 60px
        image = DrawGraphics(1, _playerGo)     'Executes procedure "DrawGraphics", passing in 1 and "_playerGo"
        btnBoard9No9.BackgroundImage = image   'Sets the background of "btnBoard9No9" to image
        _board9(2, 2) = GamePlay(2, 2, _board9, 9)
        'Sets the value indexed at (2, 2) in _board9 to the result of the function "GamePlay", passing in 2, 2,
        '"_board9" and 9
        FindNextBoard(9)              'Runs subroutines "FindNextBoard", passing in 9 and "_board9"
    End Sub
    
    Private Sub FindNextBoard(boardNumber)
        Dim ctrl As Control
        Dim image As New Bitmap(656, 656)
        Dim penPoints As Point()
                
        If (_board1Win = True And boardNumber = 1) Or (_board2Win = True And boardNumber = 2) Or _
           (_board3Win = True And boardNumber = 3) Or (_board4Win = True And boardNumber = 4) Or _
           (_board5Win = True And boardNumber = 5) Or (_board6Win = True And boardNumber = 6) Or _
           (_board7Win = True And boardNumber = 7) Or (_board8Win = True And boardNumber = 8) Or _
           (_board9Win = True And boardNumber = 9) Then
            For Each ctrl In Controls
                If TypeOf ctrl Is Button Then
                    ctrl.Enabled = True
                End If
            Next
            btnStart.Enabled = False
            btnBoard1.Enabled = False
            btnBoard2.Enabled = False
            btnBoard3.Enabled = False
            btnBoard4.Enabled = False
            btnBoard5.Enabled = False
            btnBoard6.Enabled = False
            btnBoard7.Enabled = False
            btnBoard8.Enabled = False
            btnBoard9.Enabled = False
            penPoints = {New Point(10, 10), New Point(646, 10), New Point(646, 646), New Point(10, 646), _ 
                         New Point(10, 10)}
            'Sets the value of "penPoints" to the points (10, 10), (646, 10), (646, 646), (10, 646) and (10, 646)
        Else 
            For Each ctrl In Controls
                If TypeOf ctrl Is Button Then
                    ctrl.Enabled = False
                End If
            Next
            Select Case boardNumber
                Case 1                             'If "boardNumber" is 1 then all the buttons for board 1 are enabled
                    btnBoard1No1.Enabled = True
                    btnBoard1No2.Enabled = True
                    btnBoard1No3.Enabled = True
                    btnBoard1No4.Enabled = True
                    btnBoard1No5.Enabled = True
                    btnBoard1No6.Enabled = True
                    btnBoard1No7.Enabled = True
                    btnBoard1No8.Enabled = True
                    btnBoard1No9.Enabled = True
                    penPoints = {New Point(10, 10), New Point(222, 10), New Point(222, 222), New Point(10, 222), _
                                 New Point(10, 10)}
                    'Sets the value of "penPoints" to the points (10, 10), (222, 10), (222, 222), (10, 222)
                    'And (10, 10)
                Case 2                             'If "boardNumber" is 2 then all the buttons for board 2 are enabled
                    btnBoard2No1.Enabled = True
                    btnBoard2No2.Enabled = True
                    btnBoard2No3.Enabled = True
                    btnBoard2No4.Enabled = True
                    btnBoard2No5.Enabled = True
                    btnBoard2No6.Enabled = True
                    btnBoard2No7.Enabled = True
                    btnBoard2No8.Enabled = True
                    btnBoard2No9.Enabled = True
                    penPoints = {New Point(222, 10), New Point(432, 10), New Point(432, 222), New Point(222, 222), _
                                 New Point(222, 10)}
                    'Sets the value of "penPoints" to the points (222, 10), (402, 10), (432, 222), (222, 222)
                    'And (222, 10)
                Case 3                             'If "boardNumber" is 3 then all the buttons for board 3 are enabled
                    btnBoard3No1.Enabled = True
                    btnBoard3No2.Enabled = True
                    btnBoard3No3.Enabled = True
                    btnBoard3No4.Enabled = True
                    btnBoard3No5.Enabled = True
                    btnBoard3No6.Enabled = True
                    btnBoard3No7.Enabled = True
                    btnBoard3No8.Enabled = True
                    btnBoard3No9.Enabled = True
                    penPoints = {New Point(432, 10), New Point(646, 10), New Point(646, 222), New Point(432, 222), _
                                 New Point(432, 10)}
                    'Sets the value of "penPoints" to the points (432, 10), (646, 10), (646, 222), (432, 222)
                    'And (432, 10)
                Case 4                             'If "boardNumber" is 4 then all the buttons for board 4 are enabled
                    btnBoard4No1.Enabled = True
                    btnBoard4No2.Enabled = True
                    btnBoard4No3.Enabled = True
                    btnBoard4No4.Enabled = True
                    btnBoard4No5.Enabled = True
                    btnBoard4No6.Enabled = True
                    btnBoard4No7.Enabled = True
                    btnBoard4No8.Enabled = True
                    btnBoard4No9.Enabled = True
                    penPoints = {New Point(10, 222), New Point(222, 222), New Point(222, 432), New Point(10, 432), _
                                 New Point(10, 222)}
                    'Sets the value of "penPoints" to the points (10, 222), (222, 222), (222, 432), (10, 432)
                    'And (10, 222)
                Case 5                             'If "boardNumber" is 5 then all the buttons for board 5 are enabled
                    btnBoard5No1.Enabled = True
                    btnBoard5No2.Enabled = True
                    btnBoard5No3.Enabled = True
                    btnBoard5No4.Enabled = True
                    btnBoard5No5.Enabled = True
                    btnBoard5No6.Enabled = True
                    btnBoard5No7.Enabled = True
                    btnBoard5No8.Enabled = True
                    btnBoard5No9.Enabled = True
                    penPoints = {New Point(222, 222), New Point(432, 222), New Point(432, 432), New Point(222, 432), _
                                 New Point(222, 222)}
                    'Sets the value of "penPoints" to the points (222, 222), (402, 222), (432, 432), (222, 432)
                    'And (222, 222)
                Case 6                             'If "boardNumber" is 6 then all the buttons for board 6 are enabled
                    btnBoard6No1.Enabled = True
                    btnBoard6No2.Enabled = True
                    btnBoard6No3.Enabled = True
                    btnBoard6No4.Enabled = True
                    btnBoard6No5.Enabled = True
                    btnBoard6No6.Enabled = True
                    btnBoard6No7.Enabled = True
                    btnBoard6No8.Enabled = True
                    btnBoard6No9.Enabled = True
                    penPoints = {New Point(432, 222), New Point(646, 222), New Point(646, 432), New Point(432, 432), _
                                 New Point(432, 222)}
                    'Sets the value of "penPoints" to the points (432, 222), (646, 222), (646, 432), (432, 432)
                    'And (432, 222)
                Case 7                             'If "boardNumber" is 7 then all the buttons for board 7 are enabled
                    btnBoard7No1.Enabled = True
                    btnBoard7No2.Enabled = True
                    btnBoard7No3.Enabled = True
                    btnBoard7No4.Enabled = True
                    btnBoard7No5.Enabled = True
                    btnBoard7No6.Enabled = True
                    btnBoard7No7.Enabled = True
                    btnBoard7No8.Enabled = True
                    btnBoard7No9.Enabled = True
                    penPoints = {New Point(10, 432), New Point(222, 432), New Point(222, 646), New Point(10, 646), _
                                 New Point(10, 432)}
                    'Sets the value of "penPoints" to the points (10, 432), (222, 432), (222, 646), (10, 646)
                    'And (10, 432)
                Case 8                             'If "boardNumber" is 8 then all the buttons for board 8 are enabled
                    btnBoard8No1.Enabled = True
                    btnBoard8No2.Enabled = True
                    btnBoard8No3.Enabled = True
                    btnBoard8No4.Enabled = True
                    btnBoard8No5.Enabled = True
                    btnBoard8No6.Enabled = True
                    btnBoard8No7.Enabled = True
                    btnBoard8No8.Enabled = True
                    btnBoard8No9.Enabled = True
                    penPoints = {New Point(222, 432), New Point(432, 432), New Point(432, 646), New Point(222, 646), _
                                 New Point(222, 432)}
                    'Sets the value of "penPoints" to the points (222, 432), (402, 432), (432, 646), (222, 646)
                    'And (222, 432)
                Case 9                             'If "boardNumber" is 9 then all the buttons for board 9 are enabled
                    btnBoard9No1.Enabled = True
                    btnBoard9No2.Enabled = True
                    btnBoard9No3.Enabled = True
                    btnBoard9No4.Enabled = True
                    btnBoard9No5.Enabled = True
                    btnBoard9No6.Enabled = True
                    btnBoard9No7.Enabled = True
                    btnBoard9No8.Enabled = True
                    btnBoard9No9.Enabled = True
                    penPoints = {New Point(432, 432), New Point(646, 432), New Point(646, 646), _
                                 New Point(432, 646), New Point(432, 432)}
                    'Sets the value of "penPoints" to the points (432, 432), (646, 432), (646, 646), (432, 646)
                    'And (432, 432)
            End Select
        End If
        
        image = BackgroundDraw(image, penPoints)
        pnlBackground.BackgroundImage = image
        
    End Sub
    
    Private Function BackgroundDraw(backgroundReturn, points) As Bitmap
        Dim colourDraw As New Color
        Dim type As String
        type = "Goth"
        If cmbThemeSelect.Text = type Then
            colourDraw = Color.White
        Else 
            colourDraw = Color.DimGray
        End If
        Dim penDraw As New Pen(colourDraw, 5)
        Dim background As Graphics
        background = Graphics.FromImage(backgroundReturn)
        
        background.DrawLines(penDraw, points)
        Return backgroundReturn
    End Function
    
    Private Function GamePlay(index1, index2, board, boardNumber)
        'Function "GamePlay", with "index1", "index2", "board" and "boardNumber" passed in
        Dim win As Boolean                               'Defines "win" as "boolean"
        Dim playerChange, total, state As Integer        'Defines these as "Integer"
        
        state = _playerGo                           'Sets the value of "state" to "_playerGo"
        If state = 1 Then                       'If "playerGo" is 1 then:
            board(index1, index2) = state       'The relevant array value is set to 1
            win = CheckNought(board)
            'Function "CheckNought" is executed, passing in "_playGame" and its results stored in "win"
            playerChange = 10                   'The value of "playerChange" is set to 10
        ElseIf state = 10 Then                  'Else if "_playerGo" is 10 then:
            board(index1, index2) = state       'The relevant array value is set to 10
            win = CheckCross(board)
            'Function "CheckCross" is executed, passing in "_playGame" and its results stored in "win"
            playerChange = 1                        'The value of "_playerChange" is set to 1
        Else                                        'Else executes procedure "ErrorHandler"
            ErrorHandler()
        End If
        
        If playerChange = 1 Then                           'If "playerChange" is 1 then "_playerGo" is set to 1
            _playerGo = 1
        ElseIf playerChange = 10 Then                      'Else if "_playerChange" is 10 then "_playerGo" is set to 10
            _playerGo = 10
        Else                                               'Else executes procedure "ErrorHandler"
            ErrorHandler()
        End If

        If win = True Then                                 'If "win" is true then 
            FindBoard(boardNumber, state)
            Exit Function
        End If

        For x = 0 To 2
            For y = 0 To 2
                If board(x, y) <> 0 Then
                    total += 1
                End If
            Next
        Next
        
        If total = 9 Then
            state = 0
            FindBoard(boardNumber, state)
            Exit Function
        End If
        
        If playerChange = 1 Then
            Return 10
        ElseIf playerChange = 10 Then
            Return 1
        Else
            ErrorHandler()
            Return Nothing
        End If
        
    End Function
    
    Private Sub FindBoard(board, state)
        Dim image As New Bitmap(192, 192)
        Dim userChoice, winner As String                 'Defines these as "strings"
        Dim win As Boolean
        Dim total As Integer
        win = False
        
        Select Case board
            Case 1
                btnBoard1No1.Hide()
                btnBoard1No2.Hide()
                btnBoard1No3.Hide()
                btnBoard1No4.Hide()
                btnBoard1No5.Hide()
                btnBoard1No6.Hide()
                btnBoard1No7.Hide()
                btnBoard1No8.Hide()
                btnBoard1No9.Hide()
                btnBoard1.Show()
                btnBoard1.Enabled = False
                image = DrawGraphics(2, state)
                btnBoard1.BackgroundImage = image
                _board(0, 0) = state
                _board1Win = True
            Case 2
                btnBoard2No1.Hide()
                btnBoard2No2.Hide()
                btnBoard2No3.Hide()
                btnBoard2No4.Hide()
                btnBoard2No5.Hide()
                btnBoard2No6.Hide()
                btnBoard2No7.Hide()
                btnBoard2No8.Hide()
                btnBoard2No9.Hide()
                btnBoard2.Show()
                btnBoard2.Enabled = False
                image = DrawGraphics(2, state)
                btnBoard2.BackgroundImage = image
                _board(0, 1) = state
                _board2Win = True
            Case 3
                btnBoard3No1.Hide()
                btnBoard3No2.Hide()
                btnBoard3No3.Hide()
                btnBoard3No4.Hide()
                btnBoard3No5.Hide()
                btnBoard3No6.Hide()
                btnBoard3No7.Hide()
                btnBoard3No8.Hide()
                btnBoard3No9.Hide()
                btnBoard3.Show()
                btnBoard3.Enabled = False
                image = DrawGraphics(2, state)
                btnBoard3.BackgroundImage = image
                _board(0, 2) = state
                _board3Win = True
            Case 4
                btnBoard4No1.Hide()
                btnBoard4No2.Hide()
                btnBoard4No3.Hide()
                btnBoard4No4.Hide()
                btnBoard4No5.Hide()
                btnBoard4No6.Hide()
                btnBoard4No7.Hide()
                btnBoard4No8.Hide()
                btnBoard4No9.Hide()
                btnBoard4.Show()
                btnBoard4.Enabled = False
                image = DrawGraphics(2, state)
                btnBoard4.BackgroundImage = image
                _board(1, 0) = state
                _board4Win = True
            Case 5
                btnBoard5No1.Hide()
                btnBoard5No2.Hide()
                btnBoard5No3.Hide()
                btnBoard5No4.Hide()
                btnBoard5No5.Hide()
                btnBoard5No6.Hide()
                btnBoard5No7.Hide()
                btnBoard5No8.Hide()
                btnBoard5No9.Hide()
                btnBoard5.Show()
                btnBoard5.Enabled = False
                image = DrawGraphics(2, state)
                btnBoard5.BackgroundImage = image
                _board(1, 1) = state
                _board5Win = True
            Case 6
                btnBoard6No1.Hide()
                btnBoard6No2.Hide()
                btnBoard6No3.Hide()
                btnBoard6No4.Hide()
                btnBoard6No5.Hide()
                btnBoard6No6.Hide()
                btnBoard6No7.Hide()
                btnBoard6No8.Hide()
                btnBoard6No9.Hide()
                btnBoard6.Show()
                btnBoard6.Enabled = False
                image = DrawGraphics(2, state)
                btnBoard6.BackgroundImage = image
                _board(1, 2) = state
                _board6Win = True
            Case 7
                btnBoard7No1.Hide()
                btnBoard7No2.Hide()
                btnBoard7No3.Hide()
                btnBoard7No4.Hide()
                btnBoard7No5.Hide()
                btnBoard7No6.Hide()
                btnBoard7No7.Hide()
                btnBoard7No8.Hide()
                btnBoard7No9.Hide()
                btnBoard7.Show()
                btnBoard7.Enabled = False
                image = DrawGraphics(2, state)
                btnBoard7.BackgroundImage = image
                _board(2, 0) = state
                _board7Win = True
            Case 8
                btnBoard8No1.Hide()
                btnBoard8No2.Hide()
                btnBoard8No3.Hide()
                btnBoard8No4.Hide()
                btnBoard8No5.Hide()
                btnBoard8No6.Hide()
                btnBoard8No7.Hide()
                btnBoard8No8.Hide()
                btnBoard8No9.Hide()
                btnBoard8.Show()
                btnBoard8.Enabled = False
                image = DrawGraphics(2, state)
                btnBoard8.BackgroundImage = image
                _board(2, 1) = state
                _board8Win = True
            Case 9
                btnBoard9No1.Hide()
                btnBoard9No2.Hide()
                btnBoard9No3.Hide()
                btnBoard9No4.Hide()
                btnBoard9No5.Hide()
                btnBoard9No6.Hide()
                btnBoard9No7.Hide()
                btnBoard9No8.Hide()
                btnBoard9No9.Hide()
                btnBoard9.Show()
                btnBoard9.Enabled = False
                image = DrawGraphics(2, state)
                btnBoard9.BackgroundImage = image
                _board(2, 2) = state
                _board9Win = True
        End Select
           
        If _playerGo = 1 Then
            win = CheckNought(_board)
        ElseIf _playerGo = 10 Then
            win = CheckCross(_board)
        Else
            ErrorHandler()
        End If
        
        For x = 0 To 2
            For y = 0 To 2
                If _board(x, y) <> 0 Then
                    total += 1
                End If
            Next
        Next
        
        If win = True Then
            winner = FindWinner()
            userChoice = MsgBox("The winner is: " & winner & vbCr & "Would you like to play again?",
                                MsgBoxStyle.YesNo, "End Of Game")
            'Generates a message box with the winner in it and stores their answer in "userChoice"
            RefreshForm(userChoice)    'Executes the procedure "RefreshForm", passing in "userChoice"
            Exit Sub                   'Exits the procedure
        ElseIf total = 9
            userChoice = MsgBox("That game was a draw. Would you like to play again?", MsgBoxStyle.YesNo,
                                "End Of Game")
            'Stores the player's option in "userChoice"
            RefreshForm(userChoice)    'Executes the procedure "RefreshForm", passing in "userChoice"
            Exit Sub                   'Exits the procedure
        End If
        
    End Sub
    
    Private Function FindWinner()                            'Function "FindWinner"
        Dim answer As String                                 'Defines "answer" as a string
        answer = ""                                          'Sets the contents of "answer" to ""
        If _playerGo = 1 Then                                'If "_playerGo" is 1 then:
            answer = tbxPlayer1.Text                         'Contents of "answer" is the text in "tbxPlayer1"
        ElseIf _playerGo = 10 Then                           'Else if "_playerGo" is 10 then:
            answer = tbxPlayer2.Text                         'Contents of "answer" is the text in "tbxPlayerTwo"
        Else                                                 'Else executes the procedure "ErrorHandler"
            ErrorHandler()
        End If

        Return answer                                        'Returns "answer"
    End Function
    
    Private Shared Function CheckNought(board)          'Function "CheckNought" with "board" passed in
        Dim result As Integer                           'Defines "result" as Integer
        result = 0                                      'Sets the value of "result" to 0

        For x = 0 To 2     'Loops through the columns in the array and adds the values to "result"
            For y = 0 To 2
                result += board(x, y)
            Next
            If result = 3 Then     'If "result" is 3 then returns true
                Return True
            End If
            result = 0             'Sets "result" to 0
        Next

        For y = 0 To 2     'Loops through the rows in the array and adds the values to "result"
            For x = 0 To 2
                result += board(x, y)
            Next
            If result = 3 Then     'If "result" is 3 then returns true
                Return True
            End If
            result = 0             'Sets "result" to 0
        Next

        For index = 0 To 2     'Loops through one diagonal in the array and adds the values to "result"
            result += board(index, index)
        Next
        If result = 3 Then     'If "result" is 3 then returns true
            Return True
        End If
        result = 0             'Sets "result" to 0

        For index = 0 To 2     'Loops through the other diagonal in the array and adds the values to "result"
            result += board(index, index)
        Next
        If result = 3 Then     'If "result" is 3 then returns true
            Return True
        End If

        Return False           'Returns false
    End Function

    Private Shared Function CheckCross(board)                   'Function "CheckCross" with "board" passed in
        Dim result As Integer                            'Defines "result" as Integer
        result = 0                                       'Sets the value of "result" to 0

        For x = 0 To 2     'Loops through the columns in the array and adds the values to "result"
            For y = 0 To 2
                result += board(x, y)
            Next
            If result = 30 Then     'If "result" is 30 then returns true
                Return True
            End If
            result = 0              'Sets "result" to 0
        Next

        For y = 0 To 2     'Loops through the rows in the array and adds the values to "result"
            For x = 0 To 2
                result += board(x, y)
            Next
            If result = 30 Then     'If "result" is 30 then returns true
                Return True
            End If
            result = 0              'Sets "result" to 0
        Next

        For index = 0 To 2     'Loops through one diagonal in the array and adds the values to "result"
            result += board(index, index)
        Next
        If result = 30 Then     'If "result" is 30 then returns true
            Return True
        End If
        result = 0              'Sets "result" to 0

        For index = 0 To 2     'Loops through the other diagonal in the array and adds the values to "result"
            result += board(index, index)
        Next
        If result = 30 Then     'If "result" is 30 then returns true
            Return True
        End If

        Return False            'Returns false
    End Function
    
End Class