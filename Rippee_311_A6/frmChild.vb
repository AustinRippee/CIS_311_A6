Imports System.Text

'------------------------------------------------------------
'-                File Name : frmChild.frm                     - 
'-                Part of Project: AssignX                  -
'------------------------------------------------------------
'-                Written By: Austin Rippee                     -
'-                Written On: March 20th, 2022         -
'------------------------------------------------------------
'- File Purpose:                                            -
'- The purpose of this file is to include a child of the
'- the parent function.
'------------------------------------------------------------
'- Program Purpose:                                         -
'-                                                          -
'- This program allows for a user to input a decimal/hex/binary
'- value and will convert and use math logic gates to compare
'- multiple values
'------------------------------------------------------------
'- Global Variable Dictionary (alphabetically):             -
'- blnIsTextBinaryVal1Active - Boolean to check if BinaryVal1 is active
'- blnIsTextBinaryVal2Active - Boolean to check if BinaryVal2 is active
'- blnIsTextDecimalVal1Active - Boolean to check if DecimalVal1 is active
'- blnIsTextDecimalVal2Active - Boolean to check if DecimalVal2 is active
'- blnIsTextHexVal1Active - Boolean to check if HexVal1 is active
'- blnIsTextHexVal2Active - Boolean to check if HexVal2 is active
'- blnIsValue1BinaryActive - Boolean to check if Value1Binary is active
'- blnIsValue1DecimalActive - Boolean to check if Value1Decimal is active
'- blnIsValue1HexActive - Boolean to check if Value1Hex is active
'- blnIsValue2BinaryActive - Boolean to check if Value2Binary is active
'- blnIsValue2DecimalActive - Boolean to check if Value2Decimal is active
'- blnIsValue2HexActive - Boolean to check if Value2Hex is active
'------------------------------------------------------------
Public Class frmChild

    'Creates global boolean values to check if the corresponding textbox is active for each button
    Public Shared blnIsTextBinaryVal1Active As Boolean = True
    Public Shared blnIsTextBinaryVal2Active As Boolean = True
    Public Shared blnIsTextDecimalVal1Active As Boolean = True
    Public Shared blnIsTextDecimalVal2Active As Boolean = True
    Public Shared blnIsTextHexVal1Active As Boolean = True
    Public Shared blnIsTextHexVal2Active As Boolean = True

    'Creates global boolean values to check if the corresponding textbox is active for each conversion
    Public Shared blnIsValue1BinaryActive As Boolean = True
    Public Shared blnIsValue1DecimalActive As Boolean = True
    Public Shared blnIsValue1HexActive As Boolean = True
    Public Shared blnIsValue2BinaryActive As Boolean = True
    Public Shared blnIsValue2DecimalActive As Boolean = True
    Public Shared blnIsValue2HexActive As Boolean = True

    '------------------------------------------------------------
    '-                Subprogram Name: frmChild_Load            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user loads the form                                        –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub frmChild_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Sets all textboxes to have 0 as the value
        txtVal1Binary.Text = "0"
        txtVal1Decimal.Text = "0"
        txtVal1Hex.Text = "0"
        txtVal2Binary.Text = "0"
        txtVal2Decimal.Text = "0"
        txtVal2Hex.Text = "0"
        txtResultBinary.Text = "0"
        txtResultDecimal.Text = "0"
        txtResultHex.Text = "0"
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: frmChild_FormClosing            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user closes the form.                                       –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the FormClosingEventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub frmChild_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'On a closing form, checks to see if the textboxes contain 0
        If txtVal1Binary.Text = "0" And txtVal1Decimal.Text = "0" And txtVal1Hex.Text = "0" And txtVal2Binary.Text = "0" And txtVal2Decimal.Text = "0" And txtVal2Hex.Text = "0" And txtResultBinary.Text = "0" And txtResultDecimal.Text = "0" And txtResultHex.Text = "0" Then
            'Closes out the program
            e.Cancel = False
        Else
            'If not, displays a messagebox that asks if the users wants to close out the child
            If MessageBox.Show("Close Dirty Child?", "Close Form", System.Windows.Forms.MessageBoxButtons.YesNo) = DialogResult.No Then
                'If no is pressed, the event is cancelled
                e.Cancel = True
            End If
        End If
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnClearVal1_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- clear value 1 button                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnClearVal1_Click(sender As Object, e As EventArgs) Handles btnClearVal1.Click
        'Sets value 1 textboxes to 0
        txtVal1Binary.Text = "0"
        txtVal1Decimal.Text = "0"
        txtVal1Hex.Text = "0"
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnClearVal2_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- clear value 2 button                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnClearVal2_Click(sender As Object, e As EventArgs) Handles btnClearVal2.Click
        'Sets value 2 textboxes to 0
        txtVal2Binary.Text = "0"
        txtVal2Decimal.Text = "0"
        txtVal2Hex.Text = "0"
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnClearResult_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- clear value 3 button                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnClearResult_Click(sender As Object, e As EventArgs) Handles btnClearResult.Click
        'Sets result textboxes to 0
        txtResultBinary.Text = "0"
        txtResultDecimal.Text = "0"
        txtResultHex.Text = "0"
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: txtVal1Binary_Focused            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user focuses
    '– on the txtVal1Binary textbox
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub txtVal1Binary_Focused(sender As Object, e As EventArgs) Handles txtVal1Binary.GotFocus
        'Enables/disables corresponding buttons to a binary textbox
        btnA.Enabled = False
        btnB.Enabled = False
        btnC.Enabled = False
        btnD.Enabled = False
        btnE.Enabled = False
        btnF.Enabled = False
        btn9.Enabled = False
        btn8.Enabled = False
        btn7.Enabled = False
        btn6.Enabled = False
        btn5.Enabled = False
        btn4.Enabled = False
        btn3.Enabled = False
        btn2.Enabled = False
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: txtVal2Binary_Focused            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user focuses
    '– on the txtVal2Binary textbox                                        –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub txtVal2Binary_Focused(sender As Object, e As EventArgs) Handles txtVal2Binary.GotFocus
        'Enables/disables corresponding buttons to a binary textbox
        btnA.Enabled = False
        btnB.Enabled = False
        btnC.Enabled = False
        btnD.Enabled = False
        btnE.Enabled = False
        btnF.Enabled = False
        btn9.Enabled = False
        btn8.Enabled = False
        btn7.Enabled = False
        btn6.Enabled = False
        btn5.Enabled = False
        btn4.Enabled = False
        btn3.Enabled = False
        btn2.Enabled = False
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: txtVal1Decimal_Focused            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user focuses
    '– on the txtVal1Decimal textbox                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub txtVal1Decimal_Focused(sender As Object, e As EventArgs) Handles txtVal1Decimal.GotFocus
        'Enables/disables corresponding buttons to a decimal textbox
        btnA.Enabled = False
        btnB.Enabled = False
        btnC.Enabled = False
        btnD.Enabled = False
        btnE.Enabled = False
        btnF.Enabled = False
        btn9.Enabled = True
        btn8.Enabled = True
        btn7.Enabled = True
        btn6.Enabled = True
        btn5.Enabled = True
        btn4.Enabled = True
        btn3.Enabled = True
        btn2.Enabled = True
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: txtVal2Decimal_Focused            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user focuses
    '– on the txtVal2Decimal textbox                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub txtVal2Decimal_Focused(sender As Object, e As EventArgs) Handles txtVal2Decimal.GotFocus
        'Enables/disables corresponding buttons to a decimal textbox
        btnA.Enabled = False
        btnB.Enabled = False
        btnC.Enabled = False
        btnD.Enabled = False
        btnE.Enabled = False
        btnF.Enabled = False
        btn9.Enabled = True
        btn8.Enabled = True
        btn7.Enabled = True
        btn6.Enabled = True
        btn5.Enabled = True
        btn4.Enabled = True
        btn3.Enabled = True
        btn2.Enabled = True
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: txtVal1Hex_Focused            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user focuses
    '– on the txtVal1Hex textbox                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub txtVal1Hex_Focused(sender As Object, e As EventArgs) Handles txtVal1Hex.GotFocus
        'Enables/disables corresponding buttons to a hex textbox
        btnA.Enabled = True
        btnB.Enabled = True
        btnC.Enabled = True
        btnD.Enabled = True
        btnE.Enabled = True
        btnF.Enabled = True
        btn9.Enabled = True
        btn8.Enabled = True
        btn7.Enabled = True
        btn6.Enabled = True
        btn5.Enabled = True
        btn4.Enabled = True
        btn3.Enabled = True
        btn2.Enabled = True
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: txtVal2Hex_Focused            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user focuses
    '– on the txtVal2Hex textbox                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub txtVal2Hex_Focused(sender As Object, e As EventArgs) Handles txtVal2Hex.GotFocus
        'Enables/disables corresponding buttons to a hex textbox
        btnA.Enabled = True
        btnB.Enabled = True
        btnC.Enabled = True
        btnD.Enabled = True
        btnE.Enabled = True
        btnF.Enabled = True
        btn9.Enabled = True
        btn8.Enabled = True
        btn7.Enabled = True
        btn6.Enabled = True
        btn5.Enabled = True
        btn4.Enabled = True
        btn3.Enabled = True
        btn2.Enabled = True
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnConvert_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- convert button.                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- intBinaryValue - Stores the textbox value of a binary number
    '- intDecimalValue - Stores the textbox value of a decimal number
    '- intHexValue - Stores the textbox value of a hex number
    '- intVal1Decimal - Stores conversion from binary to decimal
    '- intVal2Decimal - Stores conversion from binary to decimal
    '- intVal1Hex - Stores conversion from binary to hex
    '- intVal2Hex - Stores conversion from binary to hex
    '------------------------------------------------------------
    Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click
        If blnIsValue1BinaryActive Then 'Checks if isValue1BinaryActive is true
            'Sets the textbox value to intBinaryValue
            Dim intBinaryValue As Integer = txtVal1Binary.Text

            'Converts the Binary value to a Decimal value and sets it to a variable
            Dim intVal1Decimal As Integer = ConvertBinaryToDecimal(intBinaryValue)
            'Sets the value to the textbox
            txtVal1Decimal.Text = intVal1Decimal

            'Converts the value from binary to hex
            Dim intVal1Hex As String = ConvertBinaryToHex(intBinaryValue)
            'Sets the value to the textbox
            txtVal1Hex.Text = intVal1Hex
        ElseIf blnIsValue1DecimalActive Then
            'Sets the textbox value to intDecimalValue
            Dim intDecimalValue As Integer = txtVal1Decimal.Text

            'Converts the decimal value to binary
            txtVal1Binary.Text = ConvertDecimalToBinary(intDecimalValue)

            'Converts the decimal value to hex
            txtVal1Hex.Text = ConvertDecimalToHex(intDecimalValue)
        ElseIf blnIsValue1HexActive Then
            'Sets the textbox value to intHexValue
            Dim intHexValue As String = txtVal1Hex.Text

            'Converts the hex value to decimal
            txtVal1Decimal.Text = ConvertHexToDecimal(intHexValue)

            'Converts the hex value to binary
            txtVal1Binary.Text = ConvertHexToBinary(intHexValue)
        ElseIf blnIsValue2BinaryActive Then
            'Sets the textbox value to intVal2Binary
            Dim intBinaryValue As Integer = txtVal2Binary.Text

            'Converts the binary value to decimal
            Dim intVal2Decimal As Integer = ConvertBinaryToDecimal(intBinaryValue)
            'sets it to the textbox
            txtVal2Decimal.Text = intVal2Decimal

            'Converts the binary value to hex
            Dim intVal2Hex As String = ConvertBinaryToHex(intBinaryValue)
            'Sets the value to the textbox
            txtVal2Hex.Text = intVal2Hex
        ElseIf blnIsValue2DecimalActive Then
            'Sets the textbox value to intDecimalValue
            Dim intDecimalValue As Integer = txtVal2Decimal.Text

            'Converts the decimal value to binary
            txtVal2Binary.Text = ConvertDecimalToBinary(intDecimalValue)

            'Converts the decimal value to hex
            txtVal2Hex.Text = ConvertDecimalToHex(intDecimalValue)
        ElseIf blnIsTextHexVal2Active Then
            'Sets the textbox value to intHexValue
            Dim intHexValue As String = txtVal2Hex.Text

            'Converts the hex value to decimal
            txtVal2Decimal.Text = ConvertHexToDecimal(intHexValue)

            'Converts the hex value to binary
            txtVal2Binary.Text = ConvertHexToBinary(intHexValue)
        End If
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: txtBox_LostFocus            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user loses focus
    '– from a textbox.
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub txtBox_LostFocus(ByVal sender As Object, ByVal e As EventArgs) Handles txtVal1Binary.LostFocus, txtVal2Binary.LostFocus, txtVal1Decimal.LostFocus, txtVal2Decimal.LostFocus, txtVal1Hex.LostFocus, txtVal2Hex.LostFocus
        If (TryCast(sender, TextBox)).Name = "txtVal1Binary" Then 'Checks what the last focused textbox was and then sets the corresponding boolean to true
            blnIsValue1BinaryActive = True
            blnIsValue1DecimalActive = False
            blnIsValue1HexActive = False
            blnIsValue2BinaryActive = False
            blnIsValue2DecimalActive = False
            blnIsTextHexVal2Active = False
        ElseIf (TryCast(sender, TextBox)).Name = "txtVal1Decimal" Then 'Checks what the last focused textbox was and then sets the corresponding boolean to true
            blnIsValue1BinaryActive = False
            blnIsValue1DecimalActive = True
            blnIsValue1HexActive = False
            blnIsValue2BinaryActive = False
            blnIsValue2DecimalActive = False
            blnIsTextHexVal2Active = False
        ElseIf (TryCast(sender, TextBox)).Name = "txtVal1Hex" Then 'Checks what the last focused textbox was and then sets the corresponding boolean to true
            blnIsValue1BinaryActive = False
            blnIsValue1DecimalActive = False
            blnIsValue1HexActive = True
            blnIsValue2BinaryActive = False
            blnIsValue2DecimalActive = False
            blnIsTextHexVal2Active = False
        ElseIf (TryCast(sender, TextBox)).Name = "txtVal2Binary" Then 'Checks what the last focused textbox was and then sets the corresponding boolean to true
            blnIsValue1BinaryActive = False
            blnIsValue1DecimalActive = False
            blnIsValue1HexActive = False
            blnIsValue2BinaryActive = True
            blnIsValue2DecimalActive = False
            blnIsTextHexVal2Active = False
        ElseIf (TryCast(sender, TextBox)).Name = "txtVal2Decimal" Then 'Checks what the last focused textbox was and then sets the corresponding boolean to true
            blnIsValue1BinaryActive = False
            blnIsValue1DecimalActive = False
            blnIsValue1HexActive = False
            blnIsValue2BinaryActive = False
            blnIsValue2DecimalActive = True
            blnIsTextHexVal2Active = False
        ElseIf (TryCast(sender, TextBox)).Name = "txtVal2Hex" Then 'Checks what the last focused textbox was and then sets the corresponding boolean to true
            blnIsValue1BinaryActive = False
            blnIsValue1DecimalActive = False
            blnIsValue1HexActive = False
            blnIsValue2BinaryActive = False
            blnIsValue2DecimalActive = False
            blnIsTextHexVal2Active = True
        Else

        End If
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnClick            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks one
    '– of the buttons on the calculator and adds that value on
    '- the button to whatever textbox was last focused on
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- strTextValue - String to hold the literal value of the text of the button                                                   -
    '------------------------------------------------------------
    Private Sub btnClick(ByVal sender As Object, ByVal e As EventArgs) Handles btn0.Click, btn1.Click, btn2.Click, btn3.Click, btn4.Click, btn5.Click, btn6.Click, btn7.Click, btn8.Click, btn9.Click, btnA.Click, btnB.Click, btnC.Click, btnD.Click, btnE.Click, btnF.Click
        Dim strTextValue As String = (TryCast(sender, Button)).Text
        If blnIsTextBinaryVal1Active Then 'Checks if isTextBinaryVal1Active is active and adds the value on the button to the textbox
            txtVal1Binary.Text += strTextValue
        ElseIf blnIsTextBinaryVal2Active Then 'Checks if isTextBinaryVal2Active is active and adds the value on the button to the textbox
            txtVal2Binary.Text += strTextValue
        ElseIf blnIsTextDecimalVal1Active Then 'Checks if isTextDecimalVal1Active is active and adds the value on the button to the textbox
            txtVal1Decimal.Text += strTextValue
        ElseIf blnIsTextDecimalVal2Active Then 'Checks if isTextDecimalVal2Active is active and adds the value on the button to the textbox
            txtVal2Decimal.Text += strTextValue
        ElseIf blnIsTextHexVal1Active Then 'Checks if isTextHexVal1Active is active and adds the value on the button to the textbox
            txtVal1Hex.Text += strTextValue
        ElseIf blnIsTextHexVal2Active Then 'Checks if isTextHexVal2Active is active and adds the value on the button to the textbox
            txtVal2Hex.Text += strTextValue
        End If
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: TextBox_LostFocus            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user loses focus
    '– of a textbox and sets the boolean to true or false depending
    '- on if it was focused in on or not.
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub TextBox_LostFocus(ByVal sender As Object, ByVal e As EventArgs) Handles txtVal1Binary.LostFocus, txtVal2Binary.LostFocus, txtVal1Decimal.LostFocus, txtVal2Decimal.LostFocus, txtVal1Hex.LostFocus, txtVal2Hex.LostFocus
        If (TryCast(sender, TextBox)).Name = "txtVal1Binary" Then 'Checks if the last focused textbox was txtVal1Binary
            blnIsTextBinaryVal1Active = True
            blnIsTextBinaryVal2Active = False
            blnIsTextDecimalVal1Active = False
            blnIsTextDecimalVal2Active = False
            blnIsTextHexVal1Active = False
            blnIsTextHexVal2Active = False
        ElseIf (TryCast(sender, TextBox)).Name = "txtVal2Binary" Then 'Checks if the last focused textbox was txtVal2Binary
            blnIsTextBinaryVal1Active = False
            blnIsTextBinaryVal2Active = True
            blnIsTextDecimalVal1Active = False
            blnIsTextDecimalVal2Active = False
            blnIsTextHexVal1Active = False
            blnIsTextHexVal2Active = False
        ElseIf (TryCast(sender, TextBox)).Name = "txtVal1Decimal" Then 'Checks if the last focused textbox was txtVal1Decimal
            blnIsTextBinaryVal1Active = False
            blnIsTextBinaryVal2Active = False
            blnIsTextDecimalVal1Active = True
            blnIsTextDecimalVal2Active = False
            blnIsTextHexVal1Active = False
            blnIsTextHexVal2Active = False
        ElseIf (TryCast(sender, TextBox)).Name = "txtVal2Decimal" Then 'Checks if the last focused textbox was txtVal2Decimal
            blnIsTextBinaryVal1Active = False
            blnIsTextBinaryVal2Active = False
            blnIsTextDecimalVal1Active = False
            blnIsTextDecimalVal2Active = True
            blnIsTextHexVal1Active = False
            blnIsTextHexVal2Active = False
        ElseIf (TryCast(sender, TextBox)).Name = "txtVal1Hex" Then 'Checks if the last focused textbox was txtVal1Hex
            blnIsTextBinaryVal1Active = False
            blnIsTextBinaryVal2Active = False
            blnIsTextDecimalVal1Active = False
            blnIsTextDecimalVal2Active = False
            blnIsTextHexVal1Active = True
            blnIsTextHexVal2Active = False
        ElseIf (TryCast(sender, TextBox)).Name = "txtVal2Hex" Then 'Checks if the last focused textbox was txtVal2Hex
            blnIsTextBinaryVal1Active = False
            blnIsTextBinaryVal2Active = False
            blnIsTextDecimalVal1Active = False
            blnIsTextDecimalVal2Active = False
            blnIsTextHexVal1Active = False
            blnIsTextHexVal2Active = True
        Else

        End If
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnAnd_Click           -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- and button. It ands two values together.                                        –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- intResultDecimalAnd - Compares decimal values from the textbox
    '- strResultHexAnd - Holds the anded conversion from decimal to hex
    '- intResultHexAndConv - Holds the AND of the decimals
    '- intVal1Dec - holds conversion from hex to decimal for val1
    '- intVal2Dec - holds conversion from hex to decimal for val2
    '- strResultBinaryAnd - String that holds the conversion from decimal to binary
    '- strVal1Hex - holds textbox value of hex value 1
    '- strVal2Hex - holds textbox value of hex value 2
    '------------------------------------------------------------
    Private Sub btnAnd_Click(sender As Object, e As EventArgs) Handles btnAnd.Click

        'Ands the values of both decimal textboxes and prints it
        Dim intResultDecimalAnd As Integer = txtVal1Decimal.Text And txtVal2Decimal.Text
        txtResultDecimal.Text = intResultDecimalAnd

        'Converts the anded decimal value to binary and prints it
        Dim strResultBinaryAnd As String = ConvertDecimalToBinary(intResultDecimalAnd).ToString
        txtResultBinary.Text = strResultBinaryAnd

        'Sets the textbox values as strings in variables
        Dim strVal1Hex As String = txtVal1Hex.Text
        Dim strVal2Hex As String = txtVal2Hex.Text

        'Converts the textbox values and sets them to integer vartiables
        Dim intVal1Dec As Integer = ConvertHexToDecimal(strVal1Hex)
        Dim intVal2Dec As Integer = ConvertHexToDecimal(strVal2Hex)

        'Ands the results, converts the decimal to hex and prints it
        Dim intResultHexAndConv As Integer = strVal1Hex And strVal2Hex
        Dim strResultHexAnd As String = ConvertDecimalToHex(intResultHexAndConv)
        txtResultHex.Text = strResultHexAnd
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnOr_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- or button. It ors two values together.                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- intResultDecimalOr - Compares decimal values from the textbox
    '- strResultHexOr - Holds the or'd conversion from decimal to hex
    '- intResultHexOrConv - Holds the OR of the decimals
    '- intVal1Dec - holds conversion from hex to decimal for val1
    '- intVal2Dec - holds conversion from hex to decimal for val2
    '- strResultBinaryAnd - String that holds the conversion from decimal to binary
    '- strVal1Hex - holds textbox value of hex value 1
    '- strVal2Hex - holds textbox value of hex value 2
    '------------------------------------------------------------
    Private Sub btnOr_Click(sender As Object, e As EventArgs) Handles btnOr.Click
        'Ors the values of both decimal textboxes and prints it
        Dim intResultDecimalOr As Integer
        intResultDecimalOr = txtVal1Decimal.Text Or txtVal2Decimal.Text
        txtResultDecimal.Text = intResultDecimalOr

        'Converts the or'd decimal value to binary and prints it
        Dim intResultBinaryOr As String = ConvertDecimalToBinary(intResultDecimalOr).ToString
        txtResultBinary.Text = intResultBinaryOr

        'Sets the textbox values as strings in variables
        Dim strVal1Hex As String = txtVal1Hex.Text
        Dim strVal2Hex As String = txtVal2Hex.Text

        'Converts the textbox values and sets them to integer vartiables
        Dim intVal1Dec As Integer = ConvertHexToDecimal(strVal1Hex)
        Dim intVal2Dec As Integer = ConvertHexToDecimal(strVal2Hex)

        'Ors the results, converts the decimal to hex and prints it
        Dim intResultHexOrConv As Integer = intVal1Dec Or intVal2Dec
        Dim strResultHexOr As String = ConvertDecimalToHex(intResultHexOrConv)
        txtResultHex.Text = strResultHexOr
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnXor_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- xor button. It ors two values together.                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- Local Variable Dictionary (alphabetically):              -
    '- intResultDecimalXor - Compares decimal values from the textbox
    '- strResultHexXor - Holds the xor'd conversion from decimal to hex
    '- intResultHexXorConv - Holds the XOR of the decimals
    '- intVal1Dec - holds conversion from hex to decimal for val1
    '- intVal2Dec - holds conversion from hex to decimal for val2
    '- strResultBinaryXor - String that holds the conversion from decimal to binary
    '- strVal1Hex - holds textbox value of hex value 1
    '- strVal2Hex - holds textbox value of hex value 2
    '------------------------------------------------------------
    Private Sub btnXor_Click(sender As Object, e As EventArgs) Handles btnXor.Click
        'Xors the values of both decimal textboxes and prints it
        Dim intResultDecimalXor As Integer
        intResultDecimalXor = txtVal1Decimal.Text Xor txtVal2Decimal.Text
        txtResultDecimal.Text = intResultDecimalXor

        'Converts the xor'd decimal value to binary and prints it
        Dim strResultBinaryXor As String = ConvertDecimalToBinary(intResultDecimalXor).ToString
        txtResultBinary.Text = strResultBinaryXor

        'Sets the textbox values as strings in variables
        Dim strVal1Hex As String = txtVal1Hex.Text
        Dim strVal2Hex As String = txtVal2Hex.Text

        'Converts the textbox values and sets them to integer vartiables
        Dim intVal1Dec As Integer = ConvertHexToDecimal(strVal1Hex)
        Dim intVal2Dec As Integer = ConvertHexToDecimal(strVal2Hex)

        'Xors the results, converts the decimal to hex and prints it
        Dim intResultHexXorConv As Integer = intVal1Dec Xor intVal2Dec
        Dim strResultHexXor As String = ConvertDecimalToHex(intResultHexXorConv)
        txtResultHex.Text = strResultHexXor
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnNotVal1_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- not value 1 button. It not values the first value.                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- intResultDecimalNotVal1 - Holds the not value of the textbox
    '- strResultHexNotVal1 - Holds the string convserion from decimal to hex
    '- intResultHexNotVal1Conv - Holds the value as an integer
    '- intVal1Dec - holds not'd conversion from hex to decimal
    '- strVal1Hex - holds hex value from textbox
    '------------------------------------------------------------
    Private Sub btnNotVal1_Click(sender As Object, e As EventArgs) Handles btnNotVal1.Click
        'Not values the 1st value and prints it
        Dim intResultDecimalNotVal1 As Integer = Not txtVal1Decimal.Text
        txtResultDecimal.Text = intResultDecimalNotVal1

        'Takes the hex values, nots the conversion to decimal and converts it back to hex
        Dim strVal1Hex As String = txtVal1Hex.Text
        Dim intVal1Dec As Integer = Not ConvertHexToDecimal(strVal1Hex)
        Dim intResultHexNotVal1Conv As Integer = intVal1Dec
        Dim strResultHexNotVal1 As String = ConvertDecimalToHex(intResultHexNotVal1Conv)
        txtResultHex.Text = strResultHexNotVal1

        '==========================================================================
        ' The code below confuses me very much. I am attempting to convert the hex
        ' value into a binary value like I did above in the normal conversion through
        ' the convert button but it is not liking it no matter what way I do it.
        '==========================================================================

        'Takes the Hex value, converts it to binary and prints it
        'Dim intResultBinaryNotVal1 As Integer = ConvertHexToBinary(strResultHexNotVal1)
        'txtResultBinary.Text = intResultBinaryNotVal1
    End Sub

    '------------------------------------------------------------
    '-                Function Name: ConvertBinaryToDecimal            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This function when a user attempts to convert a binary
    '– to decimal number
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- num - Integer holding binary value
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- intValDecimal - holds decimal value                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Integer – decimal value           -
    '------------------------------------------------------------
    Function ConvertBinaryToDecimal(ByVal num As Integer) As Integer
        Dim intValDecimal As Integer = Convert.ToInt64(num, 2) 'Converts to a base 2 number
        Return intValDecimal
    End Function

    '------------------------------------------------------------
    '-                Function Name: ConvertBinaryToHex            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This function is called when a user attempts to convert
    '– a binary to hex value.
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- num - integer holding binary value
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- intValDecimal - holds converted value (Decimal)
    '- strValHex - Holds hex version of string value
    '- strValHexTemp - holds string of decimal value
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String – returns hex value            -
    '------------------------------------------------------------
    Function ConvertBinaryToHex(ByVal num As Integer) As String
        Dim intValDecimal As Integer = Convert.ToInt64(num, 2) 'Converts to a base 2 number
        Dim strValHexTemp As String = intValDecimal.ToString 'Converts to a string
        Dim strValHex As String = Hex(strValHexTemp) 'Uses the hex function
        Return strValHex
    End Function

    '------------------------------------------------------------
    '-                Function Name: ConvertDecimalToBinary            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This function is called whenever a user attempts to convert
    '– a decimal value to a binary value.
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- num - integer that holds decimal value
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- arrValue - Holds a char array where each character holds one element of the array
    '- intBinaryValue - holds converted decimal value
    '- intFinalBinaryValue - Adds both strRestZeroes & intBinaryValue
    '- intRestOfString - Finds how many values are missing from a 32 value string
    '- strConversionResult - Holds converted result of the string
    '- strRestZeroes - Holds literal values of how many zeroes need to be added
    '- strReturnValue - Holds the entire string result as it gets built up through the loop
    '- strValue - Holds the entire binary string as it gets iterated through
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String – returns binary value
    '------------------------------------------------------------
    Function ConvertDecimalToBinary(ByVal num As Integer) As String
        Dim intBinaryValue As Integer

        intBinaryValue = Convert.ToString(num, 2) 'Converts to a base 2 number

        'Initlializes variables for the binary conversion
        Dim strConversionResult As String, strValue As String, strReturnValue As String = ""
        Dim i As Integer = 0
        Dim arrValue As Char()

        'Finds how many values are missing from a 32 value string
        Dim intRestOfString As Integer = 32 - intBinaryValue.ToString.Length()
        Dim strRestZeroes As String = StrDup(intRestOfString, "0")

        'Adds the rest of the 0s to the binary string to make it 32 values total
        Dim intFinalBinaryValue As String = strRestZeroes & intBinaryValue

        'Reverses the string in order to separate it properly
        strConversionResult = StrReverse(intFinalBinaryValue)

        'Creates a char array in which each char is its own array element
        arrValue = strConversionResult.Take(strConversionResult.Length).ToArray
        'Iterates through every element
        For Each strValue In arrValue
            'Skips the first element and adds the space when the returned value is 0
            'to determine that it is the 4th character
            If (i <> 0) And (i Mod 4 = 0) Then
                'Adds the space
                strReturnValue = strReturnValue + " "
            End If
            'Adds the value to the returned value
            strReturnValue = strReturnValue + strValue
            'incrememets the value counter by 1
            i = i + 1
        Next
        'Reverses the output again to put it in the correct order
        strReturnValue = StrReverse(strReturnValue)
        Return strReturnValue
    End Function

    '------------------------------------------------------------
    '-                Function Name: ConvertDecimalToHex            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This function is called when a user attempts to convert a
    '– decimal value to a hex value
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- num - integer that holds decimal value
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- intValHexTemp - holds the decimal value as a string
    '- strValHex - holds the hex version of the decimal value
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - returns hex value            -
    '------------------------------------------------------------
    Function ConvertDecimalToHex(ByVal num As Integer) As String
        Dim intValHexTemp As Integer = num.ToString 'Takes the decimal value and puts it in a temp variable
        Dim strValHex As String = Hex(intValHexTemp) 'Takes the Hex function and puts it into the string
        Return strValHex
    End Function

    '------------------------------------------------------------
    '-                Function Name: ConvertHexToBinary            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This function is called whenever a user attempts to convert
    '– a hex value to a binary value.
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- num - integer that holds decimal value
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- arrValue - Holds a char array where each character holds one element of the array
    '- intBinaryValue - holds converted decimal value
    '- intDecimalValue - holds the converted hex value
    '- intFinalBinaryValue - Adds both strRestZeroes & intBinaryValue
    '- intRestOfString - Finds how many values are missing from a 32 value string
    '- strConversionResult - Holds converted result of the string
    '- strRestZeroes - Holds literal values of how many zeroes need to be added
    '- strReturnValue - Holds the entire string result as it gets built up through the loop
    '- strValue - Holds the entire binary string as it gets iterated through
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String – returns hex value            -
    '------------------------------------------------------------
    Function ConvertHexToBinary(ByVal num As String) As String
        Dim strValDecimal As String = Convert.ToInt64(num, 16) 'Converts to a base 16 number

        'Initlializes variables
        Dim intDecimalValue As Integer
        Dim intBinaryValue As Integer

        intDecimalValue = strValDecimal
        intBinaryValue = Convert.ToString(intDecimalValue, 2) 'Converts to a base 2 number

        'Initlializes variables for the binary conversion
        Dim strConversionResult As String, strValue As String, strReturnValue As String = ""
        Dim i As Integer = 0
        Dim arrValue As Char()

        'Finds how many values are missing from a 32 value string
        Dim intRestOfString As Integer = 32 - intBinaryValue.ToString.Length()
        Dim strRestZeroes As String = StrDup(intRestOfString, "0")

        'Adds the rest of the 0s to the binary string to make it 32 values total
        Dim intFinalBinaryValue As String = strRestZeroes & intBinaryValue

        'Reverses the string in order to separate it properly
        strConversionResult = StrReverse(intFinalBinaryValue)

        'Creates a char array in which each char is its own array element
        arrValue = strConversionResult.Take(strConversionResult.Length).ToArray
        'Iterates through every element
        For Each strValue In arrValue
            'Skips the first element and adds the space when the returned value is 0
            'to determine that it is the 4th character
            If (i <> 0) And (i Mod 4 = 0) Then
                'Adds the space
                strReturnValue = strReturnValue + " "
            End If
            'Adds the value to the returned value
            strReturnValue = strReturnValue + strValue
            'incrememets the value counter by 1
            i = i + 1
        Next
        'Reverses the output again to put it in the correct order
        strReturnValue = StrReverse(strReturnValue)
        Return strReturnValue
    End Function

    '------------------------------------------------------------
    '-                Function Name: ConvertHexToDecimal            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This function is called whenever a user attempts to convert
    '– a hex value to a decimal value
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- num - integer that holds decimal value
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- intValDecimal - holds the decimal conversion from hex                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Integer – returns decimal value            -
    '------------------------------------------------------------
    Function ConvertHexToDecimal(ByVal num As String) As Integer
        Dim intValDecimal As Integer = Convert.ToInt64(num, 16) 'Converts to a base 16 number
        Return intValDecimal
    End Function

End Class
