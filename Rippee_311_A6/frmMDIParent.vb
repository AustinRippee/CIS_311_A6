'------------------------------------------------------------
'-                File Name : frmMDIParent.frm                     - 
'-                Part of Project: Parent Form                  -
'------------------------------------------------------------
'-                Written By: Austin Rippee                     -
'-                Written On: March 20th, 2022         -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the main form of the application                                      - 
'------------------------------------------------------------
'- Program Purpose:                                         -
'-                                                          -
'- This program creates a parent form in which the user may create
'- multiple child forms
'------------------------------------------------------------
'- Global Variable Dictionary (alphabetically):             -
'- intChildCount - Keeps track of how many child forms have been created
'------------------------------------------------------------
Public Class frmMDIParent

    Dim intChildCount As Integer = 0 'Keeps track of how many child forms have been created

    '------------------------------------------------------------
    '-                Subprogram Name: mnuNew_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- menu new button.                                        –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- aNewChildForm - An instance of the frmChild()                                                   -
    '------------------------------------------------------------
    Private Sub mnuNew_Click(sender As Object, e As EventArgs) Handles mnuNew.Click
        'Creates a new instance of frmChild
        Dim aNewChildForm As frmChild = New frmChild()
        aNewChildForm.MdiParent = Me
        'Increments the child counter
        intChildCount += 1
        'Changes the titles and shows that one
        aNewChildForm.Text = "Calc " & intChildCount
        aNewChildForm.Show()
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: mnuFileExit_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- menu exit button.                                       –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuExit.Click
        'Closes the current form
        Me.Close()
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: frmMDIParent_FormClosing            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user exits the
    '– program
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the FormClosingEventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub frmMDIParent_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '========================================================
        '
        ' I attempted many different ways to create a proper way
        ' of closing but wasn't able to
        '
        '========================================================
        'If frmChild.txtVal1Binary.Text = "0" And frmChild.txtVal1Decimal.Text = "0" And frmChild.txtVal1Hex.Text = "0" And frmChild.txtVal2Binary.Text = "0" And frmChild.txtVal2Decimal.Text = "0" And frmChild.txtVal2Hex.Text = "0" And frmChild.txtResultBinary.Text = "0" And frmChild.txtResultDecimal.Text = "0" And frmChild.txtResultHex.Text = "0" Then
        '    frmChanged = False
        'Else
        '    frmChanged = True
        'End If


        'For Each child As Form In Me.MdiParent.MdiChildren
        '    If frmChanged = True Then
        '        e.Cancel = True
        '        MsgBox("Cancel = TRUE")
        '    Else
        '        'child.Close()
        '        e.Cancel = True
        '        MsgBox("Cancel = FALSE")
        '    End If
        'Next

        'If frmChild.txtVal1Binary.Text = "0" And frmChild.txtVal1Decimal.Text = "0" And frmChild.txtVal1Hex.Text = "0" And frmChild.txtVal2Binary.Text = "0" And frmChild.txtVal2Decimal.Text = "0" And frmChild.txtVal2Hex.Text = "0" And frmChild.txtResultBinary.Text = "0" And frmChild.txtResultDecimal.Text = "0" And frmChild.txtResultHex.Text = "0" Then
        '    ActiveMdiChild.Close()
        'End If
        'If Application.OpenForms().OfType(Of frmChild).Any Then
        '    If frmChild.txtVal1Binary.Text = "0" And frmChild.txtVal1Decimal.Text = "0" And frmChild.txtVal1Hex.Text = "0" And frmChild.txtVal2Binary.Text = "0" And frmChild.txtVal2Decimal.Text = "0" And frmChild.txtVal2Hex.Text = "0" And frmChild.txtResultBinary.Text = "0" And frmChild.txtResultDecimal.Text = "0" And frmChild.txtResultHex.Text = "0" Then
        '        frmChild.Close()
        '    Else
        '        MsgBox("Error")
        '        'If MessageBox.Show("Are you sure you want to quit?", "Calculator", MessageBoxButtons.YesNo) = DialogResult.Yes Then
        '        '    e.Cancel = False
        '        'Else
        '        '    e.Cancel = True
        '    End If
        'End If
        'Else
        'If MessageBox.Show("Are you sure you want to quit?", "Calculator", MessageBoxButtons.YesNo) = DialogResult.Yes Then
        'e.Cancel = False
        'Else
        '    e.Cancel = True
        'End If
        'End If
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: mnuAbout_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- menu about button.                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub mnuAbout_Click(sender As Object, e As EventArgs) Handles mnuAbout.Click
        'Shows the abount form
        frmAbout.ShowDialog()
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: mnuCascade_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- menu cascade button.                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub mnuCascade_Click(sender As Object, e As EventArgs) Handles mnuCascade.Click
        'Changes to a cascade layout
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: mnuTileHorizontal_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- menu tile horizontal button.                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub mnuTileHorizontal_Click(sender As Object, e As EventArgs) Handles mnuTileHorizontal.Click
        'Changes to a horizontal tile layout
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: mnuTileVertical_Click            -
    '------------------------------------------------------------
    '-                Written By: Austin Rippee                     -
    '-                Written On: March 20th, 2022         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called whenever the user clicks the   -
    '- menu tile vertical button.                                         –
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub mnuTileVertical_Click(sender As Object, e As EventArgs) Handles mnuTileVertical.Click
        'Changes to a vertical tile layout
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub
End Class