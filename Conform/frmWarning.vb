Public NotInheritable Class frmWarning
    Private Sub frmWarning_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        chkWarnAgree.Checked = False 'If we are displaying then ensure that the check box is unchecked and that the OK button is greyed out
        btnWarnOK.Enabled = False
    End Sub

    Private Sub chkWarnAgree_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkWarnAgree.CheckedChanged
        If chkWarnAgree.Checked = True Then ' User has agreed so enable the OK button and disable Exit
            btnWarnOK.Enabled = True
            btnWarnExit.Enabled = False
            g_Settings.WarningMessage = True ' Save the setting
        Else
            btnWarnOK.Enabled = False ' User has not agreed so disable the OK button and enable Exit
            btnWarnExit.Enabled = True
            g_Settings.WarningMessage = False ' Save the setting
        End If
    End Sub

    Private Sub btnWarnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWarnOK.Click

        Me.Close()
    End Sub

    Private Sub btnWarnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWarnExit.Click
        End ' End the program if the user clicks the exit button
    End Sub
End Class