Public NotInheritable Class frmSwitchWarning

    Private Sub btnWarnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWarnOK.Click
        g_Settings.SwitchWarningMessage = True ' Set a flag to show that we have displayed the warning message
        Me.Close() ' Close the form
    End Sub

End Class