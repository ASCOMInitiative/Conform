Friend Class Stat
    WriteOnly Property Test() As String
        Set(ByVal value As String)
            FrmConformMain.staTest.Text = value
            Application.DoEvents()
        End Set
    End Property
    WriteOnly Property Action() As String
        Set(ByVal value As String)
            FrmConformMain.staAction.Text = value
            Application.DoEvents()
        End Set
    End Property
    WriteOnly Property Status() As String
        Set(ByVal value As String)
            FrmConformMain.staStatus.Text = value
            Application.DoEvents()
        End Set
    End Property
    Sub Clear()
        Me.Test = ""
        Me.Action = ""
        Me.Status = ""
    End Sub
End Class
