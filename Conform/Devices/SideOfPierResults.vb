Imports ASCOM
Imports ASCOM.Interface

Friend Class SideOfPierResults
    Private m_SOP, m_DSOP As PierSide
    Sub New()
        m_SOP = PierSide.pierUnknown
        m_DSOP = PierSide.pierUnknown
    End Sub
    Friend Property SideOfPier() As PierSide
        Get
            Return m_SOP
        End Get
        Set(ByVal value As PierSide)
            m_SOP = value
        End Set
    End Property
    Friend Property DestinationSideOfPier() As PierSide
        Get
            Return m_DSOP
        End Get
        Set(ByVal value As PierSide)
            m_DSOP = value
        End Set
    End Property

End Class
