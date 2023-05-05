Public Class SapMmMdExcelAddin

    Private Sub SapMmMdExcelAddin_Startup() Handles Me.Startup
        log4net.Config.XmlConfigurator.Configure()
    End Sub

    Private Sub SapMmMdExcelAddin_Shutdown() Handles Me.Shutdown

    End Sub

End Class
