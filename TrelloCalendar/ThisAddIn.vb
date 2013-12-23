Public Class ThisAddIn

    Private ribbon As Ribbon1

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ribbon.startup()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        ribbon = New Ribbon1()
        Return (New Ribbon1())
    End Function





End Class
