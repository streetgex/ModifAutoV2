Imports System.Management.Automation
Imports System.Management.Automation.Host
Imports System.Management.Automation.Runspaces



Public Class PwShell

    Dim runSpace As Runspace
    Dim pipeLine As Pipeline
    Sub New()
        runSpace = RunspaceFactory.CreateRunspace()
        runSpace.Open()

        'exec("Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin")
        exec("add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010")
    End Sub
    Function exec(ByVal cmd As String) 'As Collection(Of PSObject)
        pipeLine = runSpace.CreatePipeline()
        pipeLine.Commands.AddScript(cmd)
        Return pipeLine.Invoke

    End Function
    Sub close()
        runSpace.Close()
    End Sub

End Class