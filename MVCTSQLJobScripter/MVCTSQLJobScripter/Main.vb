Imports Microsoft.SqlServer.Management.Smo
Imports Microsoft.SqlServer.Management.Sdk.Sfc

Module Main

    Sub Main()

        Dim smoConnectionSettings As New Microsoft.SqlServer.Management.Common.ServerConnection
        Dim smoServer As Microsoft.SqlServer.Management.Smo.Server
        Dim sServerNane As String = ""
        Dim JobList As Microsoft.SqlServer.Management.Smo.Agent.JobCollection
        Dim Job As Microsoft.SqlServer.Management.Smo.Agent.Job
        Dim a As Int32 = 0
        Dim sJobScript As String = ""
        Dim sFileName As String = ""

        If My.Application.CommandLineArgs.Count > 0 Then
            For a = 0 To My.Application.CommandLineArgs.Count - 1
                If My.Application.CommandLineArgs(a).ToLower.StartsWith("/server:") Then
                    sServerNane = Trim(Right(My.Application.CommandLineArgs(a), Len(My.Application.CommandLineArgs(a)) - 8))
                End If
                If My.Application.CommandLineArgs(a).ToLower.StartsWith("/outfile:") Then
                    sFileName = Trim(Right(My.Application.CommandLineArgs(a), Len(My.Application.CommandLineArgs(a)) - 9))
                End If
            Next
        End If

        If sServerNane <> "" And sFileName <> "" Then
            smoConnectionSettings.ApplicationName = "MVCT Job Scripter"
            smoConnectionSettings.ServerInstance = sServerNane

            smoServer = New Microsoft.SqlServer.Management.Smo.Server(smoConnectionSettings)
            JobList = smoServer.JobServer.Jobs
            Dim sqlOptions As New Microsoft.SqlServer.Management.Smo.ScriptingOptions With
            {
                .AllowSystemObjects = True,
                .AnsiFile = True,
                .AnsiPadding = True,
                .FileName = sFileName,
                .AppendToFile = False
            }
            For Each Job In JobList
                Job.Script(sqlOptions)
                sqlOptions.AppendToFile = True
            Next
        Else
            Console.WriteLine("Useage is:  MVCTSQLJobScripter /server:[ServerName] /outfile:[output file name]")
        End If

    End Sub

End Module
