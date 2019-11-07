Imports System.Data.SqlClient
Imports System.ComponentModel



'// How to use this program
'// 1. Change ConnectionString
'// 2. Change SQL Command for you need monitor Ex. I want to see  QCFirstLotMode, QCFinishLotMode,LotStartTime,LotEndTime
'// 3. Click button 1 for active SqlDepencyCon
'// 4. Edit data on database for checking event 
'// 5. When you edited data on database it will call dependency1_OnChange

'//  THARIN YENPAIROJ    07/11/2019


Public Class Form1
    Private connection As SqlConnection = Nothing
    Private command As SqlCommand = Nothing
    Private dataToWatch As DataSet = Nothing
    Private Const tableName1 As String = "DBData" '"DailyProcessOperationRate"

    Private Sub SqlDependencyFunction(ByVal lotNo As String)
        'Dim lotno As String = tbLotnO.Text
        'changeCount1 = 0
        'Me.Label2.Text = String.Format(statusMessage, changeCount)

        ' Remove any existing dependency connection, then create a new one.
        SqlDependency.Stop(My.Settings.SqlDepencyCon)
        SqlDependency.Start(My.Settings.SqlDepencyCon)

        If connection Is Nothing Then
            connection = New SqlConnection(My.Settings.SqlDepencyCon)
        End If

        'If command1 Is Nothing Then
        ' GetSQL is a local procedure that returns
        ' a paramaterized SQL string. You might want
        ' to use a stored procedure in your application.
        Command = New SqlCommand(GetSQLTest(lotNo), connection)
        'End If


        If dataToWatch Is Nothing Then
            dataToWatch = New DataSet()
        End If

        GetData1()


    End Sub

    Private Function GetSQLTest(ByVal LotNo As String) As String
        '****** WARNING *****
        'http://www.codeproject.com/Articles/12335/Using-SqlDependency-for-data-change-events
        'If the query is not correct, an event will immediately be sent with:
        'EX: if statement is not include "dbo" the error will be notified
        'Return "SELECT ID,Name,Class,HP,SP,LVL,EXP,EXP_MAX FROM dbo.Charactor"
        'Return "SELECT RohmDate,ProcessName,OPRate,LoadTime FROM dbo.DailyProcessOperationRate WHERE ProcessName = 'DB'"
        If LotNo Is Nothing Then
            LotNo = ""
        End If
        Return "SELECT QCFirstLotMode, QCFinishLotMode,LotStartTime,LotEndTime FROM dbo.DBData WHERE Lotno = '" & LotNo & "' and MCNo = '" & tbMcno.Text & "'"
    End Function

    Private Sub GetData1()
        ' Empty the dataset so that there is only
        ' one batch worth of data displayed.
        dataToWatch.Clear()

        ' Make sure the command object does not already have
        ' a notification object associated with it.
        Command.Notification = Nothing

        ' Create and bind the SqlDependency object
        ' to the command object.
        Dim dependency1 As New SqlDependency(Command)
        AddHandler dependency1.OnChange, AddressOf dependency1_OnChange

        Using adapter As New SqlDataAdapter(Command)
            adapter.Fill(dataToWatch, tableName1)

            Me.DataGridView1.DataSource = dataToWatch
            Me.DataGridView1.DataMember = tableName1

        End Using

    End Sub
    Private Sub dependency1_OnChange(ByVal sender As Object, ByVal e As SqlNotificationEventArgs)

        ' This event will occur on a thread pool thread.
        ' It is illegal to update the UI from a worker thread
        ' The following code checks to see if it is safe
        ' update the UI.
        Dim i As ISynchronizeInvoke = CType(Me, ISynchronizeInvoke)

        ' If InvokeRequired returns True, the code
        ' is executing on a worker thread.
        If i.InvokeRequired Then
            ' Create a delegate to perform the thread switch
            Dim tempDelegate As New OnChangeEventHandler(AddressOf dependency1_OnChange)

            Dim args() As Object = {sender, e}

            ' Marshal the data from the worker thread
            ' to the UI thread.
            i.BeginInvoke(tempDelegate, args)

            Return
        End If

        ' Remove the handler since it's only good
        ' for a single notification
        Dim dependency1 As SqlDependency = CType(sender, SqlDependency)

        RemoveHandler dependency1.OnChange, AddressOf dependency1_OnChange

        ' At this point, the code is executing on the
        ' UI thread, so it is safe to update the UI.
        'changeCount1 += 1
        'Me.Label2.Text = String.Format(statusMessage, changeCount1)

        ' Add information from the event arguments to the list box
        ' for debugging purposes only.
        'With Me.ListBox2.Items
        '    .Clear()
        '    .Add("Info:   " & e.Info.ToString())
        '    .Add("Source: " & e.Source.ToString())
        '    .Add("Type:   " & e.Type.ToString())
        'End With


        If e.Type <> SqlNotificationType.Change Then
            Exit Sub
        End If



        ' Reload the dataset that's bound to the grid.
        GetData1()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SqlDependencyFunction(tbLotNo.Text)
    End Sub
End Class
