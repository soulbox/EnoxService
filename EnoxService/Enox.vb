Imports System.ServiceProcess

Public Class Enox
    Dim tr As Threading.Thread
    Dim tr1 As Threading.Thread
    Dim tr2 As Threading.Thread
    Dim tr3 As Threading.Thread
    Dim tr4 As Threading.Thread
    Dim tr5 As Threading.Thread
    Dim tr6 As Threading.Thread
    Dim asd As New KütüpHane
    Dim MyLib As New NewLibrary
    Public Sub New()
        MyBase.New()

        ' This call is required by the designer.
        InitializeComponent()
        If Not System.Diagnostics.EventLog.SourceExists("Enox") Then
            System.Diagnostics.EventLog.CreateEventSource("Enox", "EnoxLog")
        End If
        ' Add any initialization after the InitializeComponent() call.
        EventLog1.Source = "Enox"
        EventLog1.Log = "EnoxLog"
        If Not System.Diagnostics.EventLog.SourceExists("EnoxSqlServis") Then
            System.Diagnostics.EventLog.CreateEventSource("EnoxSqlServis", "EnoxSqlServisLog")
        End If
        EventLog2.Source = "EnoxSqlServis"
        EventLog2.Log = "EnoxSqlServisLog"
    End Sub



    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Timer1.Start()
        Timer3.Start()
        EventLog1.WriteEntry("timerlala beraber başladı")

    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        EventLog1.WriteEntry("servis durdu")
    End Sub
    Private Sub Timer1_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles Timer1.Elapsed
        ' asd = New KütüpHane
        'asd = New KütüpHane
        'tr = New Threading.Thread(New Threading.ThreadStart(Function() asd.GetLogs(EventLog1, e.SignalTime)))
        'tr.SetApartmentState(Threading.ApartmentState.STA)
        'If tr.IsAlive = False Then tr.Start()
        'tr3 = New Threading.Thread(New Threading.ThreadStart(Function() asd.PressureGetAndSet(EventLog1, e.SignalTime)))
        'tr3.SetApartmentState(Threading.ApartmentState.STA)
        'If tr3.IsAlive = False Then tr3.Start()
        'tr4 = New Threading.Thread(New Threading.ThreadStart(Function() asd.HumidityGetAndSet(EventLog1, e.SignalTime)))
        'tr4.SetApartmentState(Threading.ApartmentState.STA)
        'If tr4.IsAlive = False Then tr4.Start()
        'tr5 = New Threading.Thread(New Threading.ThreadStart(Function() asd.HeatGetAndSet(EventLog1, e.SignalTime)))
        'tr5.SetApartmentState(Threading.ApartmentState.STA)
        'If tr5.IsAlive = False Then tr5.Start()

        tr = New Threading.Thread(New Threading.ThreadStart(Sub() MyLib.CollectData(EventLog1, e.SignalTime)))
        tr.SetApartmentState(Threading.ApartmentState.STA)
        If tr.IsAlive = False Then tr.Start()
        tr3 = New Threading.Thread(New Threading.ThreadStart(Sub() MyLib.PressureGetAndSet(EventLog1, e.SignalTime)))
        tr3.SetApartmentState(Threading.ApartmentState.STA)
        If tr3.IsAlive = False Then tr3.Start()
        tr4 = New Threading.Thread(New Threading.ThreadStart(Sub() MyLib.HumidityGetAndSet(EventLog1, e.SignalTime)))
        tr4.SetApartmentState(Threading.ApartmentState.STA)
        If tr4.IsAlive = False Then tr4.Start()
        tr5 = New Threading.Thread(New Threading.ThreadStart(Sub() MyLib.HeatGetAndSet(EventLog1, e.SignalTime)))
        tr5.SetApartmentState(Threading.ApartmentState.STA)
        If tr5.IsAlive = False Then tr5.Start()
    End Sub
  
    ''' <summary>
    ''' RApor Tablosunu OLuşturur
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Timer3_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles Timer3.Elapsed
        'tr6 = New Threading.Thread(New Threading.ThreadStart(Function() asd.Rapor_Table(e.SignalTime, EventLog1)))
        'tr6.SetApartmentState(Threading.ApartmentState.STA)
        'If tr6.IsAlive = False Then tr6.Start()
        tr6 = New Threading.Thread(New Threading.ThreadStart(Sub() MyLib.Report(e.SignalTime, EventLog1)))
        tr6.SetApartmentState(Threading.ApartmentState.STA)
        If tr6.IsAlive = False Then tr6.Start()

        If e.SignalTime.Second Mod 1 = 0 Then
            Dim servicename As String = "MSSQLSERVER"
            Try
                Dim service As ServiceController = New ServiceController(servicename)
                If ((service.Status.Equals(ServiceControllerStatus.Stopped)) Or (service.Status.Equals(ServiceControllerStatus.StopPending))) Then
                    Dim myService As ServiceController = (From svc In ServiceController.GetServices() Where svc.ServiceName = servicename).ToArray()(0)
                    myService.Start()
                    myService.WaitForStatus(ServiceControllerStatus.Running)
                    EventLog2.WriteEntry(e.SignalTime & " Mssqlservisi Zorla Başlatıldı.")
                End If
            Catch ex As Exception
                EventLog2.WriteEntry(ex.Message)
            End Try
        End If
    End Sub

  
End Class
