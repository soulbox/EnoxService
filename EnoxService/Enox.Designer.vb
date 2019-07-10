Imports System.ServiceProcess

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Enox
    Inherits System.ServiceProcess.ServiceBase

    'UserService overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    <System.Diagnostics.DebuggerNonUserCode()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New Enox}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.EventLog1 = New System.Diagnostics.EventLog()
        Me.Timer1 = New System.Timers.Timer()
        Me.Timer2 = New System.Timers.Timer()
        Me.Timer3 = New System.Timers.Timer()
        Me.EventLog2 = New System.Diagnostics.EventLog()
        CType(Me.EventLog1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Timer2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Timer3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EventLog2, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000.0R
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000.0R
        '
        'Timer3
        '
        Me.Timer3.Enabled = True
        Me.Timer3.Interval = 1000.0R
        '
        'Enox
        '
        Me.ServiceName = "Enox Service"
        CType(Me.EventLog1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Timer2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Timer3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EventLog2, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Friend WithEvents EventLog1 As System.Diagnostics.EventLog
    Friend WithEvents Timer1 As System.Timers.Timer
    Friend WithEvents Timer2 As System.Timers.Timer
    Friend WithEvents Timer3 As System.Timers.Timer
    Friend WithEvents EventLog2 As EventLog
End Class
