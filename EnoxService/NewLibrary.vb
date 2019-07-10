Imports EasyModbus
Imports System.Data.SqlClient
Public Class NewLibrary
    <STAThread>
    Public Sub CollectData(log As EventLog, time As DateTime)
        Dim modbus As New ModbusClient
        Dim context As New EnoxDatabaseDataContext
        Dim response(1) As Integer
        Dim adres As Integer = Nothing
        Dim Response1(1) As Integer
        Dim adres1 As Integer = Nothing
        Dim id As Integer = Nothing
        Dim id1 As Integer = Nothing
        Dim modbusadres As Integer = Nothing
        Dim modbusadres1 As Integer = Nothing
        Dim deviceid As Integer = Nothing
        Try
            Dim Query = (From it In context.DeviceSets Select it).ToList
            For Each DeviceRow As DeviceSet In Query

                Dim query2 = (From it In context.DataPointSets Where it.DeviceId = DeviceRow.Id Select it).ToList
                For Each DatapointRow As DataPointSet In query2
                    id = DatapointRow.Id
                    deviceid = DeviceRow.Id

                    modbusadres = Convert.ToInt32(DatapointRow.Address.ToString.Trim)
                    Dim StartTime As DateTime = DeviceRow.StartTime.ToString
                    Dim EndTime As DateTime = DeviceRow.EndTime.ToString
                    '<<==başlangıç zamanı ile bitiş zamanı içinde interval aralıklarıyla database ekler ve Günceller(bizim databasede her dakikada çalışıyor burası ==>>
                    If DateTime.Compare(time.TimeOfDay.ToString, StartTime) >= 0 And DateTime.Compare(time.TimeOfDay.ToString, EndTime) <= 0 And time.Minute Mod DeviceRow.Interval = 0 And time.Second = 0 And DeviceRow.Description.ToString <> "Bağlanma" Then
                        modbus.IPAddress = DeviceRow.Host.ToString
                        modbus.Port = DeviceRow.Port
                        modbus.ConnectionTimeout = 100
                        If modbus.Connected = False Then modbus.Connect()

                        If Convert.ToInt32(DatapointRow.Address.ToString.Trim) >= 3000 And DatapointRow.Unit = 2 Then
                            adres = Convert.ToInt32(DatapointRow.Address.ToString.Trim) + 4097 - 1
                            response(0) = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(adres, 2))
                            modbus.Disconnect()
                        ElseIf Convert.ToInt32(DatapointRow.Address.ToString.Trim) >= 170 And Convert.ToInt32(DatapointRow.Address.ToString.Trim) < 3000 Then
                            adres = Convert.ToInt32(DatapointRow.Address.ToString.Trim) + 4097 - 1
                            response = modbus.ReadHoldingRegisters(adres, 1)
                        Else
                            adres = Convert.ToInt32(DatapointRow.Address.ToString.Trim)
                            response = modbus.ReadHoldingRegisters(adres, 1)
                            modbus.Disconnect()
                        End If


                        Dim AddDataValue As New DataValueSet
                        AddDataValue.ValueDate = time.ToString("yyyy-MM-dd HH:mm:ss")
                        AddDataValue.Value = response(0)
                        AddDataValue.DataPointId = DatapointRow.Id
                        context.DataValueSets.InsertOnSubmit(AddDataValue)
                        context.SubmitChanges()
                        Dim update = (From it In context.DataPointSets Where it.Id = DatapointRow.Id).First
                        update.CurrentValue = response(0)
                        context.SubmitChanges()

                    End If
                    Try
                        '===========================
                        If time.Second >= 10 And time.Second Mod 2 = 0 And DeviceRow.Description.ToString.Trim <> "Bağlanma" Then
                            id1 = DatapointRow.Id
                            modbusadres1 = Convert.ToInt32(DatapointRow.Address.ToString.Trim)
                            modbus.IPAddress = DeviceRow.Host.ToString
                            modbus.Port = DeviceRow.Port
                            modbus.ConnectionTimeout = 100
                            If modbus.Connected = False Then modbus.Connect()

                            If Convert.ToInt32(DatapointRow.Address.ToString.Trim) >= 3000 And DatapointRow.Unit = 2 Then
                                adres1 = Convert.ToInt32(DatapointRow.Address.ToString.Trim) + 4097 - 1
                                Response1(0) = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(adres1, 2))
                                modbus.Disconnect()
                            ElseIf Convert.ToInt32(DatapointRow.Address.ToString.Trim) >= 170 And Convert.ToInt32(DatapointRow.Address.ToString.Trim) < 3000 Then
                                adres1 = Convert.ToInt32(DatapointRow.Address.ToString.Trim) + 4097 - 1
                                Response1 = modbus.ReadHoldingRegisters(adres1, 1)
                            Else
                                adres1 = Convert.ToInt32(DatapointRow.Address.ToString.Trim)
                                Response1 = modbus.ReadHoldingRegisters(adres1, 1)
                                modbus.Disconnect()
                            End If
                            Dim update = (From it In context.DataPointSets Where it.Id = DatapointRow.Id).First
                            update.CurrentValue = Response1(0)
                            context.SubmitChanges()
                        End If
                        '===========================
                    Catch ex As Exception
                        log.WriteEntry("2 Sn de bir veri Çekmede sorun var. " & "ip:" & modbus.IPAddress & " Port:" & modbus.Port & " Adress: " & adres1 & " Id: " & id1 & " ModbusAdres:" & modbusadres1 & vbCrLf & ex.Message & vbCrLf & ex.InnerException.ToString)
                        If ex.Message.Contains(modbus.IPAddress) Then
                            log.WriteEntry(modbus.IPAddress & ":" & modbus.Port & "2 Sn de bir( Bağlanamadı ve bi dahaki sefere bağlanamiyacak)")
                            Dim update = (From it In context.DeviceSets Where it.Id = deviceid Select it).First
                            update.Description = "Bağlanma"
                            context.SubmitChanges()
                        End If
                        modbus.Disconnect()
                    End Try

                Next
                modbus.Disconnect()
            Next

        Catch ex As Exception
            log.WriteEntry("ip:" & modbus.IPAddress & " Port:" & modbus.Port & " Adress: " & adres & " Id: " & id & " ModbusAdres:" & modbusadres & vbCrLf & ex.Message)
            If ex.Message.Contains(modbus.IPAddress) Then
                log.WriteEntry(modbus.IPAddress & ":" & modbus.Port & "  Bağlanamadı ve bi dahaki sefere bağlanamiyacak")
                Dim update = (From it In context.DeviceSets Where it.Id = deviceid Select it).First
                update.Description = "Bağlanma"
                context.SubmitChanges()
            End If
            modbus.Disconnect()
        End Try


    End Sub

    <STAThread>
    Public Sub Report(time As DateTime, log As EventLog)
        '===================================================================================================================================
        '===================================================================================================================================
        '=================================AŞAĞIDAKİ REGİSTERLERİ AYARLA EXCELDEN============================================================
        '===================================================================================================================================
        '===================================================================================================================================
        '===================================================================================================================================
        Dim IsıRegister() As Integer = {170, 171, 172, 173, 174, 260, 261, 262, 263, 264, 265, 266, 267, 360, 361, 362, 363, 364, 365, 366, 367, 368, 369, 370, 371, 372}
        Dim NemRegister() As Integer = {195, 196, 197, 198, 199, 285, 286, 287, 288, 289, 290, 291, 292, 385, 386, 387, 388, 389, 390, 391, 392, 393, 394, 395, 396, 397}
        Dim BasınçRegister() As Integer = {3006, 3002, 3042, 3044, 3022, 3014, 3012, 3018, 3048, 3024, 3010, 3028, 3016, 3040, 3020, 3000, 3030, 3008, 3038, 3036, 3046, 3004, 3050, 3034, 3026, 3032}
        Dim adressler() As Integer =
{170, 195, 3006,
171, 196, 3002,
172, 197, 3042,
173, 198, 3044,
174, 199, 3022,
260, 285, 3014,
261, 286, 3012,
262, 287, 3018,
263, 288, 3048,
264, 289, 3024,
265, 290, 3010,
266, 291, 3028,
267, 292, 3016,
360, 385, 3040,
361, 386, 3020,
362, 387, 3000,
363, 388, 3030,
364, 389, 3008,
365, 390, 3038,
366, 391, 3036,
367, 392, 3046,
368, 393, 3004,
369, 394, 3050,
370, 395, 3034,
371, 396, 3026,
372, 397, 3032
}
        '================================================================================================
        '================================================================================================
        '================================================================================================
        '================================================================================================
        '================================================================================================
        '================================================================================================
        Dim context As New EnoxDatabaseDataContext
        Dim modbus As New ModbusClient
        Dim enoxip As String = "10.0.0.9"

        Dim Isı(1) As Integer
        Dim nem(1) As Integer
        Dim basınç(1) As Integer
        modbus.ConnectionTimeout = 100

        Try

            Dim StartTimes As DateTime = (From it In context.DeviceSets Where it.Host = enoxip
                                          Select it.StartTime).SingleOrDefault.ToString
            Dim EndTimes As DateTime = (From it In context.DeviceSets Where it.Host = enoxip Select it.EndTime).SingleOrDefault.ToString
            Dim intervals As Integer = (From it In context.DeviceSets Where it.Host = enoxip Select it.Interval).SingleOrDefault
            Dim descriptions As String = (From it In context.DeviceSets Where it.Host = enoxip Select it.Description).SingleOrDefault
            modbus.IPAddress = (From it In context.DeviceSets Where it.Host = enoxip Select it.Host).SingleOrDefault
            modbus.Port = (From it In context.DeviceSets Where it.Host = enoxip Select it.Port).SingleOrDefault

            If DateTime.Compare(time.TimeOfDay.ToString, StartTimes) >= 0 And DateTime.Compare(time.TimeOfDay.ToString, EndTimes) <= 0 And time.Minute Mod intervals = 0 And
                time.Second = 0 And descriptions <> "Bağlanma" Then
                For Each IsıReg In IsıRegister
                    Dim index As Integer = Array.LastIndexOf(IsıRegister, IsıReg)
                    'ISI
                    Dim Unitısı = (From it In context.DataPointSets Where it.Address = IsıReg Select it.Unit).SingleOrDefault
                    Dim labelısı = (From it In context.DataPointSets Where it.Address = IsıReg Select it.Label).SingleOrDefault
                    If Unitısı = 2 Then
                        If modbus.Connected = False Then modbus.Connect()
                        Isı(0) = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(IsıReg + 4097 - 1, 2))
                        modbus.Disconnect()
                    Else
                        If modbus.Connected = False Then modbus.Connect()
                        Isı = modbus.ReadHoldingRegisters(IsıReg + 4097 - 1, 1)
                        modbus.Disconnect()
                    End If
                    'Nem
                    Dim unitnem = (From it In context.DataPointSets Where it.Address = NemRegister(index) Select it.Unit).SingleOrDefault
                    If unitnem = 2 Then
                        If modbus.Connected = False Then modbus.Connect()
                        nem(0) = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(NemRegister(index) + 4097 - 1, 2))
                        modbus.Disconnect()
                    Else
                        If modbus.Connected = False Then modbus.Connect()
                        nem = modbus.ReadHoldingRegisters(NemRegister(index) + 4097 - 1, 1)
                        modbus.Disconnect()
                    End If
                    'Basınç
                    Dim unitbasınç = (From it In context.DataPointSets Where it.Address = BasınçRegister(index) Select it.Unit).SingleOrDefault
                    If unitbasınç = 2 Then
                        If modbus.Connected = False Then modbus.Connect()
                        basınç(0) = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(BasınçRegister(index) + 4097 - 1, 2))
                        modbus.Disconnect()
                    Else
                        If modbus.Connected = False Then modbus.Connect()
                        basınç = modbus.ReadHoldingRegisters(BasınçRegister(index) + 4097 - 1, 1)
                        modbus.Disconnect()
                    End If
                    Dim addrapor As New Rapor
                    addrapor.Tarih = time.ToString("yyyy-MM-dd HH:mm:ss")
                    addrapor.OdaAdı = labelısı
                    addrapor.Isı = Isı(0)
                    addrapor.Nem = nem(0)
                    addrapor.Basınç = basınç(0)
                    context.Rapors.InsertOnSubmit(addrapor)
                    context.SubmitChanges()
                Next
            End If
        Catch ex As Exception
            log.WriteEntry("Raporlamada Sıkıntı Var:" & ex.Message)
            modbus.Disconnect()
        End Try
    End Sub
    <STAThread>
    Public Sub PressureGetAndSet(log As EventLog, time As DateTime)
        Dim modbus As New ModbusClient
        Try
            Dim context As New EnoxDatabaseDataContext
            Dim query = (From it In context.PressureSets Select it)
            Dim response(1) As Integer
            For Each Row As PressureSet In query
                Dim descriptions As String = (From it In context.DeviceSets Where it.Id = (From a In context.DataPointSets Where a.Id = Row.DatapointId Select a.DeviceId).SingleOrDefault Select it.Description).SingleOrDefault
                Dim ipadress As String = (From it In context.DeviceSets Where it.Id = (From a In context.DataPointSets Where a.Id = Row.DatapointId Select a.DeviceId).SingleOrDefault Select it.Host).SingleOrDefault
                Dim ports As Integer = (From it In context.DeviceSets Where it.Id = (From a In context.DataPointSets Where a.Id = Row.DatapointId Select a.DeviceId).SingleOrDefault Select it.Port).SingleOrDefault
                modbus.IPAddress = ipadress
                modbus.Port = ports
                modbus.ConnectionTimeout = 100

                If time.Second Mod 3 = 0 And descriptions <> "Bağlanma" Then
                    If modbus.Connected = False Then modbus.Connect()
                    response(0) = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(Row.Adres + 4097 - 1, 2))
                    modbus.Disconnect()
                    Dim update = (From it In context.PressureSets Where it.Id = Row.Id Select it).First
                    update.Value = response(0)
                    context.SubmitChanges()
                    If Row.CheckValue <> Row.Value And Row.CheckValue <> response(0) Then
                        If Row.CheckValue = 0 Then
                            Dim update1 = (From it In context.PressureSets Where it.Id = Row.Id Select it).First
                            update1.CheckValue = response(0)
                            context.SubmitChanges()
                        Else
                            If modbus.Connected = False Then modbus.Connect()
                            modbus.WriteMultipleRegisters(Row.Adres + 4097 - 1, ModbusClient.ConvertDoubleToTwoRegisters(Row.CheckValue))
                            modbus.Disconnect()
                        End If

                    End If

                End If


            Next
        Catch ex As Exception
            log.WriteEntry("Basınç Offsetlerinde sıkıntı var:" & ex.Message)
            modbus.Disconnect()
        End Try


    End Sub
    <STAThread>
    Public Sub HeatGetAndSet(log As EventLog, time As DateTime)
        Dim modbus As New ModbusClient
        Try
            Dim context As New EnoxDatabaseDataContext
            Dim query = (From it In context.HeatSets Select it)
            Dim response(1) As Integer
            For Each Row As HeatSet In query
                Dim descriptions As String = (From it In context.DeviceSets Where it.Id = (From a In context.DataPointSets Where a.Id = Row.DatapointId Select a.DeviceId).SingleOrDefault Select it.Description).SingleOrDefault
                Dim ipadress As String = (From it In context.DeviceSets Where it.Id = (From a In context.DataPointSets Where a.Id = Row.DatapointId Select a.DeviceId).SingleOrDefault Select it.Host).SingleOrDefault
                Dim ports As Integer = (From it In context.DeviceSets Where it.Id = (From a In context.DataPointSets Where a.Id = Row.DatapointId Select a.DeviceId).SingleOrDefault Select it.Port).SingleOrDefault
                modbus.IPAddress = ipadress
                modbus.Port = ports
                modbus.ConnectionTimeout = 100

                If time.Second Mod 3 = 0 And descriptions <> "Bağlanma" Then
                    If modbus.Connected = False Then modbus.Connect()
                    response = modbus.ReadHoldingRegisters(Row.Adres + 4097 - 1, 1)
                    modbus.Disconnect()
                    Dim update = (From it In context.HeatSets Where it.Id = Row.Id Select it).First
                    update.Value = response(0)
                    context.SubmitChanges()
                    If Row.CheckValue <> Row.Value And Row.CheckValue <> response(0) Then
                        If Row.CheckValue = 0 Then
                            Dim update1 = (From it In context.HeatSets Where it.Id = Row.Id Select it).First
                            update1.CheckValue = response(0)
                            context.SubmitChanges()
                        Else
                            If modbus.Connected = False Then modbus.Connect()
                            modbus.WriteSingleRegister(Row.Adres + 4097 - 1, Row.CheckValue)
                            modbus.Disconnect()
                        End If

                    End If

                End If


            Next
        Catch ex As Exception
            log.WriteEntry("Sıcaklık Offsetlerinde sıkıntı var:" & ex.Message)
            modbus.Disconnect()
        End Try

    End Sub
    <STAThread>
    Public Sub HumidityGetAndSet(log As EventLog, time As DateTime)
        Dim modbus As New ModbusClient
        Try
            Dim context As New EnoxDatabaseDataContext
            Dim query = (From it In context.HumiditySets Select it)

            Dim response(1) As Integer
            For Each Row As HumiditySet In query
                Dim descriptions As String = (From it In context.DeviceSets Where it.Id = (From a In context.DataPointSets Where a.Id = Row.DatapointId Select a.DeviceId).SingleOrDefault Select it.Description).SingleOrDefault
                Dim ipadress As String = (From it In context.DeviceSets Where it.Id = (From a In context.DataPointSets Where a.Id = Row.DatapointId Select a.DeviceId).SingleOrDefault Select it.Host).SingleOrDefault
                Dim ports As Integer = (From it In context.DeviceSets Where it.Id = (From a In context.DataPointSets Where a.Id = Row.DatapointId Select a.DeviceId).SingleOrDefault Select it.Port).SingleOrDefault
                modbus.IPAddress = ipadress
                modbus.Port = ports
                modbus.ConnectionTimeout = 100

                If time.Second Mod 3 = 0 And descriptions <> "Bağlanma" Then
                    If modbus.Connected = False Then modbus.Connect()
                    response = modbus.ReadHoldingRegisters(Row.Adres + 4097 - 1, 1)
                    modbus.Disconnect()
                    Dim update = (From it In context.HumiditySets Where it.Id = Row.Id Select it).First
                    update.Value = response(0)
                    context.SubmitChanges()
                    If Row.CheckValue <> Row.Value And Row.CheckValue <> response(0) Then
                        If Row.CheckValue = 0 Then
                            Dim update1 = (From it In context.HumiditySets Where it.Id = Row.Id Select it).First
                            update1.CheckValue = response(0)
                            context.SubmitChanges()
                        Else
                            If modbus.Connected = False Then modbus.Connect()
                            modbus.WriteSingleRegister(Row.Adres + 4097 - 1, Row.CheckValue)
                            modbus.Disconnect()
                        End If

                    End If

                End If


            Next
        Catch ex As Exception
            log.WriteEntry("Nem Offsetlerinde sıkıntı var:" & ex.Message)
            modbus.Disconnect()
        End Try

    End Sub
End Class
