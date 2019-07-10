
Imports Microsoft.Win32.SafeHandles
Imports System.Runtime.InteropServices
Imports EasyModbus
Imports System.Data.SqlClient


Public Class KütüpHane
    ''' <summary>
    ''' Sql Connection
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    Public Function con() As SqlConnection
        con = New SqlConnection(My.Settings.Connection)
        con.Open()
        SqlConnection.ClearPool(con)
        SqlConnection.ClearAllPools()

        Return con
    End Function
    ''' <summary>
    ''' modbus bağlantısı
    ''' </summary>
    ''' <param name="ip"></param>
    ''' <param name="port"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ModCon(ip As String, port As Integer, timeout As Integer) As Boolean
        Dim con As New ModbusClient(ip, port)
        con.UDPFlag = True
        con.ConnectionTimeout = timeout
        con.Connect()
        Return con.Connected
    End Function


    ''' <summary>
    ''' Databasedeki deviceset tablosunda ekli olan her bölüme ait 
    ''' her noktayı teker teker gezerek modbas ile bağlanır db ye
    ''' intervaal değerine göre ekleler ve güncel veriyi gösterir.
    ''' Parametreler:Eklerken adressler 4000 den yüksek olanları -1 ekleyerek çekmek gerekiyor
    ''' timerı 1000 ms e ayarlaman gerekiyor
    ''' eğer cihazlara bağlanamazsa bir dahaki sefere bağlanmayı red eder taki manuel olarak
    ''' deviceset tablosundaki descripton "bağlanma" yazzısı değişene kadar
    ''' </summary>
    ''' <param name="log">eventlog1</param>
    ''' <param name="time">timere ait e.spintime eklemen gerekir</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <STAThread()>
    Public Function GetLogs(log As EventLog, time As DateTime) As SqlCommand
        Dim cmd As New SqlCommand("select * from deviceset", con)
        Dim modbus As New ModbusClient
        Dim modbusadres As String = ""
        Dim description As String = ""
        Dim adress As Integer
        'Dim MReg As Integer
        Dim id As Integer


        Try
            Dim dr As SqlDataReader = cmd.ExecuteReader
            Do While dr.Read
                modbus.IPAddress = dr("Host")
                modbus.Port = dr("Port")
                modbus.ConnectionTimeout = 100
                Dim response(1) As Integer
                description = dr("Id")
                Dim StartTime As DateTime = dr("StartTime").ToString
                Dim EndTime As DateTime = dr("EndTime").ToString
                Dim cmd1 As New SqlCommand("select * from datapointset where DeviceId=" & dr("Id"), con)
                Dim dr1 As SqlDataReader = cmd1.ExecuteReader
                Do While dr1.Read
                    'intervala göre değerleri datavaluesete kayıt eder
                    '  log.w

                    modbusadres = dr1("Address").ToString
                    id = dr1("Id")
                    If DateTime.Compare(time.TimeOfDay.ToString, StartTime) >= 0 And DateTime.Compare(time.TimeOfDay.ToString, EndTime) <= 0 And time.Minute Mod dr("Interval") = 0 And time.Second = 0 And dr("Description") <> "Bağlanma" Then

                        If modbus.Connected = False Then modbus.Connect()
                        '    'delta plc lerde modbus d registerlerde adresleri -1 alınır Not:4097=0.registere denk bizim dll'de 4096=0.reg denk
                        If Convert.ToInt32(dr1("Address").ToString.Trim) >= 3000 And dr1("Unit") = 2 Then
                            adress = Convert.ToInt32(dr1("Address").ToString.Trim) + 4097 - 1
                            response(0) = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(adress, 2))
                            modbus.Disconnect()
                        ElseIf Convert.ToInt32(dr1("Address").ToString.Trim) >= 170 And Convert.ToInt32(dr1("Address").ToString.Trim) < 3000 Then
                            adress = Convert.ToInt32(dr1("Address").ToString) + 4097 - 1
                            response = modbus.ReadHoldingRegisters(adress, 1)
                            modbus.Disconnect()
                        Else
                            adress = Convert.ToInt32(dr1("Address").ToString.Trim)
                            response = modbus.ReadHoldingRegisters(adress, 1)
                            modbus.Disconnect()
                        End If

                        Dim cmd2 As New SqlCommand("insert into DataValueset (ValueDate,Value,DatapointId) VALUES (@1,@2,@3)", con)
                        cmd2.Parameters.Add("@1", SqlDbType.DateTime).Value = time.ToString("yyyy-MM-dd HH:mm:ss")
                        cmd2.Parameters.Add("@2", SqlDbType.Decimal).Value = response(0)
                        cmd2.Parameters.Add("@3", SqlDbType.SmallInt).Value = dr1("Id")
                        'log.WriteEntry(time.ToString("yyyy-MM-dd HH:mm:ss"))


                        cmd2.ExecuteNonQuery()
                        Dim cmd3 As New SqlCommand("UPDATE [DataPointSet] SET [CurrentValue] =" & response(0) & " Where Id=" & dr1("Id"), con)
                        cmd3.ExecuteNonQuery()
                    End If
                    '=======================================================================================


                    '========================================================================================
                    If time.Second >= 10 And time.Second Mod 2 = 0 And dr("Description") <> "Bağlanma" Then 'burası 2 saniyede bir verilerin güncel geldiği yer 10-60s arası
                        'burayı tes etmek gerekir 1-2 gün
                        If modbus.Connected = False Then modbus.Connect()
                        If Convert.ToInt32(dr1("Address").ToString.Trim) >= 3000 And dr1("Unit") = 2 Then
                            adress = Convert.ToInt32(dr1("Address").ToString.Trim) + 4097 - 1
                            response(0) = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(adress, 2))
                            modbus.Disconnect()
                        ElseIf Convert.ToInt32(dr1("Address").ToString.Trim) >= 170 And Convert.ToInt32(dr1("Address").ToString.Trim) < 3000 Then
                            adress = Convert.ToInt32(dr1("Address").ToString) + 4097 - 1
                            response = modbus.ReadHoldingRegisters(adress, 1)
                            modbus.Disconnect()
                        Else
                            adress = Convert.ToInt32(dr1("Address").ToString.Trim)
                            response = modbus.ReadHoldingRegisters(adress, 1)
                            modbus.Disconnect()
                        End If
                        Dim cmd3 As New SqlCommand("UPDATE [DataPointSet] SET [CurrentValue] =" & response(0) & " Where Id=" & dr1("Id"), con)
                        cmd3.ExecuteNonQuery()

                        'Enox da alarm uygulanır
                        '===================================

                    End If
                    '=======================================================================================
                    '=======================================================================================








                Loop
                modbus.Disconnect()
            Loop
        Catch ex As Exception
            log.WriteEntry("ip:" & modbus.IPAddress & " Port:" & modbus.Port & " Adress: " & adress & " Id: " & id & " ModbusAdres:" & modbusadres & vbCrLf & ex.Message)

            If ex.Message.Contains(modbus.IPAddress) Then
                log.WriteEntry(modbus.IPAddress & modbus.Port & "  Bağlanamadı ve bi dahaki sefere bağlanamiyacak")
                Dim cmd3 As New SqlCommand("UPDATE Deviceset SET Description='Bağlanma' where Id=" & description, con)
                cmd3.ExecuteNonQuery()
            End If
        End Try
        Return (cmd)
    End Function

    <STAThread()>
    Public Function Rapor_Table(time As DateTime, log As EventLog)
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

        Dim cmd As New SqlCommand("select* from deviceset where Id=5", con)
        Dim dr As SqlDataReader = cmd.ExecuteReader
        Dim modbus As New ModbusClient
        Dim Label As String = ""
        Dim Isı(1) As Integer
        Dim nem(1) As Integer
        Dim basınç(1) As Integer
        modbus.ConnectionTimeout = 100
        ' log.WriteEntry("Raportablos Başladı")
        'Dim i As Integer
        Try
            If dr.Read Then
                Dim StartTime As DateTime = dr("StartTime").ToString
                Dim EndTime As DateTime = dr("EndTime").ToString
                If DateTime.Compare(time.TimeOfDay.ToString, StartTime) >= 0 And DateTime.Compare(time.TimeOfDay.ToString, EndTime) <= 0 And time.Minute Mod dr("Interval") = 0 And
                time.Second = 0 And dr("Description") <> "Bağlanma" Then
                    modbus.IPAddress = dr("Host")
                    modbus.Port = dr("Port")

                    For Each IsıReg In IsıRegister
                        Dim index As Integer = Array.LastIndexOf(IsıRegister, IsıReg)

                        Dim cmd1 As New SqlCommand("Select* from datapointset where Address=" & IsıReg, con) ''ISIlar
                        Dim dr1 As SqlDataReader = cmd1.ExecuteReader

                        If dr1.Read Then
                            Label = dr1("Label").ToString.Replace("Isı", "")
                            If dr1("Unit") = 2 Then
                                If modbus.Connected = False Then modbus.Connect()
                                Isı(0) = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(IsıReg + 4097 - 1, 2))
                                modbus.Disconnect()
                            Else
                                If modbus.Connected = False Then modbus.Connect()
                                Isı = modbus.ReadHoldingRegisters(IsıReg + 4097 - 1, 1)
                                modbus.Disconnect()
                            End If
                        End If

                        Dim cmd2 As New SqlCommand("Select * from datapointset where Address=" & NemRegister(index), con) ''NEM
                        Dim dr2 As SqlDataReader = cmd2.ExecuteReader
                        If dr2.Read Then
                            If dr2("Unit") = 2 Then
                                If modbus.Connected = False Then modbus.Connect()
                                nem(0) = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(NemRegister(index) + 4097 - 1, 2))
                                modbus.Disconnect()
                            Else
                                If modbus.Connected = False Then modbus.Connect()
                                nem = modbus.ReadHoldingRegisters(NemRegister(index) + 4097 - 1, 1)
                                modbus.Disconnect()
                            End If
                        End If
                        Dim cmd3 As New SqlCommand("Select * from datapointset where Address=" & BasınçRegister(index), con) ''BASINÇ
                        Dim dr3 As SqlDataReader = cmd3.ExecuteReader
                        If dr3.Read Then
                            If dr3("Unit") = 2 Then
                                If modbus.Connected = False Then modbus.Connect()
                                basınç(0) = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(BasınçRegister(index) + 4097 - 1, 2))
                                modbus.Disconnect()
                            Else
                                If modbus.Connected = False Then modbus.Connect()
                                basınç = modbus.ReadHoldingRegisters(BasınçRegister(index) + 4097 - 1, 1)
                                modbus.Disconnect()
                            End If
                        End If

                        Dim cmd4 As New SqlCommand("insert into Rapor (Tarih,OdaAdı,Isı,Nem,Basınç) Values (@1,@2,@3,@4,@5)", con)
                        cmd4.Parameters.Add("@1", SqlDbType.DateTime).Value = time.ToString("yyyy-MM-dd HH:mm:ss")
                        cmd4.Parameters.Add("@2", SqlDbType.NVarChar).Value = Label
                        cmd4.Parameters.Add("@3", SqlDbType.Int).Value = Isı(0)
                        cmd4.Parameters.Add("@4", SqlDbType.Int).Value = nem(0)
                        cmd4.Parameters.Add("@5", SqlDbType.Int).Value = basınç(0)
                        cmd4.ExecuteNonQuery()

                    Next
                End If
            End If
        Catch ex As Exception
            log.WriteEntry(ex.Message)
            'log.WriteEntry(i)
        End Try
        Return cmd
    End Function



    ''' <summary>
    ''' PLC DEN Alarm verilerini alır ve değişiklikleri plcye gönderir.
    ''' </summary>
    ''' <param name="ip"></param>
    ''' <param name="log"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadEnoxAlertValue(ip As String, log As EventLog) As SqlCommand
        Dim tablo() As String = {"heat", "humidity", "pressure"}
        Dim kat() As String = {"CCLASS", "BCLASS", "DCLASS"}
        Dim modbus As New ModbusClient
        Dim cmd As New SqlCommand
        ReadEnoxAlertValue = New SqlCommand
        Try
            If modbus.Connected = False Then modbus.Connect(ip, "502")
        Catch ex As Exception
            log.WriteEntry(ex.Message)
            ' listbox.Items.Add(ex.Message)
            'listbox.SetSelected(listbox.Items.Count - 1, True)
        End Try
        Try


            Dim BIsı() As Integer = modbus.ReadHoldingRegisters(4547 - 1, 3) '0. set,1.critic,2.danger
            Dim BNem() As Integer = modbus.ReadHoldingRegisters(4587 - 1, 3) '0. set,1.critic,2.danger
            Dim BBasınç() As Integer = modbus.ReadHoldingRegisters(4627 - 1, 3) '0. set,1.critic,2.danger
            Dim CIsı() As Integer = modbus.ReadHoldingRegisters(4557 - 1, 3) '0. set,1.critic,2.danger
            Dim CNem() As Integer = modbus.ReadHoldingRegisters(4597 - 1, 3) '0. set,1.critic,2.danger
            Dim CBasınç() As Integer = modbus.ReadHoldingRegisters(4637 - 1, 3) '0. set,1.critic,2.danger
            Dim DIsı() As Integer = modbus.ReadHoldingRegisters(4567 - 1, 3) '0. set,1.critic,2.danger
            Dim DNem() As Integer = modbus.ReadHoldingRegisters(4607 - 1, 3) '0. set,1.critic,2.danger
            Dim DBasınç() As Integer = modbus.ReadHoldingRegisters(4647 - 1, 3) '0. set,1.critic,2.danger
            Dim adress() As Integer = {4556, 4557, 4558, 4546, 4547, 4548, 4566, 4567, 4568, 4596, 4597, 4598, 4586, 4587, 4588, 4606, 4607, 4608, 4636, 4637, 4638, 4626, 4627, 4628, 4646, 4647, 4648}
            Dim Value() As Integer = {CIsı(0), CIsı(1), CIsı(2), BIsı(0), BIsı(1), BIsı(2), DIsı(0), DIsı(1), DIsı(2), CNem(0), CNem(1), CNem(2), BNem(0), BNem(1), BNem(2), DNem(0), DNem(1), DNem(2), CBasınç(0), CBasınç(1), CBasınç(2), BBasınç(0), BBasınç(1), BBasınç(2), DBasınç(0), DBasınç(1), DBasınç(2)}

            modbus.Disconnect()

            Dim j As Integer = 0
            For Each tablolar In tablo
                For Each katlar In kat
                    cmd = New SqlCommand("select * from " & tablolar & " where [level]='" & katlar & "'", con)
                    Dim dr As SqlDataReader = cmd.ExecuteReader
                    If dr.Read Then
                        Try
                            If dr("DSet") <> dr("set") And dr("DSet") <> Value(j) Then
                                modbus.ConnectionTimeout = 100
                                If modbus.Connected = False Then modbus.Connect(ip, "502")
                                modbus.WriteSingleRegister(adress(j), dr("DSet"))
                                modbus.Disconnect()
                            End If
                            If dr("DCrit") <> dr("Crit") And dr("DCrit") <> Value(j + 1) Then
                                modbus.ConnectionTimeout = 100
                                If modbus.Connected = False Then modbus.Connect(ip, "502")
                                modbus.WriteSingleRegister(adress(j + 1), dr("DCrit"))
                                modbus.Disconnect()

                            End If
                            If dr("DDanger") <> dr("Danger") And dr("DDanger") <> Value(j + 2) Then
                                modbus.ConnectionTimeout = 100
                                If modbus.Connected = False Then modbus.Connect(ip, "502")
                                modbus.WriteSingleRegister(adress(j + 2), dr("DDanger"))
                                modbus.Disconnect()
                            End If
                        Catch ex As Exception
                            log.WriteEntry("Error Write Single Register : " & ex.Message)
                            'listbox.SetSelected(listbox.Items.Count - 1, True)
                        End Try
                        Dim cmd1 As New SqlCommand("update " & tablolar & " SET [Set]=@1, Crit=@2,Danger=@3 where [level]=@4", con)
                        cmd1.Parameters.Add("@1", SqlDbType.Int).Value = Value(j)
                        cmd1.Parameters.Add("@2", SqlDbType.Int).Value = Value(j + 1)
                        cmd1.Parameters.Add("@3", SqlDbType.Int).Value = Value(j + 2)
                        cmd1.Parameters.Add("@4", SqlDbType.NVarChar).Value = katlar
                        cmd1.ExecuteNonQuery()
                    Else
                        Dim cmd1 As New SqlCommand("insert into " & tablolar & "([Level],[Set],Crit,Danger,DSet,DCrit,DDanger,datapoint) VALUES (@1,@2,@3,@4,@2,@3,@4,@5)", con)
                        cmd1.Parameters.Add("@1", SqlDbType.NVarChar).Value = katlar
                        cmd1.Parameters.Add("@2", SqlDbType.Int).Value = Value(j)
                        cmd1.Parameters.Add("@3", SqlDbType.Int).Value = Value(j + 1)
                        cmd1.Parameters.Add("@4", SqlDbType.Int).Value = Value(j + 2)
                        cmd1.Parameters.Add("@5", SqlDbType.Int).Value = 0
                        cmd1.ExecuteNonQuery()
                    End If
                    j = j + 3
                Next
            Next
        Catch ex As Exception
            log.WriteEntry("AlarmSet Sorunu:" & ex.Message)
        End Try
        Return cmd
    End Function



    ''''ARTIK KULLANILMIYOR
    ''' <summary>
    ''' MLeri OKuyarak db ye akyıt ediyor
    ''' </summary>
    ''' <param name="ipadress"></param>
    ''' <param name="log"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadCoilsM(ipadress As String, log As EventLog)
        Dim cmd As New SqlCommand
        Dim modbus As New ModbusClient
        ReadCoilsM = Nothing

        ' If modbus.Connected = False Then modbus.Connect(ipadress, "503")
        Dim M40x295(256) As Boolean 'M40 dan M295 e kadar 2089 dan 2344 e kadar
        Dim M296x345(50) As Boolean 'm296 dan m 345 e kadar 2345 den 2394 e kadar
        Dim M346x395(50) As Boolean  'm346 dan m 395 e kadar 2345 den 2444 e kadar
        Try
            modbus.ConnectionTimeout = 100
            If modbus.Connected = False Then modbus.Connect(ipadress, "502")
            M40x295 = modbus.ReadCoils(2089, 256) 'M40 dan M295 e kadar 2089 dan 2344 e kadar
            M296x345 = modbus.ReadCoils(2345, 50) 'm296 dan m 345 e kadar 2345 den 2394 e kadar
            M346x395 = modbus.ReadCoils(2395, 50) 'm346 dan m 395 e kadar 2345 den 2444 e kadar
            modbus.Disconnect()
            modbus.Disconnect()

        Catch ex As Exception
            log.WriteEntry("M okumada Sorun :" & ex.Message)
        End Try
        Dim MCoil(512) As Boolean
        Dim m As Integer = 40
        For Each Mcoils In M40x295
            MCoil(m) = Mcoils
            m = m + 1
        Next
        m = 296
        For Each Mcoils In M296x345
            MCoil(m) = Mcoils
            m = m + 1
        Next
        m = 346
        For Each MCoils In M346x395
            MCoil(m) = MCoils
            m = m + 1
        Next

        Dim Label As String = ""
        Dim BClass() As String = {"Dolum Tesisi", "Personel Çıkış 2", "Koridor-3", "Personel Çıkış 1", "Personel Giriş 1", "Personel Giriş 2", "Pess BOX 4", "Otoklav Primer Alan"}
        Dim BOdaNo() As Integer = {709, 711, 707, 710, 705, 706, 0, 708}
        'Dim BClassValue(11) As Boolean = {}
        Dim Cclass() As String = {"Tartım Odası", "Solisyon Hazırlama", "Koridor 4", "Malzeme Airlock C", "Personal Airlock C"}
        Dim COdaNo() As Integer = {715, 714, 713, 716, 712}
        Dim Dclass() As String = {"LYO", "Koridor 2", "Yıkama Alanı", "Pess Box 3", "Primer Malzeme Cep Deposu", "Malzeme Airlock D", "Personal Airlock D", "Koridor 1", "Yarı Mamul Bekleme", "Personel Airlock", "Su Tesisi", "Pess Box 1", "Pess Box 2"}
        Dim DOdaNo() As Integer = {717, 703, 704, 0, 722, 702, 701, 700, 719, 720, 721, 0, 0}

        Dim b As Integer = 0
        Dim c As Integer = 0
        Dim d As Integer = 0
        For Each BClas In BClass
            Dim index As Integer = Array.IndexOf(BClass, BClas, 0)
            cmd = New SqlCommand("select * from EnoxAlert where Label='" & BClas & "'", con)
            Dim dr As SqlDataReader = cmd.ExecuteReader
            If dr.Read Then
                Dim cmd1 As New SqlCommand("update EnoxAlert Set Class=@0,label=@1,RoomNo=@2,CritHeatUp=@3,CritHeatDown=@4,DangerHeatUp=@5,DangerHeatDown=@6,CritHumidityUp=@7,CritHumidityDown=@8,DangerHumidityUp=@9,DangerHumidityDown=@10,CritPressureUp=@11,CritPressureDown=@12,DangerPressureUp=@13,DangerPressureDown=@14 where label=@15", con)
                cmd1.Parameters.Add("@0", SqlDbType.NVarChar).Value = "B-Class"
                cmd1.Parameters.Add("@1", SqlDbType.NVarChar).Value = BClas
                cmd1.Parameters.Add("@2", SqlDbType.Int).Value = BOdaNo(index)
                cmd1.Parameters.Add("@3", SqlDbType.Bit).Value = MCoil(b + 40)
                cmd1.Parameters.Add("@4", SqlDbType.Bit).Value = MCoil(b + 41)
                cmd1.Parameters.Add("@5", SqlDbType.Bit).Value = MCoil(b + 42)
                cmd1.Parameters.Add("@6", SqlDbType.Bit).Value = MCoil(b + 43)
                cmd1.Parameters.Add("@7", SqlDbType.Bit).Value = MCoil(b + 72)
                cmd1.Parameters.Add("@8", SqlDbType.Bit).Value = MCoil(b + 73)
                cmd1.Parameters.Add("@9", SqlDbType.Bit).Value = MCoil(b + 74)
                cmd1.Parameters.Add("@10", SqlDbType.Bit).Value = MCoil(b + 75)
                cmd1.Parameters.Add("@11", SqlDbType.Bit).Value = MCoil(b + 104)
                cmd1.Parameters.Add("@12", SqlDbType.Bit).Value = MCoil(b + 105)
                cmd1.Parameters.Add("@13", SqlDbType.Bit).Value = MCoil(b + 106)
                cmd1.Parameters.Add("@14", SqlDbType.Bit).Value = MCoil(b + 107)
                cmd1.Parameters.Add("@15", SqlDbType.NVarChar).Value = BClas
                cmd1.ExecuteNonQuery()
            Else
                Dim cmd1 As New SqlCommand("insert into EnoxAlert (Class,label,RoomNo,CritHeatUp,CritHeatDown,DangerHeatUp,DangerHeatDown,CritHumidityUp,CritHumidityDown,DangerHumidityUp,DangerHumidityDown,CritPressureUp,CritPressureDown,DangerPressureUp,DangerPressureDown) VALUES (@0,@1,@2,@3,@4,@5,@6,@7,@8,@9,@10,@11,@12,@13,@14)", con)
                cmd1.Parameters.Add("@0", SqlDbType.NVarChar).Value = "B-Class"
                cmd1.Parameters.Add("@1", SqlDbType.NVarChar).Value = BClas
                cmd1.Parameters.Add("@2", SqlDbType.Int).Value = BOdaNo(index)
                cmd1.Parameters.Add("@3", SqlDbType.Bit).Value = MCoil(b + 40)
                cmd1.Parameters.Add("@4", SqlDbType.Bit).Value = MCoil(b + 41)
                cmd1.Parameters.Add("@5", SqlDbType.Bit).Value = MCoil(b + 42)
                cmd1.Parameters.Add("@6", SqlDbType.Bit).Value = MCoil(b + 43)
                cmd1.Parameters.Add("@7", SqlDbType.Bit).Value = MCoil(b + 72)
                cmd1.Parameters.Add("@8", SqlDbType.Bit).Value = MCoil(b + 73)
                cmd1.Parameters.Add("@9", SqlDbType.Bit).Value = MCoil(b + 74)
                cmd1.Parameters.Add("@10", SqlDbType.Bit).Value = MCoil(b + 75)
                cmd1.Parameters.Add("@11", SqlDbType.Bit).Value = MCoil(b + 104)
                cmd1.Parameters.Add("@12", SqlDbType.Bit).Value = MCoil(b + 105)
                cmd1.Parameters.Add("@13", SqlDbType.Bit).Value = MCoil(b + 106)
                cmd1.Parameters.Add("@14", SqlDbType.Bit).Value = MCoil(b + 107)
                cmd1.ExecuteNonQuery()
            End If

            b = b + 4
        Next

        For Each CClas In Cclass
            Dim index As Integer = Array.IndexOf(Cclass, CClas, 0)
            cmd = New SqlCommand("select * from EnoxAlert where Label='" & CClas & "'", con)
            Dim dr As SqlDataReader = cmd.ExecuteReader
            If dr.Read Then
                Dim cmd1 As New SqlCommand("update EnoxAlert Set Class=@0,label=@1,RoomNo=@2,CritHeatUp=@3,CritHeatDown=@4,DangerHeatUp=@5,DangerHeatDown=@6,CritHumidityUp=@7,CritHumidityDown=@8,DangerHumidityUp=@9,DangerHumidityDown=@10,CritPressureUp=@11,CritPressureDown=@12,DangerPressureUp=@13,DangerPressureDown=@14 where label=@15", con)
                cmd1.Parameters.Add("@0", SqlDbType.NVarChar).Value = "C-Class"
                cmd1.Parameters.Add("@1", SqlDbType.NVarChar).Value = CClas
                cmd1.Parameters.Add("@2", SqlDbType.Int).Value = COdaNo(index)
                cmd1.Parameters.Add("@3", SqlDbType.Bit).Value = MCoil(c + 136)
                cmd1.Parameters.Add("@4", SqlDbType.Bit).Value = MCoil(c + 137)
                cmd1.Parameters.Add("@5", SqlDbType.Bit).Value = MCoil(c + 138)
                cmd1.Parameters.Add("@6", SqlDbType.Bit).Value = MCoil(c + 139)
                cmd1.Parameters.Add("@7", SqlDbType.Bit).Value = MCoil(c + 156)
                cmd1.Parameters.Add("@8", SqlDbType.Bit).Value = MCoil(c + 157)
                cmd1.Parameters.Add("@9", SqlDbType.Bit).Value = MCoil(c + 158)
                cmd1.Parameters.Add("@10", SqlDbType.Bit).Value = MCoil(c + 159)
                cmd1.Parameters.Add("@11", SqlDbType.Bit).Value = MCoil(c + 176)
                cmd1.Parameters.Add("@12", SqlDbType.Bit).Value = MCoil(c + 177)
                cmd1.Parameters.Add("@13", SqlDbType.Bit).Value = MCoil(c + 178)
                cmd1.Parameters.Add("@14", SqlDbType.Bit).Value = MCoil(c + 179)
                cmd1.Parameters.Add("@15", SqlDbType.NVarChar).Value = CClas
                cmd1.ExecuteNonQuery()
            Else
                Dim cmd1 As New SqlCommand("insert into EnoxAlert (Class,label,RoomNo,CritHeatUp,CritHeatDown,DangerHeatUp,DangerHeatDown,CritHumidityUp,CritHumidityDown,DangerHumidityUp,DangerHumidityDown,CritPressureUp,CritPressureDown,DangerPressureUp,DangerPressureDown) VALUES (@0,@1,@2,@3,@4,@5,@6,@7,@8,@9,@10,@11,@12,@13,@14)", con)
                cmd1.Parameters.Add("@0", SqlDbType.NVarChar).Value = "C-Class"
                cmd1.Parameters.Add("@1", SqlDbType.NVarChar).Value = CClas
                cmd1.Parameters.Add("@2", SqlDbType.Int).Value = COdaNo(index)
                cmd1.Parameters.Add("@3", SqlDbType.Bit).Value = MCoil(c + 136)
                cmd1.Parameters.Add("@4", SqlDbType.Bit).Value = MCoil(c + 137)
                cmd1.Parameters.Add("@5", SqlDbType.Bit).Value = MCoil(c + 138)
                cmd1.Parameters.Add("@6", SqlDbType.Bit).Value = MCoil(c + 139)
                cmd1.Parameters.Add("@7", SqlDbType.Bit).Value = MCoil(c + 156)
                cmd1.Parameters.Add("@8", SqlDbType.Bit).Value = MCoil(c + 157)
                cmd1.Parameters.Add("@9", SqlDbType.Bit).Value = MCoil(c + 158)
                cmd1.Parameters.Add("@10", SqlDbType.Bit).Value = MCoil(c + 159)
                cmd1.Parameters.Add("@11", SqlDbType.Bit).Value = MCoil(c + 176)
                cmd1.Parameters.Add("@12", SqlDbType.Bit).Value = MCoil(c + 177)
                cmd1.Parameters.Add("@13", SqlDbType.Bit).Value = MCoil(c + 178)
                cmd1.Parameters.Add("@14", SqlDbType.Bit).Value = MCoil(c + 179)
                cmd1.ExecuteNonQuery()
            End If

            c = c + 4
        Next

        For Each DClas In Dclass
            Dim index As Integer = Array.IndexOf(Dclass, DClas, 0)
            cmd = New SqlCommand("select * from EnoxAlert where Label='" & DClas & "'", con)
            Dim dr As SqlDataReader = cmd.ExecuteReader
            If dr.Read Then
                Dim cmd1 As New SqlCommand("update EnoxAlert Set Class=@0,label=@1,RoomNo=@2,CritHeatUp=@3,CritHeatDown=@4,DangerHeatUp=@5,DangerHeatDown=@6,CritHumidityUp=@7,CritHumidityDown=@8,DangerHumidityUp=@9,DangerHumidityDown=@10,CritPressureUp=@11,CritPressureDown=@12,DangerPressureUp=@13,DangerPressureDown=@14 where label=@15", con)
                cmd1.Parameters.Add("@0", SqlDbType.NVarChar).Value = "D-Class"
                cmd1.Parameters.Add("@1", SqlDbType.NVarChar).Value = DClas
                cmd1.Parameters.Add("@2", SqlDbType.Int).Value = DOdaNo(index)
                cmd1.Parameters.Add("@3", SqlDbType.Bit).Value = MCoil(d + 196)
                cmd1.Parameters.Add("@4", SqlDbType.Bit).Value = MCoil(d + 197)
                cmd1.Parameters.Add("@5", SqlDbType.Bit).Value = MCoil(d + 198)
                cmd1.Parameters.Add("@6", SqlDbType.Bit).Value = MCoil(d + 199)
                cmd1.Parameters.Add("@7", SqlDbType.Bit).Value = MCoil(d + 248)
                cmd1.Parameters.Add("@8", SqlDbType.Bit).Value = MCoil(d + 249)
                cmd1.Parameters.Add("@9", SqlDbType.Bit).Value = MCoil(d + 250)
                cmd1.Parameters.Add("@10", SqlDbType.Bit).Value = MCoil(d + 251)
                cmd1.Parameters.Add("@11", SqlDbType.Bit).Value = MCoil(d + 300)
                cmd1.Parameters.Add("@12", SqlDbType.Bit).Value = MCoil(d + 301)
                cmd1.Parameters.Add("@13", SqlDbType.Bit).Value = MCoil(d + 302)
                cmd1.Parameters.Add("@14", SqlDbType.Bit).Value = MCoil(d + 303)
                cmd1.Parameters.Add("@15", SqlDbType.NVarChar).Value = DClas
                cmd1.ExecuteNonQuery()
            Else
                Dim cmd1 As New SqlCommand("insert into EnoxAlert (Class,label,RoomNo,CritHeatUp,CritHeatDown,DangerHeatUp,DangerHeatDown,CritHumidityUp,CritHumidityDown,DangerHumidityUp,DangerHumidityDown,CritPressureUp,CritPressureDown,DangerPressureUp,DangerPressureDown) VALUES (@0,@1,@2,@3,@4,@5,@6,@7,@8,@9,@10,@11,@12,@13,@14)", con)
                cmd1.Parameters.Add("@0", SqlDbType.NVarChar).Value = "D-Class"
                cmd1.Parameters.Add("@1", SqlDbType.NVarChar).Value = DClas
                cmd1.Parameters.Add("@2", SqlDbType.Int).Value = DOdaNo(index)
                cmd1.Parameters.Add("@3", SqlDbType.Bit).Value = MCoil(d + 196)
                cmd1.Parameters.Add("@4", SqlDbType.Bit).Value = MCoil(d + 197)
                cmd1.Parameters.Add("@5", SqlDbType.Bit).Value = MCoil(d + 198)
                cmd1.Parameters.Add("@6", SqlDbType.Bit).Value = MCoil(d + 199)
                cmd1.Parameters.Add("@7", SqlDbType.Bit).Value = MCoil(d + 248)
                cmd1.Parameters.Add("@8", SqlDbType.Bit).Value = MCoil(d + 249)
                cmd1.Parameters.Add("@9", SqlDbType.Bit).Value = MCoil(d + 250)
                cmd1.Parameters.Add("@10", SqlDbType.Bit).Value = MCoil(d + 251)
                cmd1.Parameters.Add("@11", SqlDbType.Bit).Value = MCoil(d + 300)
                cmd1.Parameters.Add("@12", SqlDbType.Bit).Value = MCoil(d + 301)
                cmd1.Parameters.Add("@13", SqlDbType.Bit).Value = MCoil(d + 302)
                cmd1.Parameters.Add("@14", SqlDbType.Bit).Value = MCoil(d + 303)
                cmd1.ExecuteNonQuery()
            End If
            d = d + 4
        Next


        Return (cmd)
    End Function

    ''' <summary>
    ''' Basınç registerlerini günceller
    ''' dbde currentvalue column'a daki değişiklik plcdeki registere gönderir
    ''' 
    ''' </summary>
    ''' <param name="time">Timerin e.spentime yaz</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <STAThread()>
    Public Function PressureGetAndSet(log As EventLog, time As DateTime) As SqlCommand
        PressureGetAndSet = New SqlCommand("select * from pressureset", con)
        Dim modbus As New ModbusClient
        '
        Dim dr As SqlDataReader = PressureGetAndSet.ExecuteReader
        Try

            Do While dr.Read
                Dim cmd As New SqlCommand("select * from deviceset where Id=(select DeviceId from datapointset where Id=" & dr("DatapointId") & ")", con)
                Dim cmdr As SqlDataReader = cmd.ExecuteReader
                Dim description As String = ""
                Dim response As Integer
                If cmdr.Read Then
                    modbus.IPAddress = cmdr("Host")
                    modbus.Port = cmdr("Port")
                    modbus.ConnectionTimeout = 100
                    description = cmdr("Description")
                End If
                If time.Second Mod 3 = 0 And description <> "Bağlanma" Then
                    If modbus.Connected = False Then modbus.Connect()
                    response = ModbusClient.ConvertRegistersToDouble(modbus.ReadHoldingRegisters(dr("Adres") + 4097 - 1, 2))
                    modbus.Disconnect()
                    Dim cmd1 As New SqlCommand("update pressureset SET Value=@1 where Id=" & dr("Id"), con)
                    cmd1.Parameters.Add("@1", SqlDbType.Int).Value = response
                    cmd1.ExecuteNonQuery()
                    'MsgBox(response)
                    If dr("CheckValue") <> dr("Value") And dr("CheckValue") <> response Then
                        If dr("CheckValue") = 0 Then
                            Dim cmd2 As New SqlCommand("Update pressureset Set CheckValue=" & response & " Where Id=" & dr("Id"), con)
                            cmd2.ExecuteNonQuery()

                        Else
                            If modbus.Connected = False Then modbus.Connect()
                            modbus.WriteMultipleRegisters(dr("Adres") + 4097 - 1, ModbusClient.ConvertDoubleToTwoRegisters(dr("CheckValue")))
                            modbus.Disconnect()
                        End If
                    End If
                End If
            Loop
        Catch ex As Exception
            log.WriteEntry(ex.Message)
        End Try
        Return PressureGetAndSet
    End Function
    <STAThread()>
    Public Function HeatGetAndSet(log As EventLog, time As DateTime) As SqlCommand
        HeatGetAndSet = New SqlCommand("select * from heatset", con)
        Dim modbus As New ModbusClient

        Dim dr As SqlDataReader = HeatGetAndSet.ExecuteReader
        Try

            Do While dr.Read
                Dim cmd As New SqlCommand("select * from deviceset where Id=(select DeviceId from datapointset where Id=" & dr("DatapointId") & ")", con)
                Dim cmdr As SqlDataReader = cmd.ExecuteReader
                Dim description As String = ""
                Dim response(1) As Integer
                If cmdr.Read Then
                    modbus.IPAddress = cmdr("Host")
                    modbus.Port = cmdr("Port")
                    modbus.ConnectionTimeout = 100
                    description = cmdr("Description")
                End If
                If time.Second Mod 3 = 0 And description <> "Bağlanma" Then
                    If modbus.Connected = False Then modbus.Connect()
                    response = modbus.ReadHoldingRegisters(dr("Adres") + 4097 - 1, 1)
                    modbus.Disconnect()
                    Dim cmd1 As New SqlCommand("update heatset SET Value=@1 where Id=" & dr("Id"), con)
                    cmd1.Parameters.Add("@1", SqlDbType.Int).Value = response(0)
                    cmd1.ExecuteNonQuery()
                    'MsgBox(response)
                    If dr("CheckValue") <> dr("Value") And dr("CheckValue") <> response(0) Then
                        If dr("CheckValue") = 0 Then
                            Dim cmd2 As New SqlCommand("Update heatset Set CheckValue=" & response(0) & " Where Id=" & dr("Id"), con)
                            cmd2.ExecuteNonQuery()
                        Else
                            If modbus.Connected = False Then modbus.Connect()
                            modbus.WriteSingleRegister(dr("Adres") + 4097 - 1, dr("CheckValue"))
                            modbus.Disconnect()
                        End If
                    End If
                End If
            Loop
        Catch ex As Exception
            log.WriteEntry(ex.Message)
        End Try
        Return HeatGetAndSet
    End Function
    <STAThread()>
    Public Function HumidityGetAndSet(log As EventLog, time As DateTime) As SqlCommand
        HumidityGetAndSet = New SqlCommand("select * from Humidityset", con)
        Dim modbus As New ModbusClient
        '
        Dim dr As SqlDataReader = HumidityGetAndSet.ExecuteReader
        Try

            Do While dr.Read
                Dim cmd As New SqlCommand("select * from deviceset where Id=(select DeviceId from datapointset where Id=" & dr("DatapointId") & ")", con)
                Dim cmdr As SqlDataReader = cmd.ExecuteReader
                Dim description As String = ""
                Dim response() As Integer
                If cmdr.Read Then
                    modbus.IPAddress = cmdr("Host")
                    modbus.Port = cmdr("Port")
                    modbus.ConnectionTimeout = 100
                    description = cmdr("Description")
                End If
                If time.Second Mod 3 = 0 And description <> "Bağlanma" Then
                    If modbus.Connected = False Then modbus.Connect()
                    response = modbus.ReadHoldingRegisters(dr("Adres") + 4097 - 1, 1)
                    modbus.Disconnect()
                    Dim cmd1 As New SqlCommand("update Humidityset SET Value=@1 where Id=" & dr("Id"), con)
                    cmd1.Parameters.Add("@1", SqlDbType.Int).Value = response(0)
                    cmd1.ExecuteNonQuery()
                    'MsgBox(response)
                    If dr("CheckValue") <> dr("Value") And dr("CheckValue") <> response(0) Then
                        If dr("CheckValue") = 0 Then
                            Dim cmd2 As New SqlCommand("Update Humidityset Set CheckValue=" & response(0) & " Where Id=" & dr("Id"), con)
                            cmd2.ExecuteNonQuery()
                        Else
                            If modbus.Connected = False Then modbus.Connect()
                            modbus.WriteSingleRegister(dr("Adres") + 4097 - 1, dr("CheckValue"))
                            modbus.Disconnect()
                        End If
                    End If
                End If
            Loop
        Catch ex As Exception
            log.WriteEntry(ex.Message)
        End Try
        Return HumidityGetAndSet
    End Function

End Class
