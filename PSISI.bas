Attribute VB_Name = "Module3"
Sub PSISI()

'**********************************************************************************************************************
'
'                                       формирование Alias DVBC_MPTS_G2X_B1
'**********************************************************************************************************************
        
    Worksheets.Add.Name = "DVBC_MPTS_G2X_B1" 'создание листа
    Call Shapka("DVBC_MPTS_G2X_B1")
    Call Config("Краснодар")
    
    Worksheets.Add.Name = "DVBC_MPTS_G2X_B1" 'создание листа
    Call Shapka("DVBC_MPTS_G2X_B1")
    Call Config("Екатеринбург")
    
    Worksheets.Add.Name = "DVBC_MPTS_G2X_B1" 'создание листа
    Call Shapka("DVBC_MPTS_G2X_B1")
    Call Config("Новосибирск")
    
    Worksheets.Add.Name = "DVBC_MPTS_G2X_B1" 'создание листа
    Call Shapka("DVBC_MPTS_G2X_B1")
    Call Config("Н. Новгород")
    
    Worksheets.Add.Name = "DVBC_MPTS_G2X_B1" 'создание листа
    Call Shapka("DVBC_MPTS_G2X_B1")
    Call Config("Владивосток")
End Sub

    
Sub Config(City As String)
 Dim i
 Dim s
 Dim Tes
 Dim test As String
  
 Dim ID_num
 Dim IDShort_num
 Dim FTempl_num
 Dim PTempl_num
 Dim Group_num
 Dim Multicast_IP
 Dim Source_IP
 
 i = 1
' поиск столбцов с данными
'-------------------------------------------------------------------------------------------------------------------------------
 Do Until Sheets("параметры цс").Cells(1, i).Value = 0 And Sheets("параметры цс").Cells(1, i + 1).Value = 0 And Sheets("параметры цс").Cells(1, i + 2).Value = 0 And Sheets("параметры цс").Cells(1, i + 3).Value = 0
 If Sheets("параметры цс").Cells(1, i).Value = "Multicast IP" Then Multicast_IP = i
 If Sheets("параметры цс").Cells(1, i).Value = "Source IP (main)" Then Source_IP = i
 If Sheets("параметры цс").Cells(1, i).Value = "ID (name IQ)" Then ID_num = i
 If Sheets("параметры цс").Cells(1, i).Value = "сокращенный ID (name IQ)" Then IDShort_num = i
 If Sheets("параметры цс").Cells(1, i).Value = "Template Flow IQ" Then FTempl_num = i
 If Sheets("параметры цс").Cells(1, i).Value = "Template Program IQ" Then PTempl_num = i
 If Sheets("параметры цс").Cells(1, i).Value = "Group" Then Group_num = i
 i = i + 1
 Loop
 
 
' формирование флоу алиас DVBC_MPTS_G2X_B1
'-------------------------------------------------------------------------------------------------------------------------------
Tes = 0
s = 3
i = 2
  Do Until Sheets("параметры цс").Cells(2 + Tes, 1).MergeArea.Cells(1, 1).Value = ""
  '  If _ист1.Cells(i, 14).Value = "-1" Then
  '          i = i + 1
Set MPTS = Sheets("параметры цс").Cells(2 + Tes, 1).MergeArea 'диапазон объединенных _чеек названи_ потоков
Set Multicast = Sheets("параметры цс").Cells(2 + Tes, Multicast_IP).MergeArea
Set Source = Sheets("параметры цс").Cells(2 + Tes, Source_IP).MergeArea
Tes = Tes + Sheets("параметры цс").Cells(2 + Tes, 1).MergeArea.Count
Set MPTS_next = Sheets("параметры цс").Cells(2 + Tes, 1).MergeArea
Set Multicast_next = Sheets("параметры цс").Cells(2 + Tes, Multicast_IP).MergeArea
Set Source_next = Sheets("параметры цс").Cells(2 + Tes, Source_IP).MergeArea

' исключение старого дубл_ потока
    If MPTS.Cells(1, 1).Value = MPTS_next.Cells(1, 1).Value Then                ' если 2 подр_д одинаковых потока
        Tes = Tes + Sheets("параметры цс").Cells(2 + Tes, 1).MergeArea.Count
        If Multicast_next.Cells(1, 1).Interior.Color = 5296274 Then              ' если зеленый цвет аливки это нужный поток
            Set MPTS = MPTS_next
            Set Multicast = Multicast_next
            Set Source = Source_next
        End If
    End If
    
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 1).Value = "Video"
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 2).Formula = "MPTS DVBC " & Right(MPTS.Cells(1, 1).Value, Len(MPTS.Cells(1, 1)) - InStr(1, MPTS.Cells(1, 1), "TS") - 1)
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 3).Value = Source.Cells(1, 1).Value 'source adress
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 4).Value = Left(Multicast.Cells(1, 1), InStr(1, Multicast.Cells(1, 1), ":") - 1) 'multicast adress
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 5).Value = "No"
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 6).Value = Right(Multicast.Cells(1, 1), Len(Multicast.Cells(1, 1)) - InStr(1, Multicast.Cells(1, 1), ":")) 'порт
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 7).Value = "On"
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 8) = Sheets("параметры цс").Cells(Tes, FTempl_num).Value
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 9).Value = "No"
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 10) = "MTSProgramDefault"
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 11).Value = "255.255.255.255"
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 12).Value = "255.255.255.255"
        Sheets("DVBC_MPTS_G2X_B1").Cells(s, 13).Value = "4"

            For n = 15 To 30
                Select Case n
                    Case 15 To 20
                    Sheets("DVBC_MPTS_G2X_B1").Cells(s, n).Value = "No"
                    Case 21
                    Sheets("DVBC_MPTS_G2X_B1").Cells(s, n).Value = 1
                    Case 22 To 24
                    Sheets("DVBC_MPTS_G2X_B1").Cells(s, n).Value = "No"
                    Case 25
                    Sheets("DVBC_MPTS_G2X_B1").Cells(s, n).Value = 0
                    Case 26 To 27
                    Sheets("DVBC_MPTS_G2X_B1").Cells(s, n).Value = "No"
                    Case 28
                    Sheets("DVBC_MPTS_G2X_B1").Cells(s, n).Value = Sheets("параметры цс").Cells(Tes, Group_num)
                    Case 29 To 30
                    Sheets("DVBC_MPTS_G2X_B1").Cells(s, n).Value = "No"
                End Select
            Next n
        i = i + 1
        s = s + 1
 Loop
 
 s = PSISI_city(City, s, "DVBC_MPTS_G2X_B1")
' формирование программ алиас
 '-----------------------------------------------------------------------------------------------------------------
     i = 1
     Tes = 0
     temp = 0 'число строк инкрементации после пропуска дубл_
     
             
               Do While i > 0
               
               ' исключение программ дубл_ старого потока
               Set MPTS = Sheets("параметры цс").Cells(2 + Tes, 1).MergeArea 'диапазон объединенных _чеек названи_ пото
               Tes = Tes + Sheets("параметры цс").Cells(2 + Tes, 1).MergeArea.Count
                Set MPTS_next = Sheets("параметры цс").Cells(2 + Tes, 1).MergeArea
                Set Multicast_next = Sheets("параметры цс").Cells(2 + Tes, Multicast_IP).MergeArea
                 If MPTS.Cells(1, 1).Value = MPTS_next.Cells(1, 1).Value Then
                    Tes = Tes + Sheets("параметры цс").Cells(2 + Tes, 1).MergeArea.Count
                   
                    If Multicast_next.Cells(1, 1).Interior.Color = 5296274 Then
                    i = i + MPTS.Cells.Count
                    Set MPTS = MPTS_next
                    Else:
                    temp = MPTS_next.Cells.Count
                     
                    End If
                  End If
                
                           For n = 1 To MPTS.Cells.Count
                           If Sheets("параметры цс").Cells(i + 1, 3).Font.Strikethrough = True Or Sheets("параметры цс").Cells(i + 1, 3).Interior.Color = 255 Then ' исключить если канал зачеркнут
                            i = i + 1
                           Else:
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 1).Value = "Video"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 2).Formula = "MPTS DVBC " & Right(MPTS.Cells(1, 1).Value, Len(MPTS.Cells(1, 1)) - InStr(1, MPTS.Cells(1, 1), "TS") - 1)
                            'Sheets("DVBC_MPTS_G2X_B1").Cells(s, 2).Value = MPTS.Cells(1, 1).Value
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 3).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 4).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 5).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 6).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 7).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 8).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 9).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 10).Value = Sheets("параметры цс").Cells(i + 1, PTempl_num).Value 'темплейт программы
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 11).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 12).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 13).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 14).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 15).Value = Sheets("параметры цс").Cells(i + 1, 7).Value 'SID
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 16).Value = Sheets("параметры цс").Cells(i + 1, ID_num).Value 'Hазвание программы
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 17).Value = Sheets("параметры цс").Cells(i + 1, 8).Value 'LCN
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 18).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 19).Value = "0_0.0:0.0"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 20).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 21).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 22).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 23).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 24).Value = Sheets("параметры цс").Cells(i + 1, ID_num).Value 'Hазвание программы
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 25).Value = 0
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 26).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 27).Value = 0
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 28).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 29).Value = "No"
                            Sheets("DVBC_MPTS_G2X_B1").Cells(s, 30).Value = "No"
                            s = s + 1
                            i = i + 1
                           End If
                            Next n
                            i = i + temp
                            temp = 0
                            'Tes = Tes + Sheets("параметры цс").Cells(2 + Tes, 1).MergeArea.Count
                    If Sheets("параметры цс").Cells(i + 1, 2).Value = "" Then Exit Do
                    
                Loop
                
                Suffix0 = "\DVBC_MPTS_G2X_B1_port1,2_"
                If City = "Владивосток" Then
                    Name_file = Suffix0 & "Комсомольск" & "_"
                Else: Name_file = Suffix0 & City & "_"
                End If
                
                Suffix1 = City
                
    Call SaveFile(Name_file)
    
End Sub
Sub SaveFile(Suffix0)
' сохранение в файл
'---------------------------------------------------------------------------------------------------------------------------------

    'Suffix0 = "\DVBC_MPTS_G2X_B1_port1,2_"
    Suffix1 = ".xls"
    NewFileName = ThisWorkbook.Path & Suffix0 & Format(Date, "DDMMYY") & Format(Time, "_hhmm") & Suffix1
    Sheets("DVBC_MPTS_G2X_B1").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=NewFileName _
        , FileFormat:=xlText, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    ActiveWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
' удаление листа
'---------------------------------------------------------------------------------------------------------------------------------
    Application.DisplayAlerts = False
    Sheets("DVBC_MPTS_G2X_B1").Delete
    Application.DisplayAlerts = True
End Sub

Sub Shapka(List As String)

    Sheets(List).Cells(1, 1).Value = "version(O)"
    Sheets(List).Cells(1, 2).Value = "name(1)"
    Sheets(List).Cells(1, 3).Value = "sourceIp(2)"
    Sheets(List).Cells(1, 4).Value = "destIp(3)"
    Sheets(List).Cells(1, 5).Value = "srcPort(4)"
    Sheets(List).Cells(1, 6).Value = "destPort(5)"
    Sheets(List).Cells(1, 7).Value = "igmpStatus(6)"
    Sheets(List).Cells(1, 8).Value = "alarmTemplate(7)"
    Sheets(List).Cells(1, 9).Value = "VLANTCI(8)"
    Sheets(List).Cells(1, 10).Value = "payloadTemplate(9)"
    Sheets(List).Cells(1, 11).Value = "srcIpMask(10)"
    Sheets(List).Cells(1, 12).Value = "destIpMask(11)"
    Sheets(List).Cells(1, 13).Value = "BroadCast(12)"
    Sheets(List).Cells(1, 14).Value = "MACforARPReply(13)"
    Sheets(List).Cells(1, 15).Value = "channelNumber(15)"
    Sheets(List).Cells(1, 16).Value = "channelName(14)"
    Sheets(List).Cells(1, 17).Value = "channelAliasNumber(18)"
    Sheets(List).Cells(1, 18).Value = "deviceRef(22)"
    Sheets(List).Cells(1, 19).Value = "channelOffPeriod(32)"
    Sheets(List).Cells(1, 20).Value = "channelOffAirTemplate(33)"
    Sheets(List).Cells(1, 21).Value = "IGMP Sets(31)"
    Sheets(List).Cells(1, 22).Value = "RTP SSRC(35)"
    Sheets(List).Cells(1, 23).Value = "NonMediaProgram(37)"
    Sheets(List).Cells(1, 24).Value = "channelXRefName(201)"
    Sheets(List).Cells(1, 25).Value = "channelSourceId(20)"
    Sheets(List).Cells(1, 26).Value = "channelShortName(40)"
    Sheets(List).Cells(1, 27).Value = "AliasDetectionMode(43)"
    Sheets(List).Cells(1, 28).Value = "Ports(34)"
    Sheets(List).Cells(1, 29).Value = "Transport Stream ID(30)"
    Sheets(List).Cells(1, 30).Value = "DetectedProgramName(38)"
    
    Sheets(List).Cells(2, 1).Value = "Video"
    Sheets(List).Cells(2, 2).Value = "None"
    Sheets(List).Cells(2, 3).Value = "No"
    Sheets(List).Cells(2, 4).Value = "No"
    Sheets(List).Cells(2, 5).Value = "No"
    Sheets(List).Cells(2, 6).Value = "No"
    Sheets(List).Cells(2, 7).Value = "Off"
    Sheets(List).Cells(2, 8).Value = "tsDefault"
    Sheets(List).Cells(2, 9).Value = "No"
    Sheets(List).Cells(2, 10).Value = "programDefault"
    Sheets(List).Cells(2, 11).Value = "255.255.255.255"
    Sheets(List).Cells(2, 12).Value = "255.255.255.255"
    Sheets(List).Cells(2, 13).Value = "No"
    Sheets(List).Cells(2, 14).Value = "No"
    Sheets(List).Cells(2, 15).Value = "No"
    Sheets(List).Cells(2, 16).Value = "No"
    Sheets(List).Cells(2, 17).Value = "No"
    Sheets(List).Cells(2, 18).Value = "No"
    Sheets(List).Cells(2, 19).Value = "No"
    Sheets(List).Cells(2, 20).Value = "No"
    Sheets(List).Cells(2, 21).Value = "0"
    Sheets(List).Cells(2, 22).Value = "No"
    Sheets(List).Cells(2, 23).Value = "No"
    Sheets(List).Cells(2, 24).Value = "No"
    Sheets(List).Cells(2, 25).Value = "0"
    Sheets(List).Cells(2, 26).Value = "No"
    Sheets(List).Cells(2, 27).Value = "No"
    Sheets(List).Cells(2, 28).Value = "1"
    Sheets(List).Cells(2, 29).Value = "No"
    Sheets(List).Cells(2, 30).Value = "No"
    
End Sub
Function PSISI_city(City As String, alias_str, List As String)

 Dim i
 Dim SRC_IP_PSI_Main
 Dim SRC_IP_PSI_Back
 Dim Multicast_PSI
 Dim City_Src_main
 Dim City_Src_backup
 Dim City_Dist
 Dim FTempl_num_PSI
 Dim ID_PSI_num_main
 Dim ID_PSI_num_backup
 Dim Status_PSI_main
 Dim Status_PSI_backup
 Dim Group_PSI_main
 Dim Group_PSI_back
 Dim Num_TS_PSI
 Dim test As String
 

 i = 1
 Do Until Sheets("Сетевой PSI").Cells(1, i).Value = 0 And Sheets("Сетевой PSI").Cells(1, i + 1).Value = 0
 If Sheets("Сетевой PSI").Cells(1, i).Value = "город установки основного генератора" Then City_Src_main = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "город установки резервного генератора" Then City_Src_backup = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "Multicast IP" Then Multicast_PSI = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "Source IP (main)" Then SRC_IP_PSI_Main = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "Source IP (backup)" Then SRC_IP_PSI_Back = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "cтатус вещания main" Then Status_PSI_main = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "cтатус вещания backup" Then Status_PSI_back = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "IQ allias_main" Then ID_PSI_num_main = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "IQ allias_backup" Then ID_PSI_num_backup = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "Template Flow IQ" Then FTempl_num_PSI = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "Group_main" Then Group_PSI_main = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "Group_backup" Then Group_PSI_back = i
 If Sheets("Сетевой PSI").Cells(1, i).Value = "№ TS назначения потока" Then Num_TS_PSI = i
  i = i + 1
 Loop
 
 
 ' формирование флоу алиас для IP source main
'-------------------------------------------------------------------------------------------------------------------------------
  s = alias_str
  i = 2
  
  Do Until Sheets("Сетевой PSI").Cells(i, 1).Value = ""
  test = Sheets("Сетевой PSI").Cells(i, City_Src_main).Value
  If Sheets("Сетевой PSI").Cells(i, Status_PSI_main).Value = "1" And Sheets("Сетевой PSI").Cells(i, City_Src_main).Value = City Then
    
        Sheets(List).Cells(s, 1).Value = "Video"
        Sheets(List).Cells(s, 2).Value = Sheets("Сетевой PSI").Cells(i, ID_PSI_num_main).Value & "-main-" & Sheets("Сетевой PSI").Cells(i, Num_TS_PSI).Value
        Sheets(List).Cells(s, 3).Value = Sheets("Сетевой PSI").Cells(i, SRC_IP_PSI_Main).Value 'source adress
        Sheets(List).Cells(s, 4).Value = Left(Sheets("Сетевой PSI").Cells(i, Multicast_PSI), InStr(1, Sheets("Сетевой PSI").Cells(i, Multicast_PSI), ":") - 1) 'multicast adress
        Sheets(List).Cells(s, 5).Value = "No"
        Sheets(List).Cells(s, 6).Value = Right(Sheets("Сетевой PSI").Cells(i, Multicast_PSI), Len(Sheets("Сетевой PSI").Cells(i, Multicast_PSI)) - InStr(1, Sheets("Сетевой PSI").Cells(i, Multicast_PSI), ":")) 'порт
        Sheets(List).Cells(s, 7).Value = "On"
        Sheets(List).Cells(s, 8).Value = Sheets("Сетевой PSI").Cells(i, FTempl_num_PSI).Value
        Sheets(List).Cells(s, 9).Value = "No"
        Sheets(List).Cells(s, 10) = "No"
        Sheets(List).Cells(s, 11).Value = "255.255.255.255"
        Sheets(List).Cells(s, 12).Value = "255.255.255.255"
        Sheets(List).Cells(s, 13).Value = "4"

            For n = 15 To 30
                Select Case n
                    Case 15 To 20
                    Sheets(List).Cells(s, n).Value = "No"
                    Case 21
                    Sheets(List).Cells(s, n).Value = 1
                    Case 22 To 24
                    Sheets(List).Cells(s, n).Value = "No"
                    Case 25
                    Sheets(List).Cells(s, n).Value = 0
                    Case 26 To 27
                    Sheets(List).Cells(s, n).Value = "No"
                    Case 28
                    Sheets(List).Cells(s, n).Value = Sheets("Сетевой PSI").Cells(i, Group_PSI_main).Value
                    Case 29 To 30
                    Sheets(List).Cells(s, n).Value = "No"
                End Select
            Next n
        s = s + 1
  End If
        i = i + 1
 Loop
 
  
' формирование флоу алиас для IP source backup
'-------------------------------------------------------------------------------------------------------------------------------
  i = 2
  Do Until Sheets("Сетевой PSI").Cells(i, 1).Value = ""
  test = Sheets("Сетевой PSI").Cells(i, City_Src_backup).Value
  If Sheets("Сетевой PSI").Cells(i, Status_PSI_back).Value = "1" And Sheets("Сетевой PSI").Cells(i, City_Src_backup).Value = City Then
    
        Sheets(List).Cells(s, 1).Value = "Video"
        Sheets(List).Cells(s, 2).Value = Sheets("Сетевой PSI").Cells(i, ID_PSI_num_backup).Value & "-backup-" & Sheets("Сетевой PSI").Cells(i, Num_TS_PSI).Value
        Sheets(List).Cells(s, 3).Value = Sheets("Сетевой PSI").Cells(i, SRC_IP_PSI_Back).Value 'source adress
        Sheets(List).Cells(s, 4).Value = Left(Sheets("Сетевой PSI").Cells(i, Multicast_PSI), InStr(1, Sheets("Сетевой PSI").Cells(i, Multicast_PSI), ":") - 1) 'multicast adress
        Sheets(List).Cells(s, 5).Value = "No"
        Sheets(List).Cells(s, 6).Value = Right(Sheets("Сетевой PSI").Cells(i, Multicast_PSI), Len(Sheets("Сетевой PSI").Cells(i, Multicast_PSI)) - InStr(1, Sheets("Сетевой PSI").Cells(i, Multicast_PSI), ":")) 'порт
        Sheets(List).Cells(s, 7).Value = "On"
        Sheets(List).Cells(s, 8).Value = Sheets("Сетевой PSI").Cells(i, FTempl_num_PSI).Value
        Sheets(List).Cells(s, 9).Value = "No"
        Sheets(List).Cells(s, 10) = "No"
        Sheets(List).Cells(s, 11).Value = "255.255.255.255"
        Sheets(List).Cells(s, 12).Value = "255.255.255.255"
        Sheets(List).Cells(s, 13).Value = "4"

            For n = 15 To 30
                Select Case n
                    Case 15 To 20
                    Sheets(List).Cells(s, n).Value = "No"
                    Case 21
                    Sheets(List).Cells(s, n).Value = 1
                    Case 22 To 24
                    Sheets(List).Cells(s, n).Value = "No"
                    Case 25
                    Sheets(List).Cells(s, n).Value = 0
                    Case 26 To 27
                    Sheets(List).Cells(s, n).Value = "No"
                    Case 28
                    Sheets(List).Cells(s, n).Value = Sheets("Сетевой PSI").Cells(i, Group_PSI_back).Value
                    Case 29 To 30
                    Sheets(List).Cells(s, n).Value = "No"
                End Select
            Next n
        s = s + 1
  End If
        i = i + 1
 Loop
 
 PSISI_city = s

End Function
