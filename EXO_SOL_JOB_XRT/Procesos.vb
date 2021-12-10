Imports System.IO
Imports System.Text
Imports System.Xml
Imports OfficeOpenXml
Imports Sap.Data.Hana
Public Class Procesos
#Region "Actualizar campos"
    Public Shared Sub Actualizar_Campos(ByRef oLog As EXO_Log.EXO_Log)
        Dim oDBSAP As HanaConnection = Nothing
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim sError As String = ""
        Dim sSQL As String = ""
        Dim OdtDatos As System.Data.DataTable = Nothing
        Dim sPass As String = ""
        Dim refDI As EXO_DIAPI.EXO_DIAPI = Nothing
        Dim oXML As String = ""
        Dim sDir As String = ""
        Try
            sDir = Application.StartupPath
            sPass = Conexiones.Datos_Confi("DI", "Password")
            Dim tipoServidor As SAPbobsCOM.BoDataServerTypes = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            Dim tipocliente As EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico
            Conexiones.Connect_SQLHANA(oDBSAP, "HANA", oLog)
            sSQL = " SELECT ""EXO_BD"" FROM ""SOL_AUTORIZ"".""EXO_SOCIEDADES"" "
            OdtDatos = New System.Data.DataTable
            Conexiones.FillDtDB(oDBSAP, OdtDatos, sSQL)
            If OdtDatos.Rows.Count > 0 Then
                oLog.escribeMensaje("Se va a proceder a recorrer las SOCIEDADES...", EXO_Log.EXO_Log.Tipo.advertencia)
                For Each dr In OdtDatos.Rows
                    Conexiones.Connect_Company(oCompany, "DI", dr("EXO_BD").ToString, oLog)
#Region "Creamos campos"
                    Try
                        refDI = New EXO_DIAPI.EXO_DIAPI(oCompany, oLog)
                    Catch ex As Exception
                        refDI = New EXO_DIAPI.EXO_DIAPI(tipoServidor, oCompany.Server.ToString, oCompany.LicenseServer.ToString, oCompany.CompanyDB.ToString, oCompany.UserName.ToString, sPass, tipocliente)
                    End Try

                    Dim fsXML As New FileStream(sDir & "\XML_BD\UDFs_EXO_OADM.xml", FileMode.Open, FileAccess.Read)
                    Dim xmldoc As New XmlDocument()
                    xmldoc.Load(fsXML)
                    oXML = xmldoc.InnerXml.ToString
                    refDI.comunes.LoadBDFromXML(oXML, sError)
                    oLog.escribeMensaje("Validado: UDFs_EXO_OADM - " & sError, EXO_Log.EXO_Log.Tipo.advertencia)

                    fsXML = New FileStream(sDir & "\XML_BD\UT_EXO_DPTO.xml", FileMode.Open, FileAccess.Read)
                    xmldoc.Load(fsXML)
                    oXML = xmldoc.InnerXml.ToString
                    refDI.comunes.LoadBDFromXML(oXML, sError)
                    oLog.escribeMensaje("Validado: UT_EXO_DPTO - " & sError, EXO_Log.EXO_Log.Tipo.advertencia)

                    fsXML = New FileStream(sDir & "\XML_BD\UDFs_EXO_DSC1.xml", FileMode.Open, FileAccess.Read)
                    xmldoc.Load(fsXML)
                    oXML = xmldoc.InnerXml.ToString
                    refDI.comunes.LoadBDFromXML(oXML, sError)
                    oLog.escribeMensaje("Validado: UDFs_EXO_DSC1 - " & sError, EXO_Log.EXO_Log.Tipo.advertencia)

                    fsXML = New FileStream(sDir & "\XML_BD\UDFs_EXO_OVPM.xml", FileMode.Open, FileAccess.Read)
                    xmldoc.Load(fsXML)
                    oXML = xmldoc.InnerXml.ToString
                    refDI.comunes.LoadBDFromXML(oXML, sError)
                    oLog.escribeMensaje("Validado: UDFs_EXO_OVPM - " & sError, EXO_Log.EXO_Log.Tipo.advertencia)

                    fsXML = New FileStream(sDir & "\XML_BD\UT_EXO_XRTFLUJOS.xml", FileMode.Open, FileAccess.Read)
                    xmldoc.Load(fsXML)
                    oXML = xmldoc.InnerXml.ToString
                    refDI.comunes.LoadBDFromXML(oXML, sError)
                    oLog.escribeMensaje("Validado: UT_EXO_XRTFLUJOS - " & sError, EXO_Log.EXO_Log.Tipo.advertencia)
                    Carga_Flujos(oLog, dr("EXO_BD").ToString, oDBSAP)



                    fsXML = New FileStream(sDir & "\XML_BD\UT_EXO_XRTCPP.xml", FileMode.Open, FileAccess.Read)
                    xmldoc.Load(fsXML)
                    oXML = xmldoc.InnerXml.ToString
                    refDI.comunes.LoadBDFromXML(oXML, sError)
                    oLog.escribeMensaje("Validado: UT_EXO_XRTCPP - " & sError, EXO_Log.EXO_Log.Tipo.advertencia)
                    Carga_CPP(oLog, dr("EXO_BD").ToString, oDBSAP)

                    fsXML = New FileStream(sDir & "\XML_BD\UDF_EXO_DOCs.xml", FileMode.Open, FileAccess.Read)
                    xmldoc.Load(fsXML)
                    oXML = xmldoc.InnerXml.ToString
                    refDI.comunes.LoadBDFromXML(oXML, sError)
                    oLog.escribeMensaje("Validado: UDF_EXO_DOCs - " & sError, EXO_Log.EXO_Log.Tipo.advertencia)
#End Region
                    Conexiones.Disconnect_Company(oCompany)
                Next
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            Conexiones.Disconnect_SQLHANA(oDBSAP)
            refDI = Nothing
        End Try

    End Sub
    Public Shared Sub Carga_Flujos(ByRef oLog As EXO_Log.EXO_Log, ByVal SBBDD As String, ByRef oDBSAP As HanaConnection)
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos    
        Dim sSQL As String = ""
        Dim sPath As String = ""
        Dim sCodigo As String = ""
        Dim sDescripcion As String = ""
        Dim pck As ExcelPackage = Nothing
        Dim iLin As Integer = 0
        Dim sError As String = ""
        Try
            oLog.escribeMensaje("Insertando Flujos XRT ... Espere por favor.", EXO_Log.EXO_Log.Tipo.advertencia)

            sPath = My.Application.Info.DirectoryPath.ToString & "\EXCEL\"
            If Not System.IO.Directory.Exists(sPath) Then
                oLog.escribeMensaje("No existe direcotrio de EXCEL para capturar datos.", EXO_Log.EXO_Log.Tipo.error)
            End If
            sPath += "\FLujos.xlsx"
            If File.Exists(sPath) Then
                Dim excel As New FileInfo(sPath)
                pck = New ExcelPackage(excel)
                Dim workbook = pck.Workbook

                Dim worksheet = workbook.Worksheets.First()
                For iLin = 2 To 65
                    sCodigo = worksheet.Cells("A" & iLin).Text
                    sDescripcion = worksheet.Cells("C" & iLin).Text
                    sSQL = "insert into """ & SBBDD & """.""@EXO_XRTFLUJOS"" values('" & Trim(sCodigo) & "','" & Trim(sDescripcion) & "')"
                    Try
                        Conexiones.ExecuteSqlDB(oDBSAP, sSQL)
                    Catch ex As Exception
                        oLog.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.advertencia)
                    End Try
                Next
                oLog.escribeMensaje("Se han actualizado los Flujos de XRT.", EXO_Log.EXO_Log.Tipo.informacion)
                pck.Dispose()
            Else
                oLog.escribeMensaje("No se ha encontrado el Path del fichero a cargar.", EXO_Log.EXO_Log.Tipo.error)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally

        End Try
    End Sub
    Public Shared Sub Carga_CPP(ByRef oLog As EXO_Log.EXO_Log, ByVal SBBDD As String, ByRef oDBSAP As HanaConnection)
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos    
        Dim sSQL As String = ""
        Dim sPath As String = ""
        Dim sCodigo As String = ""
        Dim sDescripcion As String = ""
        Dim pck As ExcelPackage = Nothing
        Dim iLin As Integer = 0
        Dim sError As String = ""
        Try
            oLog.escribeMensaje("Insertando Cod.. Presupuestarios XRT ... Espere por favor.", EXO_Log.EXO_Log.Tipo.advertencia)

            sPath = My.Application.Info.DirectoryPath.ToString & "\EXCEL\"
            If Not System.IO.Directory.Exists(sPath) Then
                oLog.escribeMensaje("No existe direcotrio de EXCEL para capturar datos.", EXO_Log.EXO_Log.Tipo.error)
            End If
            sPath += "\CPP.xlsx"
            If File.Exists(sPath) Then
                Dim excel As New FileInfo(sPath)
                pck = New ExcelPackage(excel)
                Dim workbook = pck.Workbook

                Dim worksheet = workbook.Worksheets.First()
                For iLin = 2 To 46
                    sCodigo = worksheet.Cells("A" & iLin).Text
                    sDescripcion = worksheet.Cells("B" & iLin).Text
                    sSQL = "insert into """ & SBBDD & """.""@EXO_XRTCPP"" values('" & Trim(sCodigo) & "','" & Trim(sDescripcion) & "')"
                    Try
                        Conexiones.ExecuteSqlDB(oDBSAP, sSQL)
                    Catch ex As Exception
                        oLog.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.advertencia)
                    End Try
                Next
                oLog.escribeMensaje("Se han actualizado los Cod. Presupuestarios de XRT.", EXO_Log.EXO_Log.Tipo.informacion)
                pck.Dispose()
            Else
                oLog.escribeMensaje("No se ha encontrado el Path del fichero a cargar.", EXO_Log.EXO_Log.Tipo.error)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally

        End Try
    End Sub
#End Region
    Public Shared Sub Prev_origen(ByRef oLog As EXO_Log.EXO_Log)
#Region "Variables"
        Dim oDBSAP As HanaConnection = Nothing
        Dim sError As String = ""
        Dim sSQL As String = "" : Dim ESprimero As Boolean = True : Dim sBBDD As String = ""
        Dim odtDatos As System.Data.DataTable = Nothing
#End Region
        Try
            Conexiones.Connect_SQLHANA(oDBSAP, "HANA", oLog)
            'Buscamos las facturas pdtes de cobro
            sSQL = " SELECT ""EXO_BD"" FROM ""SOL_AUTORIZ"".""EXO_SOCIEDADES"" "
            odtDatos = New System.Data.DataTable
            Conexiones.FillDtDB(oDBSAP, odtDatos, sSQL)
            If odtDatos.Rows.Count > 0 Then
                oLog.escribeMensaje("Se va a proceder a recorrer las SOCIEDADES...", EXO_Log.EXO_Log.Tipo.advertencia)
                sSQL = "SELECT * FROM (" : ESprimero = True
                For Each dr As DataRow In odtDatos.Rows
                    sBBDD = dr.Item("EXO_BD").ToString

                    If ESprimero = False Then
                        sSQL &= " UNION ALL "
                    Else
                        ESprimero = False
                    End If
                    Dim sCtaFicticia As String = Conexiones.GetValueDB(oDBSAP, """" & sBBDD & """.""OADM""", """U_EXO_CTAXRT""", "")
                    oLog.escribeMensaje("La Cta. ficticia de la empresa " & sBBDD & " es:" & sCtaFicticia, EXO_Log.EXO_Log.Tipo.informacion)
                    If sCtaFicticia = "" Then
                        oLog.escribeMensaje("No se ha asignado. Por defecto se pondrá la CTA CTANOASIGN", EXO_Log.EXO_Log.Tipo.advertencia)
                        sCtaFicticia = "CTANOASIGN"
                    End If
                    'Proveedor
                    sSQL &= "(Select '" + sBBDD + "' ""BD"",T0.""NumAtCard"" ""REFERENCIA"" ,'PPOR' ""Flujo"", T0.""U_EXO_XRTCPP"" ""PRESU"", T0.""DocDate"" ""F_OPE"", T0.""DocDueDate"" ""F_VTO"","
                    sSQL &= " round(T0.""DocTotalFC""- T0.""PaidFC"",2) ""ImpDIV"", T0.""DocCur"" ""DIV"",  round(T0.""DocTotal""-T0.""PaidToDate"",2) ""ImpT"", 'EUR' ""DIVEUR"", '" & sCtaFicticia & "' ""Cuenta"", "
                    sSQL &= " ifnull(T3.""Name"",'TRANSFERENCIA PROVEEDOR') ""TEXTO"", T0.""CardCode"" ""ZONA1"", T0.""DocNum"" ""ZONA2"", T2.""Descript"" ""ZONA3"", T1.""PymntGroup"", '' ""ZONA4"", "
                    sSQL &= " T0.""CardName"" ""ZONA5"", 'OPCH' ""Tabla"", T0.""ObjType"" ""OBJTYPE"" "
                    sSQL &= " FROM """ + sBBDD + """.""OPCH"" T0 "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""@EXO_XRTFLUJOS"" T3 on T0.""U_EXO_XRTFlujo""=T3.""Code"" "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""OCTG"" T1 on T0.""GroupNum""=T1.""GroupNum"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OPYM"" T2 on T0.""PeyMethod""=T2.""PayMethCod"" "
                    sSQL &= " WHERE T0.""CANCELED"" ='N'  AND T0.""DocStatus"" <>'C' and T0.""DocTotal"">0 and year(T0.""DocDueDate"")>=" & Now.Year.ToString("0000") & ") "
                    sSQL &= " UNION ALL "
                    sSQL &= " (Select '" + sBBDD + "' ""BD"",T0.""NumAtCard"" ""REFERENCIA"" ,'PPOR' ""Flujo"", T0.""U_EXO_XRTCPP"" ""PRESU"", T0.""DocDate"" ""F_OPE"", T0.""DocDueDate"" ""F_VTO"","
                    sSQL &= " round(T0.""DocTotalFC""- T0.""PaidFC"",2) ""ImpDIV"", T0.""DocCur"" ""DIV"",  round(T0.""DocTotal""-T0.""PaidToDate"",2) ""ImpT"", 'EUR' ""DIVEUR"", '" & sCtaFicticia & "' ""Cuenta"", "
                    sSQL &= " ifnull(T3.""Name"",'TRANSFERENCIA PROVEEDOR') ""TEXTO"", T0.""CardCode"" ""ZONA1"", T0.""DocNum"" ""ZONA2"", T2.""Descript"" ""ZONA3"", T1.""PymntGroup"", '' ""ZONA4"", "
                    sSQL &= " T0.""CardName"" ""ZONA5"",'ORPC' ""Tabla"", T0.""ObjType"" ""OBJTYPE"" "
                    sSQL &= " FROM """ + sBBDD + """.""ORPC"" T0 "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""@EXO_XRTFLUJOS"" T3 on T0.""U_EXO_XRTFlujo""=T3.""Code"" "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""OCTG"" T1 on T0.""GroupNum""=T1.""GroupNum"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OPYM"" T2 on T0.""PeyMethod""=T2.""PayMethCod"" "
                    sSQL &= " WHERE T0.""CANCELED"" ='N'  AND T0.""DocStatus"" <>'C' and year(T0.""DocDueDate"")>=" & Now.Year.ToString("0000") & ") "
                    'Cliente
                    sSQL &= " UNION ALL "
                    sSQL &= "(Select '" + sBBDD + "' ""BD"",T0.""NumAtCard"" ""REFERENCIA"" ,'CPOR' ""Flujo"", T0.""U_EXO_XRTCPP"" ""PRESU"", T0.""DocDate"" ""F_OPE"", T0.""DocDueDate"" ""F_VTO"","
                    sSQL &= " round(T0.""DocTotalFC""- T0.""PaidFC"",2) ""ImpDIV"", T0.""DocCur"" ""DIV"",  round(T0.""DocTotal""-T0.""PaidToDate"",2) ""ImpT"", 'EUR' ""DIVEUR"", '" & sCtaFicticia & "' ""Cuenta"", "
                    sSQL &= " ifnull(T3.""Name"",'TRANSFERENCIA PROVEEDOR') ""TEXTO"", T0.""CardCode"" ""ZONA1"", T0.""DocNum"" ""ZONA2"", T2.""Descript"" ""ZONA3"", T1.""PymntGroup"", '' ""ZONA4"", "
                    sSQL &= " T0.""CardName"" ""ZONA5"",'OINV' ""Tabla"", T0.""ObjType"" ""OBJTYPE"" "
                    sSQL &= " FROM """ + sBBDD + """.""OINV"" T0 "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""@EXO_XRTFLUJOS"" T3 on T0.""U_EXO_XRTFlujo""=T3.""Code"" "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""OCTG"" T1 on T0.""GroupNum""=T1.""GroupNum"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OPYM"" T2 on T0.""PeyMethod""=T2.""PayMethCod"" "
                    sSQL &= " WHERE T0.""CANCELED"" ='N'  AND T0.""DocStatus"" <>'C' and T0.""DocTotal"">0 and year(T0.""DocDueDate"")>=" & Now.Year.ToString("0000") & ") "
                    sSQL &= " UNION ALL "
                    sSQL &= " (Select '" + sBBDD + "' ""BD"",T0.""NumAtCard"" ""REFERENCIA"" ,'CPOR' ""Flujo"", T0.""U_EXO_XRTCPP"" ""PRESU"", T0.""DocDate"" ""F_OPE"", T0.""DocDueDate"" ""F_VTO"","
                    sSQL &= " round(T0.""DocTotalFC""- T0.""PaidFC"",2) ""ImpDIV"", T0.""DocCur"" ""DIV"",  round(T0.""DocTotal""-T0.""PaidToDate"",2) ""ImpT"", 'EUR' ""DIVEUR"", '" & sCtaFicticia & "' ""Cuenta"", "
                    sSQL &= " ifnull(T3.""Name"",'TRANSFERENCIA PROVEEDOR') ""TEXTO"", T0.""CardCode"" ""ZONA1"", T0.""DocNum"" ""ZONA2"", T2.""Descript"" ""ZONA3"", T1.""PymntGroup"", '' ""ZONA4"", "
                    sSQL &= " T0.""CardName"" ""ZONA5"",'ORIN' ""Tabla"", T0.""ObjType"" ""OBJTYPE"" "
                    sSQL &= " FROM """ + sBBDD + """.""ORIN"" T0 "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""@EXO_XRTFLUJOS"" T3 on T0.""U_EXO_XRTFlujo""=T3.""Code"" "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""OCTG"" T1 on T0.""GroupNum""=T1.""GroupNum"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OPYM"" T2 on T0.""PeyMethod""=T2.""PayMethCod"" "
                    sSQL &= " WHERE T0.""CANCELED"" ='N'  AND T0.""DocStatus"" <>'C' and year(T0.""DocDueDate"")>=" & Now.Year.ToString("0000") & ") "
                Next
                sSQL &= ") T ORDER BY  T.""BD"", T.""F_OPE"" "
                Procesos.GestionarPrev_Origen(oDBSAP, oLog, sSQL)
            Else
                oLog.escribeMensaje("No existen sociedades definidas", EXO_Log.EXO_Log.Tipo.error)
                Exit Sub
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            Conexiones.Disconnect_SQLHANA(oDBSAP)
        End Try
    End Sub
    Public Shared Sub GestionarPrev_Origen(ByRef oDBSAP As HanaConnection, ByRef oLog As EXO_Log.EXO_Log, ByVal sSQL As String)
#Region "Variables"
        Dim sError As String = ""
        Dim odtDatos As System.Data.DataTable = Nothing
        Dim sPath As String = "" : Dim sRutaFich As String = "" : Dim sNomFich As String = ""
        Dim sLinea As String = ""
        Dim sSQLQuery As String = "" : Dim odtTabla As System.Data.DataTable = Nothing
        Dim sCodEmprXRT As String = ""
#End Region
        Try
            odtDatos = New System.Data.DataTable
            Conexiones.FillDtDB(oDBSAP, odtDatos, sSQL)
            If odtDatos.Rows.Count > 0 Then
                oLog.escribeMensaje("Recorriendo Facturas pendientes de envío...", EXO_Log.EXO_Log.Tipo.advertencia)
#Region "Comprobación de ruta para generar fichero y apertura"
                sPath = My.Application.Info.DirectoryPath.ToString

                If Not System.IO.Directory.Exists(sPath & "\PREV_ORIGEN") Then
                    System.IO.Directory.CreateDirectory(sPath & "\PREV_ORIGEN")
                End If
                sNomFich = "PREV_ORIGEN_" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "_" & Now.Hour.ToString("00") & Now.Minute.ToString("00")
                sRutaFich = Path.Combine(sPath & "\PREV_ORIGEN\" & sNomFich & ".txt")
                FileOpen(1, sRutaFich, OpenMode.Output)
                oLog.escribeMensaje("Generando fichero - " & sRutaFich, EXO_Log.EXO_Log.Tipo.advertencia)
#End Region
#Region "Generar Fichero"
                For Each dr As DataRow In odtDatos.Rows
                    'Tenemos que buscar el código de la empresa
                    'Si no existe no ponemos la línea
                    sSQLQuery = "SELECT ""U_EXO_XRTCOD"" ""CODXRT"" FROM """ + dr.Item("BD").ToString + """.""OADM""  "
                    odtTabla = New System.Data.DataTable
                    Conexiones.FillDtDB(oDBSAP, odtTabla, sSQLQuery)
                    If odtTabla.Rows.Count > 0 Then
                        sCodEmprXRT = odtTabla.Rows(0).Item("CODXRT").ToString
                    End If
                    If sCodEmprXRT <> "" Then
                        If CDbl(dr.Item("ImpT").ToString) > 0 Then
                            'sLinea = GENERALES.FormateaString(sCodEmprXRT, 4)
                            sLinea = GENERALES.FormateaString("", 4)
                            sLinea &= GENERALES.FormateaString(dr.Item("Cuenta").ToString, 10)
                            sLinea &= GENERALES.FormateaString(dr.Item("Flujo").ToString, 4)
                            sLinea &= GENERALES.FormateaString(dr.Item("PRESU").ToString, 10)
                            sLinea &= GENERALES.FormateaString(dr.Item("F_VTO").ToString, 10)
                            Dim dFecha As Date = dr.Item("F_VTO").ToString
                            If dFecha < Now.Date Then
                                Dim sfVto As String = Now.Date.AddDays(1).ToString
                                sLinea &= GENERALES.FormateaString(sfVto, 10)
                            Else
                                sLinea &= GENERALES.FormateaString(dr.Item("F_VTO").ToString, 10)
                            End If

                            If CDbl(IIf(dr.Item("ImpDIV").ToString = "", 0, dr.Item("ImpDIV").ToString)) = 0 Then
                                sLinea &= GENERALES.FormateaNumeroSinCeros(dr.Item("ImpT").ToString, 15, 2, True)
                                sLinea &= GENERALES.FormateaString(dr.Item("DIVEUR").ToString, 3)
                            Else
                                sLinea &= GENERALES.FormateaNumeroSinCeros(dr.Item("ImpDIV").ToString, 15, 2, True)
                                sLinea &= GENERALES.FormateaString(dr.Item("DIV").ToString, 3)
                            End If
                            sLinea &= GENERALES.FormateaNumeroSinCeros(dr.Item("ImpT").ToString, 15, 2, True)
                            sLinea &= GENERALES.FormateaString(dr.Item("DIVEUR").ToString, 3)
                            If dr.Item("Flujo").ToString <> "" Then
                                'Dim sTexto As String = Conexiones.GetValueDB(oDBSAP, """" & dr.Item("BD").ToString & """.""CUFD"" T0 INNER JOIN """ & dr.Item("BD").ToString & """.""UFD1"" T1 ON T0.""TableID""=T1.""TableID"" and T0.""FieldID""=T1.""FieldID""", "T1.""Descr""", " T0.""AliasID"" ='EXO_FlujoXRT' and T1.""FldValue""='" & dr.Item("Flujo").ToString & "'")
                                Dim sTexto As String = Conexiones.GetValueDB(oDBSAP, """" & dr.Item("BD").ToString & """.""@EXO_XRTFLUJOS""", """Name""", " ""Code""='" & dr.Item("Flujo").ToString & "'")
                                If sTexto = "" Then
                                    Select Case dr.Item("Flujo").ToString
                                        Case "CPOR" : sLinea &= GENERALES.FormateaString("Cobros Previsiones Origen", 30)
                                        Case "PPOR" : sLinea &= GENERALES.FormateaString("Pagos Previsiones Origen", 30)
                                        Case Else : sLinea &= GENERALES.FormateaString(dr.Item("TEXTO").ToString, 30)
                                    End Select
                                Else
                                    sLinea &= GENERALES.FormateaString(sTexto, 30)
                                End If
                            Else
                                sLinea &= GENERALES.FormateaString(dr.Item("TEXTO").ToString, 30)
                            End If
                            sLinea &= GENERALES.FormateaString(dr.Item("REFERENCIA").ToString, 10)
                            sLinea &= GENERALES.FormateaString(dr.Item("ZONA1").ToString, 30)
                            sLinea &= GENERALES.FormateaString(dr.Item("ZONA2").ToString, 30)
                            If dr.Item("ZONA3").ToString <> "" Then
                                sLinea &= GENERALES.FormateaString(dr.Item("ZONA3").ToString, 30)
                            Else
                                sLinea &= GENERALES.FormateaString(dr.Item("PymntGroup").ToString, 30)
                            End If
                            sLinea &= GENERALES.FormateaString(dr.Item("ZONA4").ToString, 30)
                            sLinea &= GENERALES.FormateaString(dr.Item("ZONA5").ToString, 30)
                            sLinea &= GENERALES.FormateaString(dr.Item("Tabla").ToString, 5)
                            sLinea &= GENERALES.FormateaString(dr.Item("OBJTYPE").ToString, 2)
                            PrintLine(1, sLinea)
                        End If
                    Else
                        oLog.escribeMensaje("En la BBDD: " & dr.Item("BD").ToString & " no existe el cód de empresa XRT", EXO_Log.EXO_Log.Tipo.error)
                    End If
                Next
                FileClose(1)
                oLog.escribeMensaje("Fichero Creado...", EXO_Log.EXO_Log.Tipo.informacion)
#End Region
#Region "Enviar por FTP"
                GENERALES.SubirFTP(sRutaFich, oLog, "PO")
#End Region

#Region "Guardar en el Hco"
                GENERALES.FicheroaHistorico(sPath & "\PREV_ORIGEN", sPath & "\PREV_ORIGEN" & "\HCOS", sNomFich, ".txt")
                oLog.escribeMensaje("El fichero fue movido al Hco: " & sPath & "\PREV_ORIGEN" & "\HCOS\" & sNomFich & ".txt", EXO_Log.EXO_Log.Tipo.advertencia)
#End Region
            Else
                oLog.escribeMensaje("No existen Facturas pendientes de envío.", EXO_Log.EXO_Log.Tipo.advertencia)
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            odtDatos = Nothing : odtTabla = Nothing
        End Try
    End Sub

    Public Shared Sub Prev_confirmadas(ByRef oLog As EXO_Log.EXO_Log)
#Region "Variables"
        Dim oDBSAP As HanaConnection = Nothing
        Dim sError As String = ""
        Dim sSQL As String = "" : Dim ESprimero As Boolean = True : Dim sBBDD As String = ""
        Dim odtDatos As System.Data.DataTable = Nothing
#End Region
        Try
            Conexiones.Connect_SQLHANA(oDBSAP, "HANA", oLog)
            'Buscamos las facturas pdtes de cobro
            sSQL = " SELECT ""EXO_BD"" FROM ""SOL_AUTORIZ"".""EXO_SOCIEDADES"" "
            odtDatos = New System.Data.DataTable
            Conexiones.FillDtDB(oDBSAP, odtDatos, sSQL)
            If odtDatos.Rows.Count > 0 Then
                oLog.escribeMensaje("Se va a proceder a recorrer las SOCIEDADES...", EXO_Log.EXO_Log.Tipo.advertencia)
                sSQL = "SELECT * FROM (" : ESprimero = True
                For Each dr As DataRow In odtDatos.Rows
                    sBBDD = dr.Item("EXO_BD").ToString

                    If ESprimero = False Then
                        sSQL &= " UNION ALL "
                    Else
                        ESprimero = False
                    End If
                    Dim sCtaFicticia As String = Conexiones.GetValueDB(oDBSAP, """" & sBBDD & """.""OADM""", """U_EXO_CTAXRT""", "")
                    oLog.escribeMensaje("La Cta. ficticia de la empresa " & sBBDD & " es:" & sCtaFicticia, EXO_Log.EXO_Log.Tipo.informacion)
                    If sCtaFicticia = "" Then
                        oLog.escribeMensaje("No se ha asignado. Por defecto se pondrá la CTA CTANOASIGN", EXO_Log.EXO_Log.Tipo.advertencia)
                        sCtaFicticia = "CTANOASIGN"
                    End If
                    sSQL &= " (Select DISTINCT '" + sBBDD + "' ""BD"",T0.""DocEntry"" ""REFERENCIA"" , T0.""DocDate"" ""F_OPE"", T0.""DocDueDate"" ""F_VTO"","
                    sSQL &= " T0.""DocTotalFC"" ""ImpDIV"", T0.""DocCurr"" ""DIV"",  T0.""DocTotal"" ""ImpT"", 'EUR' ""DIVEUR"", T2.""BankCode"", ifnull(T2.""U_EXO_XRTCOD"",'" & sCtaFicticia & "') ""Cuenta"", "
                    sSQL &= " T0.""JrnlMemo"", CASE T3.""InvType"" when 18 then TF.""NumAtCard"" WHEN 19 then TA.""NumAtCard"" ELSE 'SIN DOCUMENTO' END ""NumAtCard"", T0.""CardCode"" ""ZONA1"", T0.""DocNum"" ""ZONA2"", "
                    sSQL &= " T1.""Descript"" ""ZONA3"", T0.""TransId"" ""ZONA4"",T0.""CardName"" ""ZONA5"", CAST(T3.""DocEntry"" AS VARCHAR) ""Tabla"", T3.""InvType"" ""OBJTYPE"" "
                    sSQL &= " FROM """ + sBBDD + """.""OVPM"" T0  "
                    sSQL &= " INNER JOIN """ + sBBDD + """.""VPM2"" T3  on T0.""DocEntry""=T3.""DocNum"" and T3.""InvoiceId""=0 "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OPYM"" T1 on T0.""PayMth""=T1.""PayMethCod"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""PWZ3"" T4 on T3.""DocEntry""=T4.""InvKey"" And T3.""InvType""=T4.""ObjType"" "
                    sSQL &= " And T4.""IdEntry"" in (SELECT ""IdNumber"" FROM """ + sBBDD + """.""OPWZ"" WHERE ""Canceled""='N') "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""DSC1"" T2 on T4.""IBAN""=T2.""IBAN"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OPCH"" TF ON TF.""DocEntry""=T3.""DocEntry"" And TF.""ObjType""=T3.""InvType"" "
                    sSQL &= " Left JOIN """ + sBBDD + """.""ORPC"" TA ON TA.""DocEntry""=T3.""DocEntry"" And TA.""ObjType""=T3.""InvType"" "
                    sSQL &= " WHERE T0.""U_EXO_XRTE"" ='N'  and ifnull(CAST(T4.""InvKey"" as varchar),'')<>'' and T0.""Canceled""='N' and year(T0.""DocDueDate"")>=" & Now.Year.ToString("0000") & ") "
                    'Efectos,Confirming
                    sSQL &= " UNION ALL "
                    sSQL &= "(Select '" + sBBDD + "' ""BD"",BOE.""BoeNum"" ""REFERENCIA"",  J.""DueDate"" ""F_OPE"", T0.""DocDueDate"" ""F_VTO"",  "
                    sSQL &= " T0.""DocTotalFC"" ""ImpDIV"",  T0.""DocCurr"" ""DIV"", T0.""DocTotal"" ""ImpT"", 'EUR' ""DIVEUR"", T2.""BankCode"", ifnull(T2.""U_EXO_XRTCOD"",'" & sCtaFicticia & "') ""Cuenta"", "
                    sSQL &= " T0.""JrnlMemo"",CASE T3.""InvType"" when 18 then TF.""NumAtCard"" WHEN 19 then TA.""NumAtCard"" ELSE 'SIN DOCUMENTO' END ""NumAtCard"", T0.""CardCode"" ""ZONA1"", T0.""DocNum"" ""ZONA2"", "
                    sSQL &= " T1.""Descript"" ""ZONA3"", T0.""TransId"" ""ZONA4"",T0.""CardName"" ""ZONA5"", 'C' ""Tabla"", BOE.""BoeType"" ""OBJTYPE"" "
                    sSQL &= " FROM """ + sBBDD + """.""OVPM"" T0  "
                    sSQL &= " INNER JOIN  """ + sBBDD + """.""VPM2"" T3  On T0.""DocEntry""=T3.""DocNum"" And T3.""InvoiceId""=0 "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""OPYM"" T1 On T0.""PayMth""=T1.""PayMethCod"" "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""DSC1"" T2 On T1.""BnkDflt""=T2.""BankCode"" And T1.""DflAccount""=T2.""Account"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OCRD"" IC On T0.""CardCode""= IC.""CardCode"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OJDT"" O On T0.""TransId""=O.""TransId"" "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""JDT1"" J On O.""TransId""=J.""TransId"" "
                    sSQL &= " INNER JOIN """ + sBBDD + """.""OBOE"" BOE ON BOE.""BoeNum""=T0.""BoeNum"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OPCH"" TF On TF.""DocEntry""=T3.""DocEntry"" And TF.""ObjType""=T3.""InvType"" "
                    sSQL &= " Left JOIN """ + sBBDD + """.""ORPC"" TA On TA.""DocEntry""=T3.""DocEntry"" And TA.""ObjType""=T3.""InvType"" "
                    sSQL &= " WHERE T0.""Canceled""='N' and left(J.""Account"",4) in('4310','4311','4312','4110','4010') and (BOE.""BoeStatus""<>'C' and BOE.""BoeStatus""<>'P')  ) "
                    sSQL &= " UNION ALL "
                    sSQL &= "(Select '" + sBBDD + "' ""BD"",BOE.""BoeNum"" ""REFERENCIA"",  J.""DueDate"" ""F_OPE"", T0.""DocDueDate"" ""F_VTO"",  "
                    sSQL &= " T0.""DocTotalFC"" ""ImpDIV"",  T0.""DocCurr"" ""DIV"", T0.""DocTotal"" ""ImpT"", 'EUR' ""DIVEUR"", T2.""BankCode"", ifnull(T2.""U_EXO_XRTCOD"",'" & sCtaFicticia & "') ""Cuenta"", "
                    sSQL &= " CAST(BOE.""DepositNum"" AS VARCHAR) ,' '  ""NumAtCard"", IC.""CardCode"" ""ZONA1"", CAST(T0.""DocNum"" as VARCHAR) ""ZONA2"", "
                    sSQL &= " T1.""Descript"" ""ZONA3"", T0.""TransId"" ""ZONA4"",IC.""CardName"" ""ZONA5"", 'C' ""Tabla"", BOE.""BoeType"" ""OBJTYPE"" "
                    sSQL &= " FROM  """ + sBBDD + """.""ORCT"" T0 "
                    sSQL &= " INNER JOIN  """ + sBBDD + """.""RCT2"" T3  On T0.""DocEntry""=T3.""DocNum"" And T3.""InvoiceId""=0  "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""OPYM"" T1 On T0.""PayMth""=T1.""PayMethCod"" "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""DSC1"" T2 On T1.""BnkDflt""=T2.""BankCode"" And T1.""DflAccount""=T2.""Account"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OCRD"" IC On T0.""CardCode""= IC.""CardCode"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OJDT"" O On T0.""TransId""=O.""TransId"" "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""JDT1"" J On O.""TransId""=J.""TransId"" "
                    sSQL &= " INNER JOIN """ + sBBDD + """.""OBOE"" BOE On BOE.""BoeNum""=T0.""BoeNum"" "
                    sSQL &= " WHERE T0.""Canceled""='N' and (T0.""BoeStatus"" =  'G') and  (J.""BalDueDeb""+J.""BalDueCred"")<>0   ) "
                    sSQL &= " UNION ALL "
                    sSQL &= " (Select   '" + sBBDD + "' ""BD"",BOE.""BoeNum"" ""REFERENCIA"",  J.""DueDate"" ""F_OPE"", O.""DueDate""  ""F_VTO"",  "
                    sSQL &= " (J.""FCDebit"" + J.""FCCredit"")  ""ImpDIV"", DPS.""DeposCurr"" ""DIV"", (J.""Debit"" + J.""Credit"")  ""ImpT"", 'EUR' ""DIVEUR"", DPS.""BanckAcct"", ifnull(DPS.""DeposAcct"",'" & sCtaFicticia & "') ""Cuenta"", "
                    sSQL &= " CAST(BOE.""DepositNum"" AS VARCHAR), ' ' ""NumAtCard"", IC.""CardCode"" ""ZONA1"", BOE.""DepositNum"" ""ZONA2"", "
                    sSQL &= " BOE.""PayMethCod"" ""ZONA3"", J.""Line_ID""  ""ZONA4"",IC.""CardName"" ""ZONA5"", 'C' ""Tabla"", BOE.""BoeType"" ""OBJTYPE"" "
                    sSQL &= " FROM  """ + sBBDD + """.""OBOE"" BOE "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""ODPS"" DPS ON BOE.""DepositNum"" = DPS.""DeposId"" "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OJDT"" O On DPS.""TransAbs"" = O.""TransId"" "
                    sSQL &= " LEFT JOIN  """ + sBBDD + """.""JDT1"" J On O.""TransId""=J.""TransId""  And J.""ShortName"" = BOE.""CardCode"" And CAST(BOE.""BoeKey"" As varchar) = CAST(J.""Ref1"" As VARCHAR)  "
                    sSQL &= " LEFT JOIN """ + sBBDD + """.""OCRD"" IC On J.""ShortName""= IC.""CardCode"" "
                    sSQL &= " WHERE DPS.""Canceled""='N' and left(J.""Account"",4) in( '4311','4312') and (BOE.""BoeStatus"" =  'D') and  (J.""BalDueDeb""+J.""BalDueCred"")<>0  ) "
                Next
                sSQL &= ") T ORDER BY  T.""BD"", T.""Tabla"", T.""F_OPE"" "
                Procesos.GestionarPrev_confirmadas(oDBSAP, oLog, sSQL)
            Else
                oLog.escribeMensaje("No existen sociedades definidas", EXO_Log.EXO_Log.Tipo.error)
                Exit Sub
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            Conexiones.Disconnect_SQLHANA(oDBSAP)
        End Try
    End Sub
    Public Shared Sub GestionarPrev_confirmadas(ByRef oDBSAP As HanaConnection, ByRef oLog As EXO_Log.EXO_Log, ByVal sSQL As String)
#Region "Variables"
        Dim sError As String = ""
        Dim odtDatos As System.Data.DataTable = Nothing
        Dim sPath As String = "" : Dim sRutaFich As String = "" : Dim sNomFich As String = ""
        Dim sLinea As String = ""
        Dim sSQLQuery As String = "" : Dim odtTabla As System.Data.DataTable = Nothing
        Dim sCodEmprXRT As String = ""
        Dim sSQLAct As String = ""
        Dim sFlujo As String = "" : Dim sCodPptario As String = ""
        Dim sDocnum As String = ""
#End Region
        Try
            odtDatos = New System.Data.DataTable
            Conexiones.FillDtDB(oDBSAP, odtDatos, sSQL)
            If odtDatos.Rows.Count > 0 Then
                oLog.escribeMensaje("Recorriendo Pagos pendientes de envío...", EXO_Log.EXO_Log.Tipo.advertencia)
#Region "Comprobación de ruta para generar fichero y apertura"
                sPath = My.Application.Info.DirectoryPath.ToString

                If Not System.IO.Directory.Exists(sPath & "\PREV_CONFIRMADAS") Then
                    System.IO.Directory.CreateDirectory(sPath & "\PREV_CONFIRMADAS")
                End If
                sNomFich = "PREV_CONFIRMADAS_" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "_" & Now.Hour.ToString("00") & Now.Minute.ToString("00")
                sRutaFich = Path.Combine(sPath & "\PREV_CONFIRMADAS\" & sNomFich & ".txt")
                FileOpen(1, sRutaFich, OpenMode.Output)
                oLog.escribeMensaje("Generando fichero - " & sRutaFich, EXO_Log.EXO_Log.Tipo.advertencia)
#End Region
#Region "Crear Fichero"
                For Each dr As DataRow In odtDatos.Rows
                    'Tenemos que buscar el código de la empresa
                    'Si no existe no ponemos la línea
                    sSQLQuery = "SELECT ""U_EXO_XRTCOD"" ""CODXRT"" FROM """ + dr.Item("BD").ToString + """.""OADM""  "
                    odtTabla = New System.Data.DataTable
                    Conexiones.FillDtDB(oDBSAP, odtTabla, sSQLQuery)
                    If odtTabla.Rows.Count > 0 Then
                        sCodEmprXRT = odtTabla.Rows(0).Item("CODXRT").ToString
                    End If
                    If sCodEmprXRT <> "" Then
                        'Comprobamos que exista
                        If dr.Item("Cuenta").ToString.ToUpper.Contains("FICTICI") Then
                            oLog.escribeMensaje(dr.Item("BD").ToString & ": Banco: " & dr.Item("BankCode").ToString & ". No contiene Cta. Se ha indicado una ficticia. Cta: " & dr.Item("Cuenta").ToString & ". Asiento: " & dr.Item("ZONA2").ToString & ", Tipo: " & dr.Item("OBJTYPE").ToString, EXO_Log.EXO_Log.Tipo.error)
                        Else
                            'sLinea = GENERALES.FormateaString(sCodEmprXRT, 4)
                            'Buscamos el flujo de la factura para ello debemos buscar en la línea del pago cual es la factura.
                            Select Case dr.Item("OBJTYPE").ToString
                                Case "18"
                                    sFlujo = Conexiones.GetValueDB(oDBSAP, """" & dr.Item("BD").ToString & """.""OPCH""", """U_EXO_XRTFlujo""", """DocEntry""=" & dr.Item("Tabla").ToString)
                                    If sFlujo = "" Then
                                        sFlujo = "PVAR"
                                    End If
                                    sCodPptario = Conexiones.GetValueDB(oDBSAP, """" & dr.Item("BD").ToString & """.""OPCH""", """U_EXO_XRTCPP""", """DocEntry""=" & dr.Item("Tabla").ToString)
                                    If sCodPptario = "" Then
                                        sCodPptario = "PROV"
                                    End If
                                    sDocnum = Conexiones.GetValueDB(oDBSAP, """" & dr.Item("BD").ToString & """.""OPCH""", """DocNum""", """DocEntry""=" & dr.Item("Tabla").ToString)
                                Case "19"
                                    sFlujo = Conexiones.GetValueDB(oDBSAP, """" & dr.Item("BD").ToString & """.""ORPC""", """U_EXO_XRTFlujo""", """DocEntry""=" & dr.Item("Tabla").ToString)
                                    If sFlujo = "" Then
                                        sFlujo = "CVAR"
                                    End If
                                    sCodPptario = Conexiones.GetValueDB(oDBSAP, """" & dr.Item("BD").ToString & """.""ORPC""", """U_EXO_XRTCPP""", """DocEntry""=" & dr.Item("Tabla").ToString)
                                    If sCodPptario = "" Then
                                        sCodPptario = "CVAR"
                                    End If
                                    sDocnum = Conexiones.GetValueDB(oDBSAP, """" & dr.Item("BD").ToString & """.""ORPC""", """Docnum""", """DocEntry""=" & dr.Item("Tabla").ToString)
                                Case Else : sFlujo = "" : sCodPptario = "" : sDocnum = "" : oLog.escribeMensaje("Nº documento DocEntry: " & dr.Item("Tabla").ToString & ", con Tipo de Documento " & dr.Item("OBJTYPE").ToString & " No está contemplado.", EXO_Log.EXO_Log.Tipo.error)
                            End Select
                            If sFlujo = "" Then
                                oLog.escribeMensaje("Nº documento:" & sDocnum & ", con Tipo de Documento " & dr.Item("OBJTYPE").ToString & ", No se ha indicado el flujo. Se tendrá en cuenta en el envío en cuanto se indique.", EXO_Log.EXO_Log.Tipo.error)
                            Else
                                If CDbl(dr.Item("ImpT").ToString) > 0 Then
                                    sLinea = GENERALES.FormateaString("", 4)
                                    sLinea &= GENERALES.FormateaString(sFlujo, 4)
                                    sLinea &= GENERALES.FormateaString(dr.Item("Cuenta").ToString, 10)
                                    sLinea &= GENERALES.FormateaString(sCodPptario, 10)
                                    sLinea &= GENERALES.FormateaString(dr.Item("F_VTO").ToString, 10)
                                    Dim dFecha As Date = dr.Item("F_VTO").ToString
                                    'If dFecha < Now.Date Then
                                    '    Dim sfVto As String = Now.Date.AddDays(1).ToString
                                    '    sLinea &= GENERALES.FormateaString(sfVto, 10)
                                    'Else
                                    sLinea &= GENERALES.FormateaString(dr.Item("F_VTO").ToString, 10)
                                    'End If

                                    If CDbl(dr.Item("ImpDIV").ToString) = 0 Then
                                        sLinea &= GENERALES.FormateaNumeroSinCeros(dr.Item("ImpT").ToString, 15, 2, True)
                                        sLinea &= GENERALES.FormateaString(dr.Item("DIVEUR").ToString, 3)
                                    Else
                                        sLinea &= GENERALES.FormateaNumeroSinCeros(dr.Item("ImpDIV").ToString, 15, 2, True)
                                        sLinea &= GENERALES.FormateaString(dr.Item("DIV").ToString, 3)
                                    End If
                                    sLinea &= GENERALES.FormateaNumeroSinCeros(dr.Item("ImpT").ToString, 15, 2, True)
                                    sLinea &= GENERALES.FormateaString(dr.Item("DIVEUR").ToString, 3)
                                    'If sFlujo <> "" Then
                                    '    Dim sTexto As String = Conexiones.GetValueDB(oDBSAP, """" & dr.Item("BD").ToString & """.""@EXO_XRTFLUJOS""", """Name""", """Code""='" & sFlujo & "' ")
                                    '    If sTexto = "" Then
                                    '        sLinea &= GENERALES.FormateaString(Trim(dr.Item("JrnlMemo").ToString) & " - Asiento Nº " & dr.Item("ZONA2").ToString, 30)
                                    '    Else
                                    '        sLinea &= GENERALES.FormateaString(Trim(dr.Item("JrnlMemo").ToString) & " " & sTexto & " - Asiento Nº " & dr.Item("ZONA2").ToString, 30)
                                    '    End If
                                    'Else
                                    '    sLinea &= GENERALES.FormateaString(Trim(dr.Item("JrnlMemo").ToString) & " - Asiento Nº " & dr.Item("ZONA2").ToString, 30)
                                    'End If
                                    If sFlujo <> "" Then
                                        Dim sTexto As String = Conexiones.GetValueDB(oDBSAP, """" & dr.Item("BD").ToString & """.""@EXO_XRTFLUJOS""", """Name""", """Code""='" & sFlujo & "' ")
                                        If sTexto = "" Then
                                            sLinea &= GENERALES.FormateaString(Trim(dr.Item("NumAtCard").ToString) & " - Asiento " & dr.Item("ZONA2").ToString, 30)
                                        Else
                                            sLinea &= GENERALES.FormateaString(Trim(dr.Item("NumAtCard").ToString) & " " & sTexto & " - Asiento " & dr.Item("ZONA2").ToString, 30)
                                        End If
                                    Else
                                        sLinea &= GENERALES.FormateaString(Trim(dr.Item("NumAtCard").ToString) & " - Asiento " & dr.Item("ZONA2").ToString, 30)
                                    End If
                                    sLinea &= GENERALES.FormateaString(dr.Item("REFERENCIA").ToString, 10)
                                    sLinea &= GENERALES.FormateaString(dr.Item("ZONA1").ToString, 30)
                                    sLinea &= GENERALES.FormateaString(dr.Item("ZONA2").ToString, 30)
                                    sLinea &= GENERALES.FormateaString(dr.Item("ZONA3").ToString, 30)
                                    sLinea &= GENERALES.FormateaString(dr.Item("ZONA4").ToString, 30)
                                    sLinea &= GENERALES.FormateaString(dr.Item("ZONA5").ToString, 30)
                                    sLinea &= GENERALES.FormateaString(dr.Item("Tabla").ToString, 5)
                                    sLinea &= GENERALES.FormateaString(dr.Item("OBJTYPE").ToString, 2)
                                    PrintLine(1, sLinea)
                                    oLog.escribeMensaje("BBDD: " & dr.Item("BD").ToString & "  - Enviando Asiento:" & dr.Item("ZONA2").ToString & ", Tipo: " & dr.Item("OBJTYPE").ToString, EXO_Log.EXO_Log.Tipo.advertencia)
                                Else
                                    oLog.escribeMensaje("BBDD: " & dr.Item("BD").ToString & "  - No se envía Importe a 0 en Asiento:" & dr.Item("ZONA2").ToString & ", Tipo: " & dr.Item("OBJTYPE").ToString, EXO_Log.EXO_Log.Tipo.advertencia)
                                End If
                                sSQLAct = "UPDATE """ & dr.Item("BD").ToString & """.""OVPM"" "
                                sSQLAct &= " Set ""U_EXO_XRTE""='Y' "
                                sSQLAct &= " Where ""DocEntry""=" & dr.Item("REFERENCIA").ToString
                                Conexiones.ExecuteSqlDB(oDBSAP, sSQLAct)
                                oLog.escribeMensaje("BBDD: " & dr.Item("BD").ToString & "  - " & sSQLAct, EXO_Log.EXO_Log.Tipo.advertencia)
                            End If

                        End If
                    Else
                        oLog.escribeMensaje("En la BBDD: " & dr.Item("BD").ToString & " no existe el cód de empresa XRT", EXO_Log.EXO_Log.Tipo.error)
                    End If
                Next
                FileClose(1)
#End Region
                oLog.escribeMensaje("Fichero Creado...", EXO_Log.EXO_Log.Tipo.informacion)

#Region "Enviar por FTP"
                GENERALES.SubirFTP(sRutaFich, oLog, "PC")
#End Region

#Region "Guardar en el Hco"
                GENERALES.FicheroaHistorico(sPath & "\PREV_CONFIRMADAS", sPath & "\PREV_CONFIRMADAS" & "\HCOS", sNomFich, ".txt")
                oLog.escribeMensaje("El fichero fue movido al Hco: " & sPath & "\PREV_CONFIRMADAS" & "\HCOS\" & sNomFich & ".txt", EXO_Log.EXO_Log.Tipo.advertencia)
#End Region
            Else
                oLog.escribeMensaje("No existen Facturas pendientes de envío.", EXO_Log.EXO_Log.Tipo.advertencia)
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            odtDatos = Nothing : odtTabla = Nothing
        End Try
    End Sub
End Class
