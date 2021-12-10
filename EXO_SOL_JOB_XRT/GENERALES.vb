Imports System.IO
Imports WinSCP

Public Class GENERALES
#Region "Funciones formateos datos"
    Public Shared Function FormateaString(ByVal dato As Object, ByVal tam As Integer) As String
        Dim retorno As String = String.Empty

        If dato IsNot Nothing Then
            retorno = dato.ToString
        End If

        If retorno.Length > tam Then
            retorno = retorno.Substring(0, tam)
        End If
        retorno = retorno.Replace("-", "").Replace("*", "")
        Return retorno.PadRight(tam, CChar(" "))
    End Function
    Public Shared Function FormateaNumero(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""

        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            retorno = retorno.Replace("-", "N")
            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If

        Return retorno
    End Function
    Public Shared Function FormateaNumeroSinCeros(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""

        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            'retorno = retorno.Replace("-", "N")
            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar(" "))
        End If

        Return retorno
    End Function
    Public Shared Function FormateaNumeroSinPunto(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""
        Dim bEsNegativo As Boolean = False
        If Left(dato, 1) = "-" Then
            dato = dato.Replace("-", "")
            bEsNegativo = True
            posiciones = posiciones - 1
            totalNum = posiciones
        End If
        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            retorno = retorno.Replace(".", "")

            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If
        If bEsNegativo = True Then
            retorno = "N" & retorno
        End If
        Return retorno
    End Function
#End Region
#Region "FTP"
    Shared Sub UploadFTP(ByVal strFileNameLocal As String, ByRef oLog As EXO_Log.EXO_Log, ByVal sQueseenvia As String)
        Dim strNomCorto As String = Path.GetFileName(strFileNameLocal)
        Dim sURLFTP As String = Conexiones.Datos_Confi("FTP", "URL")
        Dim sPTFTP As String = Conexiones.Datos_Confi("FTP", "PUERTO")
        Dim sdirFTP As String = ""
        Select Case sQueseenvia
            Case "PO" : sdirFTP = Conexiones.Datos_Confi("FTP", "PREVORI")
        End Select
        Dim sUser As String = Conexiones.Datos_Confi("FTP", "US")
        Dim sPass As String = Conexiones.Datos_Confi("FTP", "PASS")
        Dim sEnviar As String = Conexiones.Datos_Confi("FTP", "ENVIAR")
        Dim clsRequest As System.Net.FtpWebRequest

        If sEnviar = "Y" Then
            'My.Computer.Network.UploadFile(strFileNameLocal, sURLFTP & sPTFTP & sdirFTP & strNomCorto, sUser, sPass, True, 100)
            'clsRequest = DirectCast(System.Net.WebRequest.Create(sURLFTP & sPTFTP & sdirFTP & strNomCorto), System.Net.FtpWebRequest)
            clsRequest = DirectCast(System.Net.WebRequest.Create(sURLFTP & "/" & sdirFTP & strNomCorto), System.Net.FtpWebRequest)
            'clsRequest.Proxy = Nothing ' Esta asignación es importantisimo con los que trabajen en windows XP ya que por defecto esta propiedad esta para ser asignado a un servidor http lo cual ocacionaria un error si deseamos conectarnos con un FTP, en windows Vista y el Seven no tube este problema.
            ' clsRequest.ImpersonationLevel = Security.Principal.TokenImpersonationLevel.Identification
            clsRequest.Credentials = New System.Net.NetworkCredential(sUser, sPass) ' Usuario y password de acceso al server FTP, si no tubiese, dejar entre comillas, osea ""
            clsRequest.Method = System.Net.WebRequestMethods.Ftp.UploadFile
            clsRequest.EnableSsl = True
            clsRequest.UsePassive = True
            clsRequest.KeepAlive = False
            clsRequest.AuthenticationLevel = System.Net.Security.AuthenticationLevel.MutualAuthRequested

            Try
                oLog.escribeMensaje("Subiendo fichero al FTP: " & sURLFTP & sPTFTP & sdirFTP & strNomCorto, EXO_Log.EXO_Log.Tipo.advertencia)
                Dim bFile() As Byte = System.IO.File.ReadAllBytes(strFileNameLocal)
                Dim clsStream As System.IO.Stream = clsRequest.GetRequestStream()
                clsStream.Write(bFile, 0, bFile.Length)
                clsStream.Close()
                clsStream.Dispose()
                oLog.escribeMensaje("Fichero subido...", EXO_Log.EXO_Log.Tipo.informacion)
            Catch ex As Exception
                oLog.escribeMensaje(ex.Message + ". El Archivo no pudo ser enviado, intente en otro momento.", EXO_Log.EXO_Log.Tipo.error)
            End Try
        Else
            oLog.escribeMensaje("No se sube el fichero: " & strNomCorto & " al FTP debido a la configuración del fichero CONNECTION", EXO_Log.EXO_Log.Tipo.advertencia)
        End If
    End Sub
    Shared Sub SubirFTP(ByVal strFileNameLocal As String, ByRef oLog As EXO_Log.EXO_Log, ByVal sQueseenvia As String)
        Dim strNomCorto As String = Path.GetFileName(strFileNameLocal)
        Dim sURLFTP As String = Conexiones.Datos_Confi("FTP", "URL")
        Dim sPTFTP As String = Conexiones.Datos_Confi("FTP", "PUERTO")
        Dim sdirFTP As String = ""
        Select Case sQueseenvia
            Case "PO" : sdirFTP = Conexiones.Datos_Confi("FTP", "PREVORI")
            Case "PC" : sdirFTP = Conexiones.Datos_Confi("FTP", "PREVCONF")
        End Select
        Dim sUser As String = Conexiones.Datos_Confi("FTP", "US")
        Dim sPass As String = Conexiones.Datos_Confi("FTP", "PASS")
        Dim sEnviar As String = Conexiones.Datos_Confi("FTP", "ENVIAR")

        If sEnviar = "Y" Then
            Try
                Dim sessionOptions As New SessionOptions
                With sessionOptions
                    .Protocol = Protocol.Ftp
                    .HostName = sURLFTP
                    .UserName = sUser
                    .Password = sPass
                    .FtpSecure = FtpSecure.Implicit
                End With

                Using session As New Session
                    ' Connect
                    session.Open(sessionOptions)

                    ' Upload files
                    Dim transferOptions As New TransferOptions
                    transferOptions.TransferMode = TransferMode.Binary

                    Dim transferResult As TransferOperationResult
                    transferResult = session.PutFiles(strFileNameLocal, sdirFTP, False, transferOptions)
                    'session.PutFiles(strFileNameLocal, "/home" & sdirFTP, False, transferOptions)


                    ' Throw on any error
                    transferResult.Check()

                    ' Print results
                    For Each transfer In transferResult.Transfers
                        oLog.escribeMensaje("Fichero subido..." & transfer.FileName, EXO_Log.EXO_Log.Tipo.informacion)
                    Next
                End Using

            Catch ex As Exception
                oLog.escribeMensaje(ex.Message + ". El Archivo no pudo ser enviado, intente en otro momento.", EXO_Log.EXO_Log.Tipo.error)
            End Try
        Else
            oLog.escribeMensaje("No se sube el fichero: " & strNomCorto & " al FTP debido a la configuración del fichero CONNECTION", EXO_Log.EXO_Log.Tipo.advertencia)
        End If
    End Sub

#End Region
#Region "Mover Fichero a Hco."
    Public Shared Sub FicheroaHistorico(ByVal folderPathOri As String, ByVal folderPathDes As String, ByVal file As String, ByVal sExtension As String)
        Try
            'Comprobamos que existe el directorio de destino
            If Not System.IO.Directory.Exists(folderPathDes) Then
                System.IO.Directory.CreateDirectory(folderPathDes)
            End If
            Dim FileDestino As String = ""
            FileDestino = file & "_" & Now.Hour.ToString("00") & Now.Minute.ToString("00") & Now.Second.ToString("00")
            My.Computer.FileSystem.MoveFile(folderPathOri & "\" & file & sExtension, folderPathDes & "\" & FileDestino & sExtension, True)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
End Class
