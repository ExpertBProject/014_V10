Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient

Public Class Conexiones

#Region "Connect to Company"

    Public Shared Sub Connect_Company(ByRef oCompany As SAPbobsCOM.Company, ByVal sDatabaseName As String, ByRef olog As EXO_Log.EXO_Log)
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing

        Try
            'Conectar DI SAP
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case "DI"
                                oCompany = New SAPbobsCOM.Company

                                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish
                                oCompany.Server = Reader.GetAttribute("Server").ToString.Trim
                                oCompany.LicenseServer = Reader.GetAttribute("LicenseServer").ToString.Trim
                                oCompany.UserName = Reader.GetAttribute("UserName").ToString.Trim
                                oCompany.Password = Reader.GetAttribute("Password").ToString.Trim
                                oCompany.UseTrusted = False
                                oCompany.DbPassword = Reader.GetAttribute("DbPassword").ToString.Trim
                                oCompany.DbUserName = Reader.GetAttribute("DbUserName").ToString.Trim
                                Dim sVersion As String = Reader.GetAttribute("Version").ToString.Trim
                                Select Case sVersion
                                    Case "2016" : oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
                                    Case "2017" : oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017
                                    Case "2018" : oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019
                                End Select

                                oCompany.CompanyDB = sDatabaseName

                                If oCompany.Connect <> 0 Then
                                    olog.escribeMensaje("Conexión Compañia " & sDatabaseName & ": " & oCompany.GetLastErrorDescription.Trim, EXO_Log.EXO_Log.Tipo.error)
                                    Throw New System.Exception("Error en la conexión a la compañia:" & oCompany.GetLastErrorDescription.Trim)
                                Else

                                End If
                        End Select
                End Select
            End While


        Catch exCOM As System.Runtime.InteropServices.COMException
            olog.escribeMensaje("Conexión Compañia: " & exCOM.ErrorCode & " - " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
            Throw exCOM
        Catch ex As Exception
            olog.escribeMensaje("Conexión SQL: " & ex.GetHashCode & " - " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub

    Public Shared Sub Disconnect_Company(ByRef oCompany As SAPbobsCOM.Company)
        Try
            If Not oCompany Is Nothing Then
                If oCompany.Connected = True Then
                    oCompany.Disconnect()
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oCompany IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCompany)
            oCompany = Nothing
        End Try
    End Sub

#End Region

#Region "Connect to SQL Server"

    Public Shared Sub Connect_SQLServer(ByRef db As SqlConnection, ByRef olog As EXO_Log.EXO_Log)
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing

        Try
            'Conectar SQL
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case "SQL"
                                If db Is Nothing OrElse db.State = ConnectionState.Closed Then
                                    Dim sCadConex As String = ""
                                    sCadConex = "Database=" & Reader.GetAttribute("Db").ToString.Trim & ";Data Source=" & Reader.GetAttribute("Server").ToString.Trim & ";User Id=" & Reader.GetAttribute("DbUser").ToString & ";Password=" & Reader.GetAttribute("DbPwd").ToString
                                    db = New SqlConnection
                                    db.ConnectionString = sCadConex
                                    olog.escribeMensaje("Conexión SQL: " & db.ConnectionString, EXO_Log.EXO_Log.Tipo.advertencia)
                                    db.Open()
                                    olog.escribeMensaje("Conectado", EXO_Log.EXO_Log.Tipo.advertencia)
                                End If

                        End Select
                End Select
            End While

        Catch exCOM As System.Runtime.InteropServices.COMException
            olog.escribeMensaje("Conexión SQL: " & exCOM.ErrorCode & " - " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
            Throw exCOM
        Catch ex As Exception
            olog.escribeMensaje("Conexión SQL: " & ex.GetHashCode & " - " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub

    Public Shared Sub Disconnect_SQLServer(ByRef db As SqlConnection)
        Try
            If Not db Is Nothing AndAlso db.State = ConnectionState.Open Then
                db.Close()
                db.Dispose()
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Shared Sub FillDtDB(ByRef db As SqlConnection, ByRef dt As System.Data.DataTable, ByVal strConsulta As String)
        Dim cmd As SqlCommand = Nothing
        Dim da As SqlDataAdapter = Nothing

        Try
            cmd = New SqlCommand(strConsulta, db)

            cmd.CommandTimeout = 0
            da = New SqlDataAdapter
            da.SelectCommand = cmd
            da.Fill(dt)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
            If Not da Is Nothing Then
                da.Dispose()
            End If
        End Try
    End Sub

    Public Shared Sub ExecuteSQLDB(ByRef db As SqlConnection, ByVal sSQL As String)
        Dim cmd As SqlCommand = Nothing

        Try
            cmd = New SqlCommand(sSQL, db)
            cmd.ExecuteNonQuery()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
        End Try
    End Sub

    Public Shared Sub ExecuteSQLDB(ByRef db As SqlConnection, ByRef oTransaction As SqlTransaction, ByVal sSQL As String)
        Dim cmd As SqlCommand = Nothing

        Try
            cmd = New SqlCommand(sSQL, db)
            cmd.Transaction = oTransaction
            cmd.ExecuteNonQuery()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
        End Try
    End Sub

    Public Shared Function GetValueDB(ByRef db As SqlConnection, ByRef sTabla As String, ByRef sCampo As String, ByRef sCondicion As String) As String
        Dim dt As System.Data.DataTable = Nothing
        Dim sSQL As String = ""
        Dim cmd As SqlCommand = Nothing
        Dim da As SqlDataAdapter = Nothing

        Try
            ''MyBase.ConnectSQLServer()

            If sCondicion = "" Then
                sSQL = "SELECT " & sCampo & " FROM " & sTabla
            Else
                sSQL = "SELECT " & sCampo & " FROM " & sTabla & " WHERE " & sCondicion
            End If

            dt = New System.Data.DataTable("Tabla")

            cmd = New SqlCommand(sSQL, db)
            cmd.CommandTimeout = 0

            da = New SqlDataAdapter

            da.SelectCommand = cmd
            da.Fill(dt)

            If dt.Rows.Count <= 0 Then
                Return ""
            Else
                If Not IsDBNull(dt.Rows.Item(0).Item(0).ToString) Then
                    Return dt.Rows.Item(0).Item(0).ToString
                Else
                    Return ""
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not dt Is Nothing Then
                dt.Dispose()
            End If
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
            If Not da Is Nothing Then
                da.Dispose()
            End If
        End Try
    End Function

    Public Shared Function GetValueDB(ByRef db As SqlConnection, ByRef oTransaction As SqlTransaction, ByRef sTabla As String, ByRef sCampo As String, ByRef sCondicion As String) As String
        Dim dt As System.Data.DataTable = Nothing
        Dim sSQL As String = ""
        Dim cmd As SqlCommand = Nothing
        Dim da As SqlDataAdapter = Nothing

        Try
            ''MyBase.ConnectSQLServer()

            If sCondicion = "" Then
                sSQL = "SELECT " & sCampo & " FROM " & sTabla
            Else
                sSQL = "SELECT " & sCampo & " FROM " & sTabla & " WHERE " & sCondicion
            End If

            dt = New System.Data.DataTable("Tabla")

            cmd = New SqlCommand(sSQL, db)
            cmd.Transaction = oTransaction
            cmd.CommandTimeout = 0

            da = New SqlDataAdapter

            da.SelectCommand = cmd
            da.Fill(dt)

            If dt.Rows.Count <= 0 Then
                Return ""
            Else
                If Not IsDBNull(dt.Rows.Item(0).Item(0).ToString) Then
                    Return dt.Rows.Item(0).Item(0).ToString
                Else
                    Return ""
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not dt Is Nothing Then
                dt.Dispose()
            End If
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
            If Not da Is Nothing Then
                da.Dispose()
            End If
        End Try
    End Function

#End Region

End Class
