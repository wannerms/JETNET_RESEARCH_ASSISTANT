Imports System.Net.Mail
Imports System.Windows.Forms
Imports System.IO
Imports MySql.Data.MySqlClient

Partial Public Class _Default2
    Inherits System.Web.UI.Page

    Dim INSPECTION_ID_ARRAY(100) As Integer
    Dim INSPECTION_DESC_ARRAY(100) As String
    Dim INSPECTION_HOURS_ARRAY(100) As Integer
    Dim INSPECTION_CYCLES_ARRAY(100) As Integer
    Dim INSPECTION_DATE_ARRAY(100) As String
    Dim ARRAY_COUNT As Integer = 0

    Dim SqlConn_JETNET As New SqlClient.SqlConnection
    Dim SqlCommand_JETNET As New SqlClient.SqlCommand
    Dim AircraftReader_JETNET As SqlClient.SqlDataReader

    Dim MySqlConn_AI As New MySql.Data.MySqlClient.MySqlConnection
    Dim MySqlCommand_AI As New MySql.Data.MySqlClient.MySqlCommand
    Dim MySqlReader_AI As MySql.Data.MySqlClient.MySqlDataReader
    Dim MySqlException As MySql.Data.MySqlClient.MySqlException

    Const JETNET_LIVE_SQL_CONN As String = "Data Source=172.30.5.58;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=Evolution;Password=k7F522#e"

    Const AI_LIVE_MYSQL_CONN As String = "Data Source=146.190.71.212;Initial Catalog=asset_insight;Persist Security Info=True;User ID=forge;Password=DSLEJY4KJ6Pb9IdL7u0D"


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Dim asset_id As String = ""
        Dim program_id As String = ""


        text_label.Text = ""
        asset_id = Trim(Request("asset"))
        program_id = Trim(Request("program"))

        If Trim(asset_id) <> "" And Trim(program_id) <> "" Then
            Call CREATE_AND_DISPLAY_AI_INSPECTIONS(asset_id, program_id)

        End If

    End Sub




    Public Sub CREATE_AND_DISPLAY_AI_INSPECTIONS(ByVal asset_id As String, ByVal program_id As String)

        Dim AI_MODEL_CODE As String = ""
        Dim model_name As String = ""
        Dim temp_return_value As Boolean = False
        Dim AI_DEFEAULT_INSPECTIONS As String = ""

        Try
            ' GO SELECT THE MODEL CODE BASED ON THE ASSET AND RETURN INTO AI_MODEL_CODE VARIABLE
            Call RETURN_AI_MODEL_CODE(AI_MODEL_CODE, asset_id, model_name)


            ' PASS IN THE VARIABLE WE JUST GOT AI_MODEL_CODE and GO GET AI_DEFAULT_INSPECTIONS
            Call GET_ASSET_INSIGHT_DEFAULT_INSPECTION_CODES(AI_MODEL_CODE, AI_DEFEAULT_INSPECTIONS)


            ' PASS THE DEFAULT INSPECTIONS IN SO THAT WE CAN SELECT THEM AND LOAD ALL OF THE INSPECTIONS INTO AN ARRAY 
            Call LOAD_MODEL_INSPECTIONS(asset_id, program_id, AI_DEFEAULT_INSPECTIONS)


            text_label.Text &= "<br>Asset ID: " & asset_id
            text_label.Text &= "<br>Program ID: " & program_id
            text_label.Text &= "<br>Model Name: " & model_name
            text_label.Text &= "<br>Model Code: " & AI_MODEL_CODE
            text_label.Text &= "<br>Default Inspections: " & AI_DEFEAULT_INSPECTIONS
            text_label.Text &= "<br>Inspections Found: " & ARRAY_COUNT
            text_label.Text &= "<br>------------------"
            For i = 0 To ARRAY_COUNT - 1

                If i = 0 Or i = 10 Or i = 11 Or Trim(INSPECTION_DESC_ARRAY(i)) = "1,600 Hour Inspection" Then
                    temp_return_value = CHECK_IF_EVOLUTION_INSPECTION_EXISTS(asset_id, INSPECTION_ID_ARRAY(i), "") ' this runs with a blank program id
                    If temp_return_value = True Then
                        temp_return_value = CHECK_IF_EVOLUTION_INSPECTION_EXISTS(asset_id, INSPECTION_ID_ARRAY(i), program_id) ' this then adds in the program id 

                        If temp_return_value = True Then  ' teste for a 2nd time, if its also true, all of the numbers line up 
                            text_label.Text &= "<br>--Inspection (<font color='green'>EXISTS - ID CORRECT " & INSPECTION_ID_ARRAY(i) & "</font>): " & " - " & INSPECTION_DESC_ARRAY(i)
                        Else ' if its false, then all of the numbers dont line up and the IDS are wrong on the program that is in there. inspection id correct, program id wrong 
                            text_label.Text &= "<br>--Inspection (<font color='green'>EXISTS</font> - <font color='red'>PROGRAM ID WRONG " & INSPECTION_ID_ARRAY(i) & "</font>) - " & INSPECTION_DESC_ARRAY(i)
                        End If
                    Else
                        text_label.Text &= "<br>--Inspections (<font color='red'>DOES NOT EXIST</font>): " & INSPECTION_ID_ARRAY(i) & " - " & INSPECTION_DESC_ARRAY(i)
                    End If

                Else
                    text_label.Text &= "<br>--Inspections: " & INSPECTION_DESC_ARRAY(i)
                End If


            Next
            text_label.Text &= "<br>------------------"

        Catch ex As Exception
            Response.Write("Error" & ex.ToString)
        End Try

    End Sub
    Public Sub GET_ASSET_INSIGHT_DEFAULT_INSPECTION_CODES(ByVal AI_MODEL_CODE As String, ByRef AI_DEFEAULT_INSPECTIONS As String)

        '// --10 planned hours And cycles
        '//--11 current hours And cycles
        '//--21 engine 1 -
        '//--22 engine 2
        '//-- 23 engine 3
        '//-- 24 engine 4
        '//-- 31 first apu
        '//--41 prop 1
        '//-- 42 prop 2
        '//   --  -- 2. based on the config, add inspection ids to inspection selection 

        Dim default_config_code As String = ""
        Dim engine_code As String = ""
        Dim prop_code As String = ""
        Dim apu_code As String = ""

        Try

            If Trim(AI_MODEL_CODE) <> "" Then

                If Len(Trim(AI_MODEL_CODE)) = 3 Then
                    default_config_code = "'10','11'" ' 10 is for PLANNED time, 11 is for CURRENT time

                    ' engine is first character, prop code is second and apu code is 3rd 
                    engine_code = Left(Trim(AI_MODEL_CODE), 1)
                    prop_code = Mid(Trim(AI_MODEL_CODE), 2, 1)
                    apu_code = Right(Trim(AI_MODEL_CODE), 1)

                    ' check the engine code to see how many engines there are 
                    If Trim(engine_code) = "0" Then
                        ' then it has no engines so dont any any defaults 
                    ElseIf Trim(engine_code) = "1" Then
                        default_config_code &= ",'21'"
                    ElseIf Trim(engine_code) = "2" Then
                        default_config_code &= ",'21','22'"
                    ElseIf Trim(engine_code) = "3" Then
                        default_config_code &= ",'21','22','23'"
                    ElseIf Trim(engine_code) = "4" Then
                        default_config_code &= ",'21','22','23','24'"
                    End If

                    ' check the prop code to see how many propellers there are 
                    If Trim(prop_code) = "0" Then
                        ' then it has no props so dont any any defaults 
                    ElseIf Trim(prop_code) = "1" Then
                        default_config_code &= ",'41'"
                    ElseIf Trim(prop_code) = "2" Then
                        default_config_code &= ",'41,'42'"
                    ElseIf Trim(prop_code) = "3" Then
                        default_config_code &= ",'41,'42,'43'"
                    ElseIf Trim(prop_code) = "4" Then
                        default_config_code &= ",'41,'42,'43,'44'"
                    End If

                    ' check the apu code to see how many apus there are 
                    If Trim(apu_code) = "0" Then
                        ' then it has no apu so dont any any defaults 
                    ElseIf Trim(apu_code) = "1" Then
                        default_config_code &= ",'31'"
                    ElseIf Trim(apu_code) = "2" Then
                        default_config_code &= ",'31','32'"
                    ElseIf Trim(apu_code) = "3" Then
                        default_config_code &= ",'31','32','33'"
                    ElseIf Trim(apu_code) = "4" Then
                        default_config_code &= ",'31','32','33','34'"
                    End If

                End If
            End If





        Catch ex As Exception
            Response.Write("Error" & ex.ToString)
        End Try

        AI_DEFEAULT_INSPECTIONS = default_config_code

    End Sub
    Public Sub RETURN_AI_MODEL_CODE(ByRef AI_MODEL_CODE As String, ByVal asset_id As String, ByRef model_name As String)
        Dim temp_ai_model_table As New DataTable

        ' take the moel id and go find both the model name and asset insight model code 
        temp_ai_model_table = Find_AI_Model_Code(asset_id)
        If temp_ai_model_table.Rows.Count > 0 Then

            For Each r As DataRow In temp_ai_model_table.Rows

                If Not IsDBNull(r.Item("config")) Then
                    If Not String.IsNullOrEmpty(r.Item("config").ToString.Trim) Then
                        AI_MODEL_CODE = r.Item("config")
                    End If
                End If

                If Not IsDBNull(r.Item("name")) Then
                    If Not String.IsNullOrEmpty(r.Item("name").ToString.Trim) Then
                        model_name = r.Item("name")
                    End If
                End If

            Next

        End If ' _dataTable.Rows.Count > 0 Then

    End Sub


    Public Function Find_AI_Model_Code(ByVal asset_id As String) As DataTable
        Dim Model_sQuery As String = ""
        Dim atemptable As New DataTable

        Try
            Find_AI_Model_Code = Nothing

            MySqlConn_AI.ConnectionString = AI_LIVE_MYSQL_CONN

            MySqlConn_AI.Open()

            Model_sQuery = "select config, name from models WHERE id = " & asset_id

            MySqlCommand_AI.Connection = MySqlConn_AI
            MySqlCommand_AI.CommandType = CommandType.Text
            MySqlCommand_AI.CommandTimeout = 60

            MySqlCommand_AI.CommandText = Model_sQuery
            MySqlReader_AI = MySqlCommand_AI.ExecuteReader()

            Try
                atemptable.Load(MySqlReader_AI)
            Catch constrExc As System.Data.ConstraintException
                Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
            End Try

            MySqlReader_AI.Dispose()

            Find_AI_Model_Code = atemptable

        Catch ex As Exception


            MySqlConn_AI.Dispose()
            MySqlCommand_AI.Dispose()

        Finally

            MySqlConn_AI.Close()
            MySqlCommand_AI.Dispose()
            MySqlConn_AI.Dispose()

        End Try

    End Function

    Public Sub LOAD_MODEL_INSPECTIONS(ByVal asset_id As String, ByVal program_id As String, ByVal ai_default_models As String)
        Dim temp_ai_model_table As New DataTable


        Dim i As Integer = 0

        For i = 0 To 99
            INSPECTION_ID_ARRAY(i) = 0
            INSPECTION_DESC_ARRAY(i) = ""
            INSPECTION_HOURS_ARRAY(i) = 0
            INSPECTION_CYCLES_ARRAY(i) = 0
            INSPECTION_DATE_ARRAY(i) = ""
        Next


        temp_ai_model_table = FIND_AI_MODEL_PROGRAM_INSPECTIONS(asset_id, program_id, ai_default_models)
        If Not IsNothing(temp_ai_model_table) Then
            If temp_ai_model_table.Rows.Count > 0 Then

                For Each r As DataRow In temp_ai_model_table.Rows

                    If Not IsDBNull(r.Item("id")) Then
                        If Not String.IsNullOrEmpty(r.Item("id").ToString.Trim) Then
                            INSPECTION_ID_ARRAY(ARRAY_COUNT) = r.Item("id")
                        End If
                    End If

                    If Not IsDBNull(r.Item("description")) Then
                        If Not String.IsNullOrEmpty(r.Item("description").ToString.Trim) Then
                            INSPECTION_DESC_ARRAY(ARRAY_COUNT) = r.Item("description")
                        End If
                    End If

                    If Not IsDBNull(r.Item("hours")) Then
                        If Not String.IsNullOrEmpty(r.Item("hours").ToString.Trim) Then
                            INSPECTION_HOURS_ARRAY(ARRAY_COUNT) = r.Item("hours")
                        End If
                    End If

                    If Not IsDBNull(r.Item("cycles")) Then
                        If Not String.IsNullOrEmpty(r.Item("cycles").ToString.Trim) Then
                            INSPECTION_CYCLES_ARRAY(ARRAY_COUNT) = r.Item("cycles")
                        End If
                    End If

                    If Not IsDBNull(r.Item("date")) Then
                        If Not String.IsNullOrEmpty(r.Item("date").ToString.Trim) Then
                            INSPECTION_DATE_ARRAY(ARRAY_COUNT) = r.Item("date")
                        End If
                    End If

                    ARRAY_COUNT += 1


                Next
            End If
        End If ' _dataTable.Rows.Count > 0 Then

    End Sub
    Public Function FIND_AI_MODEL_PROGRAM_INSPECTIONS(ByVal asset_id As String, ByVal program_id As String, ByVal ai_default_models As String) As DataTable
        Dim Inspection_sQuery As String = ""
        Dim atemptable As New DataTable

        Try
            FIND_AI_MODEL_PROGRAM_INSPECTIONS = Nothing

            MySqlConn_AI.ConnectionString = AI_LIVE_MYSQL_CONN

            MySqlConn_AI.Open()


            Inspection_sQuery = " select inspections.id,  inspection_descriptions.description, hours, cycles, date  from inspections"
            Inspection_sQuery &= " inner join inspection_descriptions on inspection_descriptions.id = inspections.inspection_description_id "
            Inspection_sQuery &= " where  ((program_id = " & program_id & " and required = 1) "

            ' add the codes in, so it should get more 
            If Trim(ai_default_models) <> "" Then
                Inspection_sQuery &= " or inspections.id in (" + ai_default_models + ") "
            End If
            Inspection_sQuery &= " ) "


            MySqlCommand_AI.Connection = MySqlConn_AI
            MySqlCommand_AI.CommandType = CommandType.Text
            MySqlCommand_AI.CommandTimeout = 60

            MySqlCommand_AI.CommandText = Inspection_sQuery
            MySqlReader_AI = MySqlCommand_AI.ExecuteReader()

            Try
                atemptable.Load(MySqlReader_AI)
            Catch constrExc As System.Data.ConstraintException
                Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
            End Try

            MySqlReader_AI.Dispose()

            FIND_AI_MODEL_PROGRAM_INSPECTIONS = atemptable

        Catch ex As Exception


            MySqlConn_AI.Dispose()
            MySqlCommand_AI.Dispose()

        Finally

            MySqlConn_AI.Close()
            MySqlCommand_AI.Dispose()
            MySqlConn_AI.Dispose()

        End Try

    End Function

    Public Function CHECK_IF_EVOLUTION_INSPECTION_EXISTS(ByVal asset_id As String, ByVal inspection_id As String, ByVal program_id As String) As Boolean
        Dim Model_sQuery As String = ""
        Dim atemptable As New DataTable
        Dim item_count As Integer = 0

        Try
            CHECK_IF_EVOLUTION_INSPECTION_EXISTS = Nothing

            SqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN

            SqlConn_JETNET.Open()

            Model_sQuery = "SELECT COUNT(*) AS itemcount FROM Asset_Insight_Model_Maintenance_Item "

            Model_sQuery &= " WHERE aimodmaint_item_id = " & inspection_id & " And aimodmaint_asset_id = " & asset_id

            If Trim(program_id) <> "" Then
                Model_sQuery &= " and aimodmaint_version_id = " & program_id
            End If

            SqlCommand_JETNET.Connection = SqlConn_JETNET
            SqlCommand_JETNET.CommandType = CommandType.Text
            SqlCommand_JETNET.CommandTimeout = 60

            SqlCommand_JETNET.CommandText = Model_sQuery
            AircraftReader_JETNET = SqlCommand_JETNET.ExecuteReader()

            Try
                atemptable.Load(AircraftReader_JETNET)
            Catch constrExc As System.Data.ConstraintException
                Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
            End Try

            AircraftReader_JETNET.Dispose()

            If atemptable.Rows.Count > 0 Then
                For Each r As DataRow In atemptable.Rows
                    If Not IsDBNull(r.Item("itemcount")) Then
                        If Not String.IsNullOrEmpty(r.Item("itemcount").ToString.Trim) Then
                            item_count = r.Item("itemcount")
                        End If
                    End If
                Next

                If item_count > 0 Then
                    CHECK_IF_EVOLUTION_INSPECTION_EXISTS = True
                End If
            End If



        Catch ex As Exception


            SqlConn_JETNET.Dispose()
            SqlCommand_JETNET.Dispose()

        Finally

            SqlConn_JETNET.Close()
            SqlCommand_JETNET.Dispose()
            SqlConn_JETNET.Dispose()

        End Try

    End Function

End Class