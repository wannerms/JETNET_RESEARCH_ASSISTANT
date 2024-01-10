Imports System.Net.Mail
Imports System.Windows.Forms
Imports System.IO

Partial Public Class _Default
  Inherits System.Web.UI.Page
  Dim TOTAL_NEWS As Integer = 0
  Dim TOTAL_COMPANIES As Integer = 0
  Dim TOTAL_COMPANIES_CONNECTED As Integer = 0
  Dim TOTAL_YACHTS_CONNECTED As Integer = 0
  Dim ac_found_count As Integer = 0
  Dim ac_count As Integer = 0
  Dim temp_details As String = ""
  Dim temp_spot As Integer = 0
  Dim temp_string As String = ""
  Dim is_due_date As Boolean = False
  Dim temp_spot2 As Integer = 0
  Dim temp_spot3 As Integer = 0
  Dim original_temp_details As String = ""
  Dim date_type As String = ""
  Dim is_found As Boolean = False
  Dim temp_ac_id As Long = 0
  Dim temp_date_sep As String = ""
  Dim length_plus_space As Integer = 0
  Dim x As Integer = 0
  Dim details_found_in_string As Boolean = False
  Dim error_count As Integer = 0
  Dim note_string As String = ""
  Dim ac_insert_count As Integer = 0
  Dim before_2000 As String = ""
  Dim TT_Count As Integer = 0
  Dim has_tt As Boolean = False
  Dim note_string_to_replace As String = ""
  Dim MySqlConn_JETNET As New SqlClient.SqlConnection
  Dim MySqlCommand_JETNET As New SqlClient.SqlCommand
  Dim MyAircraftReader_JETNET As SqlClient.SqlDataReader

  Dim MySqlConn_JETNET2 As New SqlClient.SqlConnection
  Dim MySqlCommand_JETNET2 As New SqlClient.SqlCommand
  Dim MyAircraftReader_JETNET2 As SqlClient.SqlDataReader
  Dim SqlConn_YPL As New SqlClient.SqlConnection
  Dim SqlCommand_YPL As New SqlClient.SqlCommand
  Dim SqlReader_YPL As SqlClient.SqlDataReader

    Dim temp_Country As String = ""
    Dim found_36_96 As String = ""
  Dim found_any_info As Boolean = False
  Dim inserted_any_info As Boolean = False
  Dim found_le_re As Boolean = False
  Dim found_possible_mismatch As Boolean = False
  Dim left_engine_string As String = ""
  Dim right_engine_string As String = ""
  Dim le_re_complied_or_due As String = ""
  Dim left_aftt As Integer = 0
  Dim right_aftt As Integer = 0
  Dim before_text As String = ""
  Dim aftt_string As String = ""
  Dim afft_found As Integer = 0
  Dim aftt_orig As String = ""
  Dim aftt_temp As String = ""
  Dim mis_count As Long = 0
  Dim mis_string As String = ""
  Dim clear_action_query As String = ""
  Dim temp_amod_id As Long = 0
  Dim temp_sale_flag As String = "N"
  Dim only_run_this_section As Boolean = False
  Dim yacht_news_name As String = ""
  Dim match_mmsi_count As Long = 0
  Dim total_mmsi As Long = 0
  Dim yt_count As Long = 0
  Dim new_yacht_mmsi As Long = 0

  Dim rows_match As String = "N"
  Dim rows_temp As String = ""
  Dim Match As Integer = 0
  Dim non_match As Integer = 0
  Dim not_found As Integer = 0
  Dim not_fs As Integer = 0
  Dim found_yt As Integer = 0
  Dim yt_table As String = ""
  Dim blank_mmsi_on_yacht As Integer = 0
  Dim wrong_ask As Integer = 0
  Dim more1_found As Integer = 0
  Dim yacht_id_sy As Long = 0
  Dim new_ypl_insert As Integer = 0
  Dim ypl_start_date As String
  Dim ypl_id As Long = 0
  Dim ypl_details As String = ""
  Dim ypl_link As String = ""
  Dim pub_yacht_id As Long = 0
  Dim found_pub_id As Long = 0
  Dim temp_where As String = ""
  Dim found_pub_match As Boolean = False
  Dim percent_off As Double = 0.03
  Dim yacht_asking_temp As String = ""
  Dim mmsi_string As String = ""
  Dim match_has_mmsi As Boolean = False
  Dim ys_mmsi As String = ""
  Dim ys_imo As String = ""
  Dim ys_dups As String = ""
  Dim dup_count As Integer = 0
  Dim bad_type As Integer = 0
  Dim pub_reg_no As String = ""
  Dim pub_landings As String = ""
  Dim pub_ser_no As String = ""
  Dim pub_desc As String = ""
  Dim pub_price As String = ""
  Dim pub_aftt As String = ""
    Dim pub_seller_info As String = ""
    Dim pub_seller_info_no_city As String = ""
    Dim pub_picture As String = ""
  Dim pub_status As String = ""
  Dim pub_url As String = ""
  Dim acpub_count As Long = 0
  Dim acpub_match_count As Long = 0
  Dim acpub_insert_count As Long = 0
  Dim acpub_original_name As String = ""
  Dim pub_comp_id As Integer = 0
  Dim acpub_process_status As String = ""
  Dim acpub_status As String = ""
  Dim acpub_controller_general_start As Integer = 0
  Dim total_pages As Long = 0
  Dim page_break_number As Long = 0
  Dim aftt_different As String = ""
  Dim landings_different As String = ""
  Dim has_pics As Boolean = False
  Dim Naughty_List_Of_Models(1000) As String
  Dim Naughty_List_Size As Integer = 0
  Dim acpub_price_details As String = ""
  Dim ypl_asking_price As String = ""
  Dim asking_within_range As Boolean = False
    Dim no_conn As String = "N"
    Dim na_skip As Boolean = False

    Dim zero_Count As Integer = 0
    Dim non_zero_Count As Integer = 0
    Dim pub_city As String = ""

    Dim xlApp As Microsoft.Office.Interop.Excel.Application
    Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
    'Dim xlBook_temp As Microsoft.Office.Interop.Excel.Workbook
    Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet

    Dim Insert_Query_Start As String = ""
    Dim Insert_Query1 As String = ""
    Dim temp_Date1 As String = ""

    Dim temp_make As String = ""
    Dim temp_model As String = ""
    Dim temp_party_first As String = ""
    Dim last_reg As String = ""
    Dim last_party As String = ""

    Dim pub_reg_no_doc_pending As String = ""


    ' Const JETNET_LIVE_SQL_CONN As String = "Data Source=www.jetnetsql2.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=sa;Password=krw32n89"
    '  Const JETNET_LIVE_SQL_CONN As String = "Data Source=www.jetnetsql1.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=sa;Password=krw32n89" 
    '  Const JETNET_LIVE_SQL_CONN As String = "Data Source=www.jetnetsql1.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=Evolution;Password=k7F522#e"

    Const JETNET_LIVE_SQL_CONN As String = "Data Source=172.30.5.58;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=Evolution;Password=k7F522#e"


    Const Inhouse_Live_Connection As String = "Data Source=10.10.254.54;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=sa;Password=moejive"
  Const Inhouse_Test_Connection As String = "Data Source=10.10.254.56;Initial Catalog=jetnet_ra_test;Persist Security Info=True;User ID=sa;Password=moejive"

    '''''' results = results & get_yacht_news_super_yacht_times()

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim temp_date As String = ""
        Dim sAdminReportString As String = ""
        Dim sReportOutputFilename As String = ""

        Try

            '  If Trim(Request("noconn")) = "Y" Then
            no_conn = "Y"
            '  End If

            text_label.Text = ""

            If Trim(Request("ae")) <> "" Then
                'Call Scrape_Aircraft_Exchange("https://aircraftexchange.com/jet-aircraft-for-sale/details/785/2013-dassault-falcon-900lx", 0)
                ' Call Scrape_Ebay_Page("https://av-info.faa.gov/data/AirOperators/fix/airopera.txt", 0)
                '  Call scrape_controller_html()



                '--------- not yet  

                '  Call Scrape_Aircraft_Exchange("https://aircraftexchange.com/aviation-brokers-dealers", 0)


                '    Call Run_AC_Exchange()

                Call Scrape_For_Business_Air("https://www.findaircraft.com/aircraft-for-sale/?AircraftType=Bombardier%20Challenger%20300")

                Call Scrape_For_Business_Air("https://www.businessair.com/aircraft?page=1")  ' beechcraft 

                '   Call Scrape_For_Business_Air("https://www.businessair.com/taxonomy/term/17726/all")  ' beechcraft 

                '  Call Scrape_For_flightmarket("https://www.flightmarket.com.br/br")



                ' --- NOT WORKING-----------
                '  Call Scrape_For_global_plane_search("https://www.globalplanesearch.com/aircraft-for-sale-worldwide/?sort=age_d")




                'working --------------
                '  Call Scrape_Ebay_Page("https://av-info.faa.gov/data/AirOperators/fix/airopera.txt", 0)
                ' Call Scrape_Ebay_Page("https://aircraftexchange.com/jet-aircraft-for-sale/details/245/2000-id76manufacturer-id1namecitation-exceldeleted-atnullcreated-at2018-07-05-023905updated-at2018-07-05-023905typejet-for-sale", 0)
                'Call Scrape_Ebay_Page("https://aircraftexchange.com/aircraft-by-broker/115/qs-partners", 0)

                '  Call Scrape_For_flightmarket("https://www.flightmarket.com.br/pt/aeronaves")
                ' Call Scrape_For_aviapages("https://aviapages.com/jet_market/")


            End If





            If Trim(Request("rep_id")) <> "" Then

                Call generateAdminReport(Trim(Request("rep_id")), sAdminReportString, 0, "", False, False, False)

                Dim sReportTitle = "adminReport_" & Trim(Request("rep_id"))
                Dim sAdminReportFileName As String = ""
                sAdminReportFileName = GenerateFileName(sReportTitle, ".xls", False)

                If write_report_string_to_file(sAdminReportString, sAdminReportFileName) Then
                    sReportOutputFilename = "pictures/" + sAdminReportFileName.Trim
                Else
                    HttpContext.Current.Session.Item("localUser").crmUser_DebugText += "Error in sending string to file"
                End If

                Dim reportURL As String = "openReportWindow(""" + sReportOutputFilename.Trim + """,""" & Trim(Request("rep_id")) & """);"
                System.Web.UI.ScriptManager.RegisterStartupScript(Me, Me.GetType(), "PopUpSelectedReportWindow", reportURL, True)
            End If


            If Trim(Request("answer")) = "Y" Then
                If Trim(Request("yacht_news")) = "Y" Then
                    Call YACHT_NEWS()
                ElseIf Trim(Request("excel")) = "Y" Then
                    Call run_excel_scraper()
                ElseIf Trim(Request("reserved")) = "Y" Then
                    run_excel_scraper_reserved()
                ElseIf Trim(Request("yacht_pub")) = "Y" Then
                    Call YACHT_PUB(Trim(Request("page")))
                ElseIf Trim(Request("jetnet_news")) = "Y" Then
                    Call JETNET_NEWS()
                ElseIf Trim(Request("mmsi")) = "Y" Then
                    Call get_MMSI_NUMBERS_top_function()
                ElseIf Trim(Request("ac_pub")) = "Y" Then
                    Call get_ac_pub_top_function()
                ElseIf Trim(Request("ac_pub_controller")) = "Y" Then

                    Call scrape_controller_html()

                ElseIf Trim(Request("data_integrity")) = "Y" Then
                    Call RUN_DATA_INTEGRITY_CHECKS()
                ElseIf Trim(Request("ac_details")) = "Y" Then
                    ' THIS HAS BEEN COMMENTED OUT TO AVOID ACCIDENT OF RE-RUN
                    ' get_AC_Details()
                    text_label.Text &= "<br><Br><br>Back to <a href='default.aspx'>Home</a>"
                ElseIf Trim(Request("ac_details")) = "S" Then
                    '   only_run_this_section = True
                    '   get_AC_Details()
                    '   text_label.Text &= "<br><Br><br>Back to <a href='default.aspx'>Home</a>"
                ElseIf Trim(Request("euro")) = "Y" Then
                    Call scrape_for_Euro_Control()
                End If
            ElseIf Trim(Request("question")) = "Y" Then
                If Trim(Request("yacht_news")) = "Y" Then
                    text_label.Text &= "Are You Sure You Want to Run the Yacht News?"
                    text_label.Text &= "<br><br><a href='default.aspx?yacht_news=Y&answer=Y'>Yes</a>"
                    text_label.Text &= "<br><br><a href='default.aspx?yacht_news=Y&answer=N'>No</a>"
                ElseIf Trim(Request("excel_question")) = "Y" Then
                    text_label.Text &= "Are You Sure You Want to Run the Doc Request?"
                    text_label.Text &= "<br><br><a href='default.aspx?excel=Y&answer=Y'>Yes</a>"
                    text_label.Text &= "<br><br><a href='default.aspx?excel=Y&answer=N'>No</a>"
                ElseIf Trim(Request("reserved_question")) = "Y" Then
                    text_label.Text &= "Are You Sure You Want to Run the Reserved List?"
                    text_label.Text &= "<br><br><a href='default.aspx?reserved=Y&answer=Y'>Yes</a>"
                    text_label.Text &= "<br><br><a href='default.aspx?reserved=Y&answer=N'>No</a>"
                ElseIf Trim(Request("ac_pub")) = "Y" Then
                    text_label.Text &= "Are You Sure You Want to Run the AC Pubs?"
                    text_label.Text &= "<br><br><a href='default.aspx?ac_pub=Y&answer=Y'>Yes</a>"
                    text_label.Text &= "<br><br><a href='default.aspx?ac_pub=Y&answer=N'>No</a>"
                    text_label.Text &= "<br><br><br><br><br><br><a href='default.aspx?ac_pub_controller=Y&answer=Y'>Controller</a>"

                ElseIf Trim(Request("data_integrity")) = "Y" Then
                    text_label.Text &= "Are You Sure You Want to Run the Data Integrity?"
                    text_label.Text &= "<br><br><a href='default.aspx?data_integrity=Y&answer=Y'>Yes</a>"
                    text_label.Text &= "<br><br><a href='default.aspx?data_integrity=Y&answer=N'>No</a>"
                ElseIf Trim(Request("yacht_pub")) = "Y" Then
                    text_label.Text &= "Are You Sure You Want to Run the Yacht Pubs?"
                    text_label.Text &= "<br><br><a href='default.aspx?yacht_pub=Y&answer=Y&page=1'>Yes - Page 1</a>"
                    text_label.Text &= "<br><br><a href='default.aspx?yacht_pub=Y&answer=Y&page=2'>Yes - Page 2</a>"
                    text_label.Text &= "<br><br><a href='default.aspx?yacht_pub=Y&answer=N'>No</a>"
                ElseIf Trim(Request("jetnet_news")) = "Y" Then
                    text_label.Text &= "Are You Sure You Want to Run the JETNET News?"
                    text_label.Text &= "<br><br><a href='default.aspx?jetnet_news=Y&answer=Y'>Yes</a>"
                    text_label.Text &= "<br><br><a href='default.aspx?jetnet_news=Y&answer=N'>No</a>"
                ElseIf Trim(Request("mmsi")) = "Y" Then
                    text_label.Text &= "Are You Sure You Want to Run the Yacht Marine Traffic MMSI?"
                    text_label.Text &= "<br><br><a href='default.aspx?mmsi=Y&answer=Y'>Yes</a>"
                    text_label.Text &= "<br><br><a href='default.aspx?mmsi=Y&answer=N'>No</a>"
                ElseIf Trim(Request("ac_details")) = "Y" Then
                    text_label.Text &= "Are You Sure You Want to Run the AC Maintenance Details Page?"
                    text_label.Text &= "<br><br><a href='default.aspx?ac_details=Y&answer=Y'>Yes</a>"
                    text_label.Text &= "<br><br><a href='default.aspx?ac_details=Y&answer=N'>No</a>"

                    text_label.Text &= "<br><br><a href='default.aspx?ac_details=S&answer=Y'>Run Selected Few Items Only</a>"
                End If
            Else

                ' If Trim(Request("assett_insight")) = "Y" Then
                '  Call asset_insight_functions() 
                'End If

                'text_label.Text &= "<b class=""aircraftHead"">Research Assistant</b>"
                text_label.Text &= "<table cellspacing='0' cellpadding='4' border='1'><tr valign='top'>"
                text_label.Text &= "<td align='left' colspan='2'><b class=""aircraftSubHead"">AIRCRAFT</b></td>"
                text_label.Text &= "<td align='left' colspan='2'><b class=""yachtSubHead"">YACHTS</b></td>"
                text_label.Text &= "</tr>"

                text_label.Text &= "<tr valign='top'>"

                If Trim(no_conn) = "Y" Then
                Else
                    temp_date = GET_LOG_DATE("Aircraft News Started", "N")
                End If
                text_label.Text &= "<td align='left' valign='top' class=""fa_size""><a href='default.aspx?jetnet_news=Y&question=Y'><i class=""fa fa-newspaper-o"" aria-hidden=""true""></i><!--<img align='texttop' src='/pictures/ytnews.jpg' width='70'>--></a></td><td align='left'><a href='default.aspx?jetnet_news=Y&question=Y'>Aircraft News.</a><br/>Last Run on " & temp_date & "</td>"


                If Trim(no_conn) = "Y" Then
                Else
                    temp_date = GET_LOG_DATE("Yacht News Started", "Y")
                End If
                text_label.Text &= "<td align='left' valign='top' class=""fa_size""><a href='default.aspx?yacht_news=Y&question=Y'><i class=""fa fa-newspaper-o"" aria-hidden=""true""></i><!--<img align='texttop' src='/pictures/ytnews.jpg' width='70'>--></a></td><td align='left'><a href='default.aspx?yacht_news=Y&question=Y'>Yacht News</a><br/>Last Run on " & temp_date & "</td>"
                text_label.Text &= "</tr>"


                If Trim(no_conn) = "Y" Then
                Else
                    temp_date = GET_LOG_DATE("Aircraft Pubs Started", "Y")
                End If
                text_label.Text &= "<tr valign='top'>"
                ' text_label.Text &= "<td align='left' valign='top'><a href='default.aspx?ac_pub=Y&question=Y'><img align='texttop' src='/pictures/for_sale.jpg' width='70'></a></td><td align='left'><a href='default.aspx?ac_pub=Y&question=Y'>Aircraft Sale Listings</a></td>"
                text_label.Text &= "<td align='left' valign='top'><a href='default.aspx?ac_pub=Y&question=Y'><i class=""fa fa-tachometer"" aria-hidden=""true""></i><!--<img align='texttop' src='/pictures/for_sale.jpg' width='70'>--></a></td><td><a href='default.aspx?ac_pub=Y&question=Y'>Aircraft Pubs</a><br/>Last Run on " & temp_date & "</td>"


                If Trim(no_conn) = "Y" Then
                Else
                    temp_date = GET_LOG_DATE("Yacht Pubs Started", "Y")
                End If
                text_label.Text &= "<td align='left' valign='top'><a href='default.aspx?yacht_pub=Y&question=Y'><i class=""fa fa-money"" aria-hidden=""true""></i><!--<img align='texttop' src='/pictures/for_sale.jpg' width='70'>--></a></td><td align='left'><a href='default.aspx?yacht_pub=Y&question=Y'>Yacht For Sale Listings</a><br/>Last Run on " & temp_date & "</td>"
                text_label.Text &= "</tr>"


                temp_date = ""
                text_label.Text &= "<tr valign='top'>"
                text_label.Text &= "<td align='left' valign='top'><a href='default.aspx?data_integrity=Y&question=Y'><i class=""fa fa-money"" aria-hidden=""true""></i><!--<img align='texttop' src='/pictures/for_sale.jpg' width='70'>--></a></td><td align='left'><a href='default.aspx?data_integrity=Y&question=Y'>Data Integrity</a><br/>Last Run on " & temp_date & "</td>"
                ' text_label.Text &= "<td align='left' valign='top'>&nbsp;</td><td align='left'><a href='default.aspx?ac_details=Y&question=Y'>Aircraft Maintenance Details</a></td>"




                temp_date = ""
                If Trim(no_conn) = "Y" Then
                Else
                    temp_date = GET_LOG_DATE("Doc Request Started", "Y")
                End If
                text_label.Text &= "<td align='left' valign='top'><a href='default.aspx?excel_question=Y&question=Y'><i class=""fa fa-money"" aria-hidden=""true""></i><!--<img align='texttop' src='/pictures/for_sale.jpg' width='70'>--></a></td><td align='left'><a href='default.aspx?excel_question=Y&question=Y'>Doc Request</a><br/>Last Run on " & temp_date & "</td>"
                ' text_label.Text &= "<td align='left' valign='top'>&nbsp;</td><td align='left'><a href='default.aspx?ac_details=Y&question=Y'>Aircraft Maintenance Details</a></td>"






                text_label.Text &= "</tr>"

                text_label.Text &= "<tr valign='top'>"

                temp_date = ""
                If Trim(no_conn) = "Y" Then
                Else
                    temp_date = GET_LOG_DATE("Doc Request Started", "Y")
                End If
                text_label.Text &= "<td align='left' valign='top'><a href='default.aspx?reserved_question=Y&question=Y'><i class=""fa fa-money"" aria-hidden=""true""></i><!--<img align='texttop' src='/pictures/for_sale.jpg' width='70'>--></a></td><td align='left'><a href='default.aspx?reserved_question=Y&question=Y'>Reserved Data</a><br/>Last Run on " & temp_date & "</td>"



                text_label.Text &= "</tr>"


                ' text_label.Text &= "<td align='left' valign='top'><i class=""fa fa-anchor"" aria-hidden=""true""></i></td><td align='left'><a href='default.aspx?mmsi=Y&question=Y'>Yacht Marine Traffic MMSI</a></td>"



                'text_label.Text &= "<tr valign='top'>"
                'text_label.Text &= "<td align='left' valign='top'>&nbsp;</td><td align='left'></td>"
                'text_label.Text &= "<td align='left' valign='top'>&nbsp;</td><td align='left'>&nbsp;</td>"
                'text_label.Text &= "</tr>"
                If Trim(no_conn) = "Y" Then
                Else
                    Call GET_DATA_INTEGRITY_INFO()
                End If

                text_label.Text &= "</table>"


            End If



            ' temp_date = GET_LAST_DATE("ABI_News_Links", "abinewslnk_date", "", "N")
            '  temp_date = GET_LOG_DATE("Yacht_News", "ytnews_action_date", "", "Y") 
            ' temp_date = GET_LAST_DATE("Temp_Publication_Log", "temp_publog_entry_date", "", "Y")
            'temp_date = GET_LAST_DATE("Yacht_Publication_Log", "ypl_source_date", "", "Y")
            'temp_date = GET_LAST_DATE("ABI_News_Links_Temp", "tmpnewslnk_date", " tmpnewslnk_date <= getdate() ", "N")


        Catch ex As Exception
            Response.Write(ex)
        Finally
            MySqlConn_JETNET = Nothing
            MyAircraftReader_JETNET = Nothing
            MySqlCommand_JETNET = Nothing
        End Try


    End Sub

    Public Sub run_excel_scraper()

        Dim temp_directory As String = "C:\Users\Matt Wanner\Desktop\"
        '   Dim temp_directory As String = "D:\jetnetassistant\DOC_REQUEST\"

        '    Dim temp_file_name As String = "doc_index.html"
        Dim temp_file_name As String = "doc_index.xlsx"
        '  Dim temp_file_name As String = "doc_index.csv"
        Dim temp_line As Long = 0
        Dim temp_String As String = ""

        Dim temp_party As String = ""
        Dim temp_party_All As String = ""
        Dim temp_serial As String = ""
        Dim last_serial As String = ""
        Dim last_date As String = ""


        Dim ac_id As Long = 0
        Dim publist_research_note As String = ""
        Dim Insert_Record As String = ""





        Insert_Query_Start = " INSERT INTO Publication_Listing"
        Insert_Query_Start &= " (publist_ac_id"
        Insert_Query_Start &= " ,publist_journ_id"
        Insert_Query_Start &= " ,publist_source"
        Insert_Query_Start &= " ,publist_reg_no"
        Insert_Query_Start &= " ,publist_ser_no"
        Insert_Query_Start &= " ,publist_description"
        Insert_Query_Start &= " ,publist_price"
        Insert_Query_Start &= " ,publist_aftt"
        Insert_Query_Start &= " ,publist_seller_info"
        Insert_Query_Start &= " ,publist_picture"
        Insert_Query_Start &= " ,publist_status"
        Insert_Query_Start &= " ,publist_url"
        Insert_Query_Start &= "  ,publist_clear_date"
        Insert_Query_Start &= "  ,publist_acct_rep"
        Insert_Query_Start &= "  ,publist_entry_date"
        Insert_Query_Start &= "  ,publist_update_date"
        Insert_Query_Start &= "  ,publist_original_desc"
        Insert_Query_Start &= "  ,publist_latest_change"
        Insert_Query_Start &= "  ,publist_user_id"
        Insert_Query_Start &= "  , publist_type "
        Insert_Query_Start &= "  , publist_comp_id "
        Insert_Query_Start &= "  ,publist_process_status, publist_research_note, publist_category)"
        Insert_Query_Start &= " VALUES( "



        Try

            MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
            '  MySqlConn_JETNET.ConnectionString = Inhouse_Test_Connection
            ' MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN
            MySqlConn_JETNET.Open()
            MySqlCommand_JETNET.Connection = MySqlConn_JETNET
            MySqlCommand_JETNET.CommandType = CommandType.Text
            MySqlCommand_JETNET.CommandTimeout = 60


        Catch ex As Exception

        End Try

        Call insert_into_eventlog("Doc Request Started", "Research Assistant")

        'Dim string_text As String = "" 
        'Using sr As New StreamReader(temp_directory & temp_file_name)
        '    '  Using sr As New StreamReader("C:\Controller\" & i & ".htm")
        '    string_text = sr.ReadToEnd()

        '    string_text = string_text
        'End Using



        xlApp = CType(CreateObject("Excel.Application"),
                  Microsoft.Office.Interop.Excel.Application)
        xlBook = CType(xlApp.Workbooks.Open(temp_directory & temp_file_name), Microsoft.Office.Interop.Excel.Workbook)

        ' xlBook = CType(xlApp.Workbooks.Add,Microsoft.Office.Interop.Excel.Workbook)
        xlSheet = CType(xlBook.Worksheets(1),
            Microsoft.Office.Interop.Excel.Worksheet)



        Dim xRng As Microsoft.Office.Interop.Excel.Range
        Dim val As Object

        ' no line 0 , and line 1 is columns 
        For temp_line = 2 To 10000
            'temp_String = xlSheet.Cells(temp_line, 1).value()
            temp_String = temp_String

            ' COLUMN 1 is N-NUMBER
            xRng = CType(xlSheet.Cells(temp_line, 1), Microsoft.Office.Interop.Excel.Range)
            val = xRng.Value()
            If Not IsNothing(val) Then
                temp_String = val.ToString
            Else
                temp_String = ""
            End If
            pub_reg_no = temp_String
            pub_reg_no = Replace(pub_reg_no, """", "")


            ' SERIAL NUMBER
            xRng = CType(xlSheet.Cells(temp_line, 2), Microsoft.Office.Interop.Excel.Range)
            val = xRng.Value()
            If Not IsNothing(val) Then
                temp_serial = val.ToString
            Else
                temp_serial = ""
            End If
            pub_ser_no = temp_serial
            pub_ser_no = Replace(pub_ser_no, """", "")

            If Trim(temp_String) = "" And Trim(temp_serial) = "" Then
                temp_line = 10001
            Else

                ' MFR
                xRng = CType(xlSheet.Cells(temp_line, 36), Microsoft.Office.Interop.Excel.Range)
                val = xRng.Value()
                temp_make = val.ToString

                ' Model
                xRng = CType(xlSheet.Cells(temp_line, 37), Microsoft.Office.Interop.Excel.Range)
                val = xRng.Value()
                temp_model = val.ToString


                ' DR DATE 
                xRng = CType(xlSheet.Cells(temp_line, 42), Microsoft.Office.Interop.Excel.Range)
                val = xRng.Value()
                temp_Date1 = val.ToString

                ' country
                xRng = CType(xlSheet.Cells(temp_line, 15), Microsoft.Office.Interop.Excel.Range)
                If Not IsNothing(xRng.Value()) Then
                    val = xRng.Value()
                    temp_Country = val.ToString
                Else
                    temp_Country = ""
                End If



                ' Party
                xRng = CType(xlSheet.Cells(temp_line, 40), Microsoft.Office.Interop.Excel.Range)
                val = xRng.Value()
                temp_party = val.ToString
                temp_party_first = temp_party
                If Trim(temp_party_All) <> "" Then
                    temp_party_All &= ", "
                End If
                temp_party_All &= temp_party


                Insert_Record = False
                ' if we have changed serial number, or doc date, then 
                'If (Trim(temp_serial) <> Trim(last_serial)) And Trim(last_serial) <> "" Then
                '    Response.Write("<br/>-----------------Serial Changed")
                '    Response.Write("<br/>LAST SERIAL: " & last_serial)
                '    Response.Write("<br/>PARTY: " & temp_party_All)
                '    Insert_Record = True
                'ElseIf (Trim(temp_Date) <> Trim(last_date)) And Trim(last_date) <> "" Then
                '    Response.Write("<br/>-----------------Date Changed")
                '    Response.Write("<br/>LAST SERIAL: " & last_serial)
                '    Response.Write("<br/>PARTY: " & temp_party_All)
                '    Insert_Record = True
                'Else

                'End If

                ' if neither has changed, then dont enter 
                If (Trim(temp_serial) = Trim(last_serial)) And (Trim(temp_Date1) = Trim(last_date)) Then
                    Insert_Record = False
                Else
                    Insert_Record = True
                End If

                If Insert_Record = True Then
                    Call CREATE_INSERT_FUNCTION()
                    temp_party_All = ""
                End If



                last_party = temp_party_first
                last_reg = pub_reg_no

                last_serial = temp_serial
                last_date = temp_Date1
            End If


            '	S	MFR MDL CODE	ENG MFR MDL	YEAR MFR	TYPE REGISTRANT	NAME	STREET	STREET2	CITY	STATE	ZIP CODE	REGION	COUNTY	COUNTRY	LAST ACTION DATE	CERT ISSUE DATE	CERTIFICATION	TYPE AIRCRAFT	TYPE ENGINE	STATUS CODE	MODE S CODE	FRACT OWNER	AIR WORTH DATE	OTHER NAMES(1)	OTHER NAMES(2)	OTHER NAMES(3)	OTHER NAMES(4)	OTHER NAMES(5)	EXPIRATION DATE	UNIQUE ID	KIT MFR	 KIT MODEL	MODE S CODE HEX	CODE	MFR	MODEL	TYPE-COLLATERAL	COLLATERAL	PARTY	DOC-ID	DRDATE	PROCESSING-DATE	CORR-DATE	CORR-ID	SERIAL-ID

        Next

        '' then we had data 
        'If Trim(last_serial) <> "" Then
        '    Call CREATE_INSERT_FUNCTION()
        'End If


    End Sub
    Public Sub run_excel_scraper_reserved()

        Dim temp_directory As String = "C:\Users\Matt Wanner\Desktop\"
        '   Dim temp_directory As String = "D:\jetnetassistant\DOC_REQUEST\"

        '    Dim temp_file_name As String = "doc_index.html"
        Dim temp_file_name As String = "reserved_data.xlsx"
        '  Dim temp_file_name As String = "doc_index.csv"
        Dim temp_line As Long = 0
        Dim temp_String As String = ""

        Dim temp_party As String = ""
        Dim temp_party_All As String = ""
        Dim temp_serial As String = ""
        Dim last_serial As String = ""
        Dim last_date As String = ""


        Dim ac_id As Long = 0
        Dim publist_research_note As String = ""
        Dim Insert_Record As String = ""
        Dim pub_insert As Boolean = False


        Try


            Insert_Query_Start = " INSERT INTO Publication_Listing"
            Insert_Query_Start &= " (publist_ac_id"
            Insert_Query_Start &= " ,publist_journ_id"
            Insert_Query_Start &= " ,publist_source"
            Insert_Query_Start &= " ,publist_reg_no"
            Insert_Query_Start &= " ,publist_ser_no"
            Insert_Query_Start &= " ,publist_description"
            Insert_Query_Start &= " ,publist_price"
            Insert_Query_Start &= " ,publist_aftt"
            Insert_Query_Start &= " ,publist_seller_info"
            Insert_Query_Start &= " ,publist_picture"
            Insert_Query_Start &= " ,publist_status"
            Insert_Query_Start &= " ,publist_url"
            Insert_Query_Start &= "  ,publist_clear_date"
            Insert_Query_Start &= "  ,publist_acct_rep"
            Insert_Query_Start &= "  ,publist_entry_date"
            Insert_Query_Start &= "  ,publist_update_date"
            Insert_Query_Start &= "  ,publist_original_desc"
            Insert_Query_Start &= "  ,publist_latest_change"
            Insert_Query_Start &= "  ,publist_user_id"
            Insert_Query_Start &= "  , publist_type "
            Insert_Query_Start &= "  , publist_comp_id "
            Insert_Query_Start &= "  ,publist_process_status, publist_research_note, publist_category)"
            Insert_Query_Start &= " VALUES( "



            Try

                MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
                '  MySqlConn_JETNET.ConnectionString = Inhouse_Test_Connection
                ' MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN
                MySqlConn_JETNET.Open()
                MySqlCommand_JETNET.Connection = MySqlConn_JETNET
                MySqlCommand_JETNET.CommandType = CommandType.Text
                MySqlCommand_JETNET.CommandTimeout = 60


            Catch ex As Exception

            End Try

            '  Call insert_into_eventlog("Reserved Data Started", "Research Assistant")

            'Dim string_text As String = "" 
            'Using sr As New StreamReader(temp_directory & temp_file_name)
            '    '  Using sr As New StreamReader("C:\Controller\" & i & ".htm")
            '    string_text = sr.ReadToEnd()

            '    string_text = string_text
            'End Using



            xlApp = CType(CreateObject("Excel.Application"),
                      Microsoft.Office.Interop.Excel.Application)
            xlBook = CType(xlApp.Workbooks.Open(temp_directory & temp_file_name), Microsoft.Office.Interop.Excel.Workbook)

            ' xlBook = CType(xlApp.Workbooks.Add,Microsoft.Office.Interop.Excel.Workbook)
            xlSheet = CType(xlBook.Worksheets(1),
                Microsoft.Office.Interop.Excel.Worksheet)



            Dim xRng As Microsoft.Office.Interop.Excel.Range
            Dim val As Object

            ' no line 0 , and line 1 is columns 
            For temp_line = 2 To 10000
                'temp_String = xlSheet.Cells(temp_line, 1).value()
                temp_String = temp_String
                acpub_original_name = ""



                ' COLUMN 1 is N-NUMBER
                xRng = CType(xlSheet.Cells(temp_line, 1), Microsoft.Office.Interop.Excel.Range)
                val = xRng.Value()
                If Not IsNothing(val) Then
                    temp_String = val.ToString
                Else
                    temp_String = ""
                End If
                pub_reg_no = temp_String
                pub_reg_no = Replace(pub_reg_no, """", "")



                xRng = CType(xlSheet.Cells(temp_line, 2), Microsoft.Office.Interop.Excel.Range)
                val = xRng.Value()
                If Not IsNothing(val) Then
                    temp_serial = val.ToString
                Else
                    temp_serial = ""
                End If
                pub_ser_no = temp_serial
                pub_ser_no = Replace(pub_ser_no, """", "")

                temp_party_first = ""
                xRng = CType(xlSheet.Cells(temp_line, 7), Microsoft.Office.Interop.Excel.Range)
                val = xRng.Value()
                If Not IsNothing(val) Then
                    temp_party_first = val.ToString
                Else
                    temp_party_first = ""
                End If

                ' country
                xRng = CType(xlSheet.Cells(temp_line, 15), Microsoft.Office.Interop.Excel.Range)
                If Not IsNothing(xRng.Value()) Then
                    val = xRng.Value()
                    temp_Country = val.ToString
                Else
                    temp_Country = ""
                End If



                ' MFR
                xRng = CType(xlSheet.Cells(temp_line, 36), Microsoft.Office.Interop.Excel.Range)
                val = xRng.Value()
                If Not IsNothing(val) Then
                    temp_make = val.ToString
                Else
                    temp_make = ""
                End If

                ' Model
                xRng = CType(xlSheet.Cells(temp_line, 37), Microsoft.Office.Interop.Excel.Range)
                val = xRng.Value()
                If Not IsNothing(val) Then
                    temp_model = val.ToString
                Else
                    temp_model = ""
                End If


                ' COLUMN AL 
                xRng = CType(xlSheet.Cells(temp_line, 38), Microsoft.Office.Interop.Excel.Range)
                val = xRng.Value()
                If Not IsNothing(val) Then
                    temp_String = val.ToString
                Else
                    temp_String = ""
                End If
                pub_reg_no_doc_pending = temp_String
                pub_reg_no_doc_pending = Replace(pub_reg_no_doc_pending, """", "")

                acpub_original_name = "Reserved Doc Reg: " & pub_reg_no & "/" & pub_reg_no_doc_pending & " Serno: " & pub_ser_no

                ' add N if its not there to all US aircraft 
                If Trim(temp_Country) = "US" Then
                    If Left(Trim(pub_reg_no), 1) <> "N" Then
                        pub_reg_no = Trim("N" & Trim(pub_reg_no))
                    End If
                    If Left(Trim(pub_reg_no_doc_pending), 1) <> "N" Then
                        pub_reg_no_doc_pending = Trim("N" & Trim(pub_reg_no_doc_pending))
                    End If
                End If


                ' if the reg docs r different, then we may need to change 
                pub_insert = False
                If Trim(pub_reg_no) <> Trim(pub_reg_no_doc_pending) Then

                    temp_ac_id = 0
                    Call FIND_AC_ID_RESERVED(pub_reg_no_doc_pending)   ' see if it matches the AL column - DOC GOING TO BE CHANGED TO HAS ALRREADY BEEN DONE BY RESEARCH 

                    ' IF WE HAVENT ALREADY CHANGED IT 
                    If temp_ac_id = 0 Then

                        ' if we havent already done it. check to see if its with the old id 
                        Call FIND_AC_ID_RESERVED(pub_reg_no)

                        If temp_ac_id = 221269 Then
                            temp_ac_id = temp_ac_id
                        End If

                        If find_Pending_Reg(pub_reg_no_doc_pending, temp_ac_id) = True Then ' SEE IF ITS ALREADY PENDING 
                            pub_insert = False ' then  we have already added it to pending 
                        Else
                            pub_insert = True
                        End If
                    Else
                        ' we have matched the AL column already and we dont need to do more 
                        pub_insert = False
                    End If


                    If pub_insert = True Then
                            insert_doc_request_reserved()
                        End If
                    End If



                '' SERIAL NUMBER
                'xRng = CType(xlSheet.Cells(temp_line, 2), Microsoft.Office.Interop.Excel.Range)
                'val = xRng.Value()
                'If Not IsNothing(val) Then
                '    temp_serial = val.ToString
                'Else
                '    temp_serial = ""
                'End If
                'pub_ser_no = temp_serial
                'pub_ser_no = Replace(pub_ser_no, """", "")

                'If Trim(temp_String) = "" And Trim(temp_serial) = "" Then
                '    temp_line = 10001
                'Else

                '    ' MFR
                '    xRng = CType(xlSheet.Cells(temp_line, 36), Microsoft.Office.Interop.Excel.Range)
                '    val = xRng.Value()
                '    temp_make = val.ToString

                '    ' Model
                '    xRng = CType(xlSheet.Cells(temp_line, 37), Microsoft.Office.Interop.Excel.Range)
                '    val = xRng.Value()
                '    temp_model = val.ToString


                '    ' DR DATE 
                '    xRng = CType(xlSheet.Cells(temp_line, 42), Microsoft.Office.Interop.Excel.Range)
                '    val = xRng.Value()
                '    temp_Date1 = val.ToString

                '    ' country
                '    xRng = CType(xlSheet.Cells(temp_line, 15), Microsoft.Office.Interop.Excel.Range)
                '    If Not IsNothing(xRng.Value()) Then
                '        val = xRng.Value()
                '        temp_Country = val.ToString
                '    Else
                '        temp_Country = ""
                '    End If



                '    ' Party
                '    xRng = CType(xlSheet.Cells(temp_line, 40), Microsoft.Office.Interop.Excel.Range)
                '    val = xRng.Value()
                '    temp_party = val.ToString
                '    temp_party_first = temp_party
                '    If Trim(temp_party_All) <> "" Then
                '        temp_party_All &= ", "
                '    End If
                '    temp_party_All &= temp_party


                '    Insert_Record = False
                '    ' if we have changed serial number, or doc date, then 
                '    'If (Trim(temp_serial) <> Trim(last_serial)) And Trim(last_serial) <> "" Then
                '    '    Response.Write("<br/>-----------------Serial Changed")
                '    '    Response.Write("<br/>LAST SERIAL: " & last_serial)
                '    '    Response.Write("<br/>PARTY: " & temp_party_All)
                '    '    Insert_Record = True
                '    'ElseIf (Trim(temp_Date) <> Trim(last_date)) And Trim(last_date) <> "" Then
                '    '    Response.Write("<br/>-----------------Date Changed")
                '    '    Response.Write("<br/>LAST SERIAL: " & last_serial)
                '    '    Response.Write("<br/>PARTY: " & temp_party_All)
                '    '    Insert_Record = True
                '    'Else

                '    'End If

                '    ' if neither has changed, then dont enter 
                '    If (Trim(temp_serial) = Trim(last_serial)) And (Trim(temp_Date1) = Trim(last_date)) Then
                '        Insert_Record = False
                '    Else
                '        Insert_Record = True
                '    End If





                '    last_party = temp_party_first
                '    last_reg = pub_reg_no

                '    last_serial = temp_serial
                '    last_date = temp_Date1
                'End If


                '	S	MFR MDL CODE	ENG MFR MDL	YEAR MFR	TYPE REGISTRANT	NAME	STREET	STREET2	CITY	STATE	ZIP CODE	REGION	COUNTY	COUNTRY	LAST ACTION DATE	CERT ISSUE DATE	CERTIFICATION	TYPE AIRCRAFT	TYPE ENGINE	STATUS CODE	MODE S CODE	FRACT OWNER	AIR WORTH DATE	OTHER NAMES(1)	OTHER NAMES(2)	OTHER NAMES(3)	OTHER NAMES(4)	OTHER NAMES(5)	EXPIRATION DATE	UNIQUE ID	KIT MFR	 KIT MODEL	MODE S CODE HEX	CODE	MFR	MODEL	TYPE-COLLATERAL	COLLATERAL	PARTY	DOC-ID	DRDATE	PROCESSING-DATE	CORR-DATE	CORR-ID	SERIAL-ID

            Next

            '' then we had data 
            'If Trim(last_serial) <> "" Then
            '    Call CREATE_INSERT_FUNCTION()
            'End If


            temp_String = temp_String

        Catch ex As Exception
        Finally
            MySqlConn_JETNET.Dispose()
        MySqlConn_JETNET.Close()
        MySqlConn_JETNET = Nothing
        End Try


    End Sub
    Public Sub CREATE_INSERT_FUNCTION()
        ' pub_ser_no = last_serial 
        acpub_status = "O"
        pub_desc = ""
        acpub_original_name = "Doc Index: Date: " & temp_Date1  ' title 


        If Trim(temp_Country) = "US" Then
            If Left(Trim(pub_reg_no), 1) <> "N" Then
                pub_reg_no = Trim("N" & Trim(pub_reg_no))
            End If
        End If



        temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
        If temp_ac_id = 0 Then
            temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
            If temp_ac_id = 0 Then
                temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                If temp_ac_id = 0 Then
                    temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                    If temp_ac_id = 0 Then
                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                        If temp_ac_id = 0 Then
                            If Trim(pub_reg_no) <> "" Then
                                temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                                If temp_ac_id = 0 Then
                                    ' one last try 
                                    temp_ac_id = find_ac_global_search("", temp_make, temp_model, pub_reg_no)
                                    If temp_ac_id = 0 Then
                                        temp_ac_id = temp_ac_id
                                    End If
                                End If
                            End If


                            If temp_ac_id = 0 Then
                                If Trim(pub_ser_no) <> "" Then
                                    temp_ac_id = find_ac_ac_search(pub_ser_no, temp_make, temp_model, "")
                                End If

                                If temp_ac_id = 0 Then
                                    If Trim(pub_reg_no) <> "" Then
                                        temp_ac_id = find_ac_ac_search("", temp_make, temp_model, pub_reg_no)
                                    End If

                                    If temp_ac_id = 0 Then
                                        If Trim(pub_ser_no) <> "" And Left(Trim(pub_ser_no), 1) = "0" Then
                                            temp_ac_id = find_ac_ac_search(Right(Trim(pub_ser_no), Len(Trim(pub_ser_no)) - 1), temp_make, temp_model, "")
                                        End If
                                    End If

                                End If
                            End If
                            temp_ac_id = temp_ac_id



                        End If

                    End If
                End If
            End If

        End If

        If temp_ac_id = 0 Then
            temp_ac_id = temp_ac_id
        End If






        If CHECK_IF_DOC_REQUEST_EXISTS(acpub_original_name, temp_ac_id, temp_party_first) = False Then



            Insert_Query1 = Insert_Query_Start
            Insert_Query1 &= " " & temp_ac_id & ""
            Insert_Query1 &= ",0"
            Insert_Query1 &= ", '0'"
            Insert_Query1 &= ", '" & pub_reg_no & "'"
            Insert_Query1 &= ", '" & pub_ser_no & "'"
            Insert_Query1 &= ", '" & Left(Replace(pub_desc, "'", ""), 799) & "'"
            '  Insert_Query1 &= ", '" & Left(pub_desc, 119) & "'"

            Insert_Query1 &= ", '" & pub_price & "'"
            Insert_Query1 &= ", '" & pub_aftt & "'"
            Insert_Query1 &= ", '" & Replace(pub_seller_info, "'", "") & "'"
            Insert_Query1 &= ", '" & pub_picture & "'"
            Insert_Query1 &= ", '" & acpub_status & "'"
            Insert_Query1 &= ", '" & pub_url & "'"
            Insert_Query1 &= ", ''"  'clear date
            Insert_Query1 &= ", 'TN03'"  'acct rep
            Insert_Query1 &= ", '" & Date.Now & "'"  'entry date
            Insert_Query1 &= ", ''"  'update date
            Insert_Query1 &= ", '" & Replace(Trim(acpub_original_name), "'", "") & "'"  'original desc
            Insert_Query1 &= ", ''"  'latest change
            Insert_Query1 &= ", 'mvit'"  'user id  
            Insert_Query1 &= ", 'Aircraft'"  'type 
            Insert_Query1 &= ", '" & pub_comp_id & "'"  'comp id

            Insert_Query1 &= ", ''"

            Insert_Query1 &= ", '" & Replace(temp_party_first, "'", "") & "','Doc Request'"

            Insert_Query1 &= ")"

            MySqlCommand_JETNET.CommandText = Insert_Query1
            MySqlCommand_JETNET.ExecuteNonQuery()
            MySqlCommand_JETNET.Dispose()

        End If

    End Sub

    Public Sub FIND_AC_ID_RESERVED(ByVal REAL_TEMP_REG As String)
        ' pub_ser_no = last_serial 
        acpub_status = "O"
        pub_desc = ""
        acpub_original_name = "Doc Index: Date: " & temp_Date1  ' title 




        temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
        If temp_ac_id = 0 Then
            temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
            If temp_ac_id = 0 Then
                temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                If temp_ac_id = 0 Then
                    temp_ac_id = find_ac_global_search(pub_ser_no, "", "", REAL_TEMP_REG)
                    If temp_ac_id = 0 Then
                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                        If temp_ac_id = 0 Then
                            If Trim(REAL_TEMP_REG) <> "" Then
                                temp_ac_id = find_ac_global_search("", "", "", REAL_TEMP_REG)
                                If temp_ac_id = 0 Then
                                    ' one last try 
                                    temp_ac_id = find_ac_global_search("", temp_make, temp_model, REAL_TEMP_REG)
                                    If temp_ac_id = 0 Then
                                        temp_ac_id = temp_ac_id
                                    End If
                                End If
                            End If


                            If temp_ac_id = 0 Then
                                If Trim(pub_ser_no) <> "" Then
                                    temp_ac_id = find_ac_ac_search(pub_ser_no, temp_make, temp_model, "")
                                End If

                                If temp_ac_id = 0 Then
                                    If Trim(REAL_TEMP_REG) <> "" Then
                                        temp_ac_id = find_ac_ac_search("", temp_make, temp_model, REAL_TEMP_REG)
                                    End If

                                    If temp_ac_id = 0 Then
                                        If Trim(pub_ser_no) <> "" And Left(Trim(pub_ser_no), 1) = "0" Then
                                            temp_ac_id = find_ac_ac_search(Right(Trim(pub_ser_no), Len(Trim(pub_ser_no)) - 1), temp_make, temp_model, "")
                                        End If
                                    End If

                                End If
                            End If
                            temp_ac_id = temp_ac_id



                        End If

                    End If
                End If
            End If

        End If

        If temp_ac_id = 0 Then
            temp_ac_id = temp_ac_id
        End If


    End Sub

    Public Sub insert_doc_request_reserved()



        If CHECK_IF_DOC_REQUEST_EXISTS(acpub_original_name, temp_ac_id, temp_party_first) = False Then



            Insert_Query1 = Insert_Query_Start
            Insert_Query1 &= " " & temp_ac_id & ""
            Insert_Query1 &= ",0"
            Insert_Query1 &= ", '0'"
            Insert_Query1 &= ", '" & pub_reg_no & "'"
            Insert_Query1 &= ", '" & pub_ser_no & "'"
            Insert_Query1 &= ", '" & Left(Replace(pub_desc, "'", ""), 799) & "'"
            '  Insert_Query1 &= ", '" & Left(pub_desc, 119) & "'"

            Insert_Query1 &= ", '" & pub_price & "'"
            Insert_Query1 &= ", '" & pub_aftt & "'"
            Insert_Query1 &= ", '" & Replace(pub_seller_info, "'", "") & "'"
            Insert_Query1 &= ", '" & pub_picture & "'"
            Insert_Query1 &= ", '" & acpub_status & "'"
            Insert_Query1 &= ", '" & pub_url & "'"
            Insert_Query1 &= ", ''"  'clear date
            Insert_Query1 &= ", 'TN03'"  'acct rep
            Insert_Query1 &= ", '" & Date.Now & "'"  'entry date
            Insert_Query1 &= ", ''"  'update date
            Insert_Query1 &= ", '" & Trim(acpub_original_name) & "'"  'original desc
            Insert_Query1 &= ", ''"  'latest change
            Insert_Query1 &= ", 'mvit'"  'user id  
            Insert_Query1 &= ", 'Aircraft'"  'type 
            Insert_Query1 &= ", '" & pub_comp_id & "'"  'comp id

            Insert_Query1 &= ", ''"

            Insert_Query1 &= ", '" & temp_party_first & "','Doc Request'"

            Insert_Query1 &= ")"


            Insert_Query1 = Insert_Query1

            '  Response.Write("<Br>" & Insert_Query1)

            Response.Write("<Br><br/>ACID: " & temp_ac_id & " SERNO " & pub_ser_no & " REGNO " & pub_reg_no & " PENDING/NEW REGNO: " & pub_reg_no_doc_pending)

            '  MySqlCommand_JETNET.CommandText = Insert_Query1
            '   MySqlCommand_JETNET.ExecuteNonQuery()
            '  MySqlCommand_JETNET.Dispose()

        End If



    End Sub

    Public Function CHECK_IF_DOC_REQUEST_EXISTS(ByVal original_Desc As String, ByVal ac_id As Long, ByVal temp_party_first As String) As Boolean
        CHECK_IF_DOC_REQUEST_EXISTS = False
        Dim atemptable As New DataTable

        Dim SqlConn As New SqlClient.SqlConnection
        Dim SqlCommand As New SqlClient.SqlCommand
        Dim SqlReader As SqlClient.SqlDataReader
        Dim Query As String = ""
        Dim original_found As Integer = 0


        Try

            SqlConn.ConnectionString = MySqlConn_JETNET.ConnectionString
            SqlConn.Open()
            SqlCommand.Connection = SqlConn
            SqlCommand.CommandType = CommandType.Text
            SqlCommand.CommandTimeout = 60


            Query = "select * from Publication_Listing with (NOLOCK) where publist_ac_id = '" & ac_id & "'   "
            Query &= " and publist_original_desc = '" & original_Desc & "' "

            Query &= " and publist_research_note like '" & temp_party_first & "%' "


            SqlCommand.CommandText = Query.ToString
            SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

            Try
                atemptable.Load(SqlReader)
            Catch constrExc As System.Data.ConstraintException
                Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
            End Try

            If atemptable.Rows.Count > 0 Then
                For Each r As DataRow In atemptable.Rows
                    ' pub_yacht_id = r.Item("publist_ac_id")
                    CHECK_IF_DOC_REQUEST_EXISTS = True
                Next
            End If

        Catch ex As Exception
            Return Nothing
            '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
        Finally
            SqlReader = Nothing

            SqlConn.Dispose()
            SqlConn.Close()
            SqlConn = Nothing

            SqlCommand.Dispose()
            SqlCommand = Nothing
        End Try


    End Function

    Public Sub RUN_AC_Exchange()


    '  Try

    'MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
    ''MySqlConn_JETNET.ConnectionString = Inhouse_Test_Connection
    '' MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN
    'MySqlConn_JETNET.Open()
    'MySqlCommand_JETNET.Connection = MySqlConn_JETNET
    'MySqlCommand_JETNET.CommandType = CommandType.Text
    'MySqlCommand_JETNET.CommandTimeout = 60

    '' Call insert_into_eventlog("Aircraft Pubs Started", "Research Assistant")

    'ypl_start_date = Date.Now

    'Call Find_Naughty_Models()


    'dont care --------------------------------------------------
    'Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/36/cirrus")


    'dont care --------------------------------------------------





    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/35/astragulfstream")



    'Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/34/aerostar")
    'Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/27/agusta")
    'Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/19/airbus")
    'Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/62/airbus-helicopter")
    'Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/59/american-champion") 
    'Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/56/bae")

    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/49/columbia")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/14/daher-socata")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/38/diamond")

    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/21/dornier")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/13/eclipse")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/3/embraer")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/4/hawker")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/22/honda")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/60/leonardo")

    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/50/nextant") 
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/10/pilatus")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/18/piper") 
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/40/quest")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/61/robinson") 
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/32/sikorsky")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/45/twin-commander")


    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/11/dassault")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/1/cessna")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/6/bombardier")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/20/boeing")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/28/bell")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/12/beechcraft")
    Call Scrape_Aircraft_Exchange_Detail("https://aircraftexchange.com/aircraft-for-sale/2/gulfstream")




    'Catch ex As Exception
    'Finally
    '  MySqlConn_JETNET.Dispose()
    '  MySqlConn_JETNET.Close()
    '  MySqlConn_JETNET = Nothing
    'End Try



  End Sub
  '  Public Sub asset_insight_functions()

  '    Dim client As New HttpClient

  '    static HttpClient   = new HttpClient();
  'client.BaseAddress = new Uri("https://api.assetinsight.com/api/v1/process/analysis?_format=json");

  '// Add Headers
  'client.DefaultRequestHeaders.Accept.Clear();
  'client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", "encrypted keys public:private");
  'client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

  'var postVariables = new {
  '  user = 9999, // User ID to tag to from previous steps
  '  asset = new {
  '    type         = 0,
  '    manufacturer = "bombardier",
  '    model        = 534,
  '    version      = 1,
  '    manufacture  = 982972800, // All Dates as Timestamps
  '    serial       = "test",
  '    tail         = "example",
  '    coverage     = new {
  '      airframe   = "program_name",
  '      engines    = "program_name",
  '      apu        = "program_name",
  '    avionics = "program_name"
  '    },
  '    inspections  = new {
  '      inspection_id = new {
  '        hours    = 4300,
  '    cycles = 2200
  '      },
  '      inspection_id = new {
  '        hours    = 4300,
  '    cycles = 2200
  '      }
  '    }
  '  }
  '};

  'try {
  '  HttpResponseMessage response = await client.PostAsJsonAsync(postVariables);

  '  if (response.IsSuccessStatusCode) {
  '    // Parse the Reponse Object
  '    var dataObjects = response.Content.ReadAsAsync<IEnumerable<DataObject>>().Result;

  '    foreach (var data in dataObjects) {
  '      // INSERT CODE HERE
  '    }
  '  }
  '  else {
  '    Console.WriteLine("{0} ({1})", (int)response.StatusCode, response.ReasonPhrase);
  '  }
  '}
  'catch (Exception e) {
  '  Console.WriteLine(e.Message);
  '}

  '  End Sub
  Public Sub JETNET_NEWS()

    Dim sQuery As String = ""
    Dim counter1 As Integer = 0
    Dim link_array(100) As String
    Dim link_id_array(100) As String
    Dim temp_string As String = ""
    Dim mail_string As String = ""
    Dim function_count As Long = 0
    Dim total_count As Long = 0
    Try
 



      If InStr(Server.MapPath(""), "jetnetabi", CompareMethod.Text) > 0 Then
        MySqlConn_JETNET.ConnectionString = "DSN=JETNET_Conn"
        MySqlConn_JETNET2.ConnectionString = "DSN=JETNET_Conn"
      Else
        '   MySqlConn_JETNET.ConnectionString = "Data Source=www.jetnetsql1.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
        '  MySqlConn_JETNET2.ConnectionString = "Data Source=www.jetnetsql1.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
        MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN
        MySqlConn_JETNET2.ConnectionString = JETNET_LIVE_SQL_CONN

      End If
      '  Response.Write(Server.MapPath("") & "<br>")
      ' Response.Write(InStr(Server.MapPath(""), "jetnetabi", CompareMethod.Text) & "<br><br>")

      ' Response.Write(MySqlConn_JETNET2.ConnectionString & "<Br>")
      ' Response.End()

      sQuery = "select abinewssrc_id, abinewssrc_feed_link "

      sQuery = sQuery & " from abi_news_source "
      sQuery = sQuery & " where abinewssrc_feed_flag='Y' "


      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120

      MySqlConn_JETNET2.Open()
      MySqlCommand_JETNET2.Connection = MySqlConn_JETNET2
      MySqlCommand_JETNET2.CommandType = CommandType.Text
      MySqlCommand_JETNET2.CommandTimeout = 120

      MySqlCommand_JETNET.CommandText = sQuery
      MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()
      If MyAircraftReader_JETNET.HasRows Then

        Do While MyAircraftReader_JETNET.Read()

          '  Response.Write(MyAircraftReader_JETNET("abinewssrc_id") & "-" & MyAircraftReader_JETNET("abinewssrc_feed_link") & "<br>")


          link_array(counter1) = MyAircraftReader_JETNET("abinewssrc_feed_link")
          link_id_array(counter1) = MyAircraftReader_JETNET("abinewssrc_id")

          counter1 = counter1 + 1
        Loop
      End If

      MyAircraftReader_JETNET.Close()


            '     Call insert_into_eventlog("Aircraft News Started", "Research Assistant")


            mail_string += "JETNET ABI NEWS SCRAPER RESULTS - " & Date.Now & "<Br><br>"
            For i = 1 To counter1

                Response.Write("<br/>Running..." & Trim(link_array(i)))
                System.Threading.Thread.Sleep(10)
                Response.Flush()
                Response.Flush()
                System.Threading.Thread.Sleep(10)


                If Trim(link_array(i)) = "www.verticalmag.com/news/press_release_archive" Then
                    Try
                        function_count = ABI_News_Scraper_Vertical("https://www.verticalmag.com/press-releases/", link_id_array(i))
                    Catch ex As Exception
                        Response.Write(ex)
                    End Try
                ElseIf Trim(link_array(i)) = "www.ainonline.com/aviation-news/business-aviation" Then
                    Try
                        function_count = ABI_News_Scraper_AIN_Online("https://www.ainonline.com/", link_id_array(i))
                    Catch ex As Exception
                        Response.Write(ex)
                    End Try
                End If

                If link_id_array(i) <> "" Then
                    mail_string += "http://" & link_array(i) & " (ID " & link_id_array(i) & ")  - Count: " & function_count & "<br>"
                    total_count = total_count + function_count
                End If
            Next


            'Try
            '    function_count = ABI_News_Scraper_Flight_Global("https://gulfstreamnews.com/en/news/", 31) ' 31 is flight global 
            'Catch ex As Exception
            '    Response.Write(ex)
            '        End Try








            'If link_id_array(i) <> "" Then
            '    mail_string += "http://" & link_array(i) & " (ID " & link_id_array(i) & ")  - Count: " & function_count & "<br>"
            '    total_count = total_count + function_count
            'End If

            'If Trim(link_array(i)) = "www.gulfstream.com/news/" Then
            '  Try

            '    'function_count = ABI_News_Scraper_Gulfstream("http://" & link_array(i), link_id_array(i))
            '    ' BAD CONNECTION

            '  Catch ex As Exception
            '    Response.Write(ex)
            '  End Try
            'ElseIf Trim(link_array(i)) = "newsroom.hawkerbeechcraft.com/news-press/" Then
            '  '  Try
            '  '  function_count = ABI_News_Scraper_Hawker_Beach("http://" & link_array(i), link_id_array(i))
            '  '  Catch ex As Exception
            '  '   Response.Write(ex)
            '  ' End Try
            'ElseIf Trim(link_array(i)) = "www.dassaultfalcon.com/en/MediaCenter/Newsd/Pages/Press-release.aspx" Then
            '  ' WORKING  -- WORKING
            '  Try
            '    function_count = ABI_News_Scraper_Dassult_Falcon("http://" & link_array(i), link_id_array(i))
            '  Catch ex As Exception
            '    Response.Write(ex)
            '  End Try
            'ElseIf Trim(link_array(i)) = "www.cessna.com/news/news-releases.html" Then
            '  ' Try
            '  'function_count = ABI_News_Scraper_Cessna("http://" & link_array(i), link_id_array(i))
            '  '   Catch ex As Exception
            '  'Response.Write(ex)
            '  ' End Try
            'ElseIf Trim(link_array(i)) = "www.flightglobal.com/news-listings/business-aviation-news-listings/" Then
            '  ' WORKING  -- WORKING
            '  Try
            '    function_count = ABI_News_Scraper_Flight_Global("https://" & link_array(i), link_id_array(i))
            '  Catch ex As Exception
            '    Response.Write(ex)
            '  End Try
            'ElseIf Trim(link_array(i)) = "www.ainonline.com/aviation-news/business-aviation" Then
            '  ' '' '' WE ARE GETTING BLOCKED BY THIS SITE 
            '  ' '' ''Try
            '  ' '' ''  function_count = ABI_News_Scraper_AIN_Online("https://" & link_array(i), link_id_array(i))
            '  ' '' ''Catch ex As Exception
            '  ' '' ''  Response.Write(ex)
            '  ' '' ''End Try
            '  ' '' ''ElseIf Trim(link_array(i)) = "www.rotorpad.com/general/" Then
            '  ' '' '' function_count = ABI_News_Scraper_Rotorpad("http://" & link_array(i), link_id_array(i))
            'ElseIf Trim(link_array(i)) = "www.bartintl.com" Then
            '            ' WORKING  -- WORKING
            '            'Try
            '            '  function_count = ABI_News_Scraper_BART("http://" & link_array(i), link_id_array(i))
            '            'Catch ex As Exception
            '            '  Response.Write(ex)
            '            'End Try
            '        ElseIf Trim(link_array(i)) = "www.avweb.com" Then
            '  Try
            '    function_count = ABI_News_Scraper_Avweb("http://" & link_array(i), link_id_array(i))
            '  Catch ex As Exception
            '    Response.Write(ex)
            '  End Try
            'ElseIf Trim(link_array(i)) = "www.verticalmag.com/news/press_release_archive" Then
            '  Try
            '    function_count = ABI_News_Scraper_Vertical("https://www.verticalmag.com/press-releases/", link_id_array(i))
            '  Catch ex As Exception
            '    Response.Write(ex)
            '  End Try


            'ElseIf Trim(link_array(i)) = "feeds.feedburner.com/FlyCorporateNews?format=xml" Then
            '  ' --WORKING - JUST NOT TOO MANY RECORDS
            '  Try
            '    function_count = ABI_News_Scraper_FlyCorperate("http://" & link_array(i), link_id_array(i))
            '  Catch ex As Exception
            '    Response.Write(ex)
            '  End Try

            'ElseIf Trim(link_array(i)) = "www.aviationweek.com/business-aviation" Then
            '    Try
            '      function_count = ABI_News_Scraper_AVIATION_WEEK("http://" & link_array(i), link_id_array(i))
            '    Catch ex As Exception
            '      Response.Write(ex)
            '    End Try

            'ElseIf Trim(link_array(i)) = "http://corpjetfin.live.subhub.com/categories/Corporate-Jet-News" Then
            '    ' Try
            '    'function_count = ABI_News_Scraper_Corp_Jet(link_array(i), link_id_array(i))
            '    '    Catch ex As Exception
            '    'Response.Write(ex)
            '    '   End Try
            'End If




            '   Next

            mail_string += "<br>Total News Articles Loaded: " & total_count

            'Dim ToAddress As String = "info@jetnetabainews.com"

            ''(1) Create the MailMessage instance
            'Dim mm As New MailMessage("jetnetabinews@test.com", ToAddress)

            ''(2) Assign the MailMessage's properties
            'mm.Subject = "JETNET ABI NEWS SCRAPER RESULTS - " & Date.Now
            'mm.Body = mail_string
            'mm.IsBodyHtml = True

            ''(3) Create the SmtpClient object
            'Dim smtp As New SmtpClient
            'smtp.Host = "www.jetnetabainews.com"
            ''(4) Send the MailMessage (will use the Web.config settings)
            'smtp.Send(mm)




            '     Call insert_into_eventlog("Aircraft News Finished", "Research Assistant")

            Response.Write(mail_string)

      'ABI_News_Scraper_Gulfstream_Whole("http://www.gulfstream.com/news/releases/index.cfm", 77, 2)
    Catch ex As Exception
      Response.Write(ex)
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()
      MySqlCommand_JETNET.Dispose()

      MySqlConn_JETNET2.Close()
      MySqlConn_JETNET2.Dispose()
      MySqlCommand_JETNET2.Dispose()


    End Try
  End Sub
  Public Sub get_MMSI_NUMBERS_top_function()

    '  yt_table = "<table cellspacing='0' cellpadding='4' border='1'>"
    '  yt_table &= "<tr><td colspan='5'><b>MarineTraffic Listings</b></td></tr>"
    '  yt_table &= "<tr><td><b>Yacht Name</b></td><td><b>YachtID</b></td><td><b>MMSI (Yacht Spot MMSI)</b></td><td><b>IMO (Yacht Spot IMO)</b></td><td><b>Status</b></td><td><b>Action</b></td></tr>"

    Try


      SqlConn_YPL.ConnectionString = Inhouse_Live_Connection
      SqlConn_YPL.Open()
      SqlCommand_YPL.Connection = SqlConn_YPL
      SqlCommand_YPL.CommandType = CommandType.Text
      SqlCommand_YPL.CommandTimeout = 60

      ypl_start_date = Date.Now

      Call insert_into_eventlog("Marine Traffic Started", "Research Assistant")

      For i = 1 To 50
        Response.Write("<br/>Running Page " & i & " ")
        System.Threading.Thread.Sleep(10)
        Response.Flush()
        Response.Flush()
        System.Threading.Thread.Sleep(10)
        Call scrape_for_mmsi(i)
      Next


      ' yt_table &= "</table>"

      yt_table &= "<table cellspacing='0' cellpadding='4' border='1'>"
      yt_table &= "<tr>"
      yt_table &= "<td align='left'>Yachts Searched: " & total_mmsi & "</td>"
      yt_table &= "<td align='left'>Yachts Found: " & yt_count & "</td>"
      yt_table &= "<td align='left'>MMSI Matches Currently: " & match_mmsi_count & "</td>"
      yt_table &= "<td align='left'>Suggested Updates: " & (total_mmsi - match_mmsi_count - blank_mmsi_on_yacht - new_yacht_mmsi) & "</td>"
      yt_table &= "<td align='left'>Suggested New Yachts: " & new_yacht_mmsi & "</td>"
      yt_table &= "<td align='left'>Suggested Automatic Adds: " & blank_mmsi_on_yacht & "</td>"
      yt_table &= "<td align='left'>Duplicate Name/Source: " & dup_count & "</td>"
      yt_table &= "<td align='left'>Types Not Allowed: " & bad_type & "</td>"
      yt_table &= "<td align='left'>Total Inserts: " & new_ypl_insert & "</td>"
      yt_table &= "</tr>"
      yt_table &= "</table>"

      Response.Write(yt_table)

      ' Call get_MMSI_NUMBERS()



      Call insert_into_eventlog("Marine Traffic Finished", "Research Assistant")

    Catch ex As Exception
    Finally

      SqlConn_YPL.Dispose()
      SqlConn_YPL.Close()
      SqlConn_YPL = Nothing

    End Try


  End Sub

  Public Sub get_ac_pub_top_function()

    '  yt_table = "<table cellspacing='0' cellpadding='4' border='1'>"
    '   yt_table &= "<tr><td colspan='5'><b>AC Pub Listings</b></td></tr>"
    '  yt_table &= "<tr><td><b>Aircraft</b></td><td><b>AC ID</b></td><td><b>Owner</b></td><td><b>Status</b></td><td><b>Action</b></td></tr>"
    Dim skip_this As String = "Y"

    Try


      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
      'MySqlConn_JETNET.ConnectionString = Inhouse_Test_Connection
      ' MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN
      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 60

      Call insert_into_eventlog("Aircraft Pubs Started", "Research Assistant")

      ypl_start_date = Date.Now

            Call Find_Naughty_Models()


            Call insert_into_eventlog("Bad Models Found", "Research Assistant")


            'Try

            '    Call Scrape_For_flightmarket("http://flightmarket.com/pt/aeronaves")

            'Catch ex As Exception

            'End Try


            'Call insert_into_eventlog("Flight Market Finished", "Research Assistant")

            Try



                Call Scrape_Barnstormers("https://www.barnstormers.com/listing.php?catid=16738")  ' beechcraft 

                Call Scrape_Barnstormers("https://www.barnstormers.com/listing.php?catid=16754")
                Call Scrape_Barnstormers("https://www.barnstormers.com/listing.php?catid=16758")
                Call Scrape_Barnstormers("https://www.barnstormers.com/listing.php?catid=16759")
                Call Scrape_Barnstormers("https://www.barnstormers.com/listing.php?catid=16776")
                Call Scrape_Barnstormers("https://www.barnstormers.com/listing.php?catid=16777")
                Call Scrape_Barnstormers("https://www.barnstormers.com/listing.php?catid=16807")



                Call Scrape_Barnstormers("https://www.barnstormers.com/listing.php")


            Catch ex As Exception

            End Try


            Call insert_into_eventlog("barnstormers Finished", "Research Assistant")


            Try

                Response.Write("<br/>Running Trade A Plane...")
                System.Threading.Thread.Sleep(10)
                Response.Flush()
                Response.Flush()
                System.Threading.Thread.Sleep(10)
                Call scrape_for_TradeAPlane(0)
                Call scrape_for_TradeAPlane(2)

                'skip_this = "Y"
                'If skip_this = "Y" Then
                'Else
            Catch ex As Exception

            End Try



            Call insert_into_eventlog("TradeAPlane Finished", "Research Assistant")


            Try



                Response.Write("<br/>Running Global AIR...")
                System.Threading.Thread.Sleep(10)
                Response.Flush()
                Response.Flush()
                System.Threading.Thread.Sleep(10)
                Call scrape_for_GlobalAIR("https://www.globalair.com/aircraft-for-sale/Private-Jet")   ' removed last / . msw- 2/10/23 
                Response.Flush()
                Response.Flush()
                System.Threading.Thread.Sleep(10)
                Call scrape_for_GlobalAIR("https://www.globalair.com/aircraft-for-sale/Twin-Engine-Turbine")
                Response.Flush()
                Response.Flush()
                System.Threading.Thread.Sleep(10)
                Call scrape_for_GlobalAIR("https://www.globalair.com/aircraft-for-sale/Gulfstream-G550")
                Response.Flush()
                Response.Flush()
                System.Threading.Thread.Sleep(10)
                Call scrape_for_GlobalAIR("https://www.globalair.com/aircraft-for-sale/Helicopters")
                Response.Flush()
                Response.Flush()
                System.Threading.Thread.Sleep(10)
                Call scrape_for_GlobalAIR("https://www.globalair.com/aircraft-for-sale/Single-Engine-Turbine")
                Response.Flush()
                Response.Flush()
                System.Threading.Thread.Sleep(10)
                Call scrape_for_GlobalAIR("https://www.globalair.com/aircraft-for-sale/Single-Engine-Piston")
                Response.Flush()
                Response.Flush()
                System.Threading.Thread.Sleep(10)
                Call scrape_for_GlobalAIR("https://www.globalair.com/aircraft-for-sale/Twin-Engine-Turbine")

            Catch ex As Exception

            End Try


            Call insert_into_eventlog("GlobalAIR Finished", "Research Assistant")


            Try


                Response.Write("<br/>Running ASO...")
                System.Threading.Thread.Sleep(10)
                Response.Flush()
                Response.Flush()
                System.Threading.Thread.Sleep(10)
                Call scrape_for_ASO()
            Catch ex As Exception

            End Try

            Call insert_into_eventlog("ASO Finished", "Research Assistant")


            Try


                Response.Write("<br/>Running AvData...")

                '  For i = 1 To 13
                'Response.Write("<br/>Running Page " & i & " ")
                System.Threading.Thread.Sleep(10)
                    Response.Flush()
                    Response.Flush()
                    System.Threading.Thread.Sleep(10)
                Call scrape_for_AvBuyer(0)
                '   Next
            Catch ex As Exception

            End Try


            Call insert_into_eventlog("AvBuyer Finished", "Research Assistant")


            Try
                Call RUN_AC_Exchange()
            Catch ex As Exception

            End Try



            Call insert_into_eventlog("AC_Exchange Finished", "Research Assistant")



            ' 58 total pages saved - 13 minutes 


            '' CONTR
            ''Response.Write("<br/>Running Controller...")


            ' ''page_break_number = 20 
            ''For i = 1 To 15
            ''  Response.Write("<br/>Running Page " & i & " ")
            ''  System.Threading.Thread.Sleep(10)
            ''  Response.Flush()
            ''  Response.Flush()
            ''  System.Threading.Thread.Sleep(10)
            ''  Response.Write("Jets, ")
            ''  System.Threading.Thread.Sleep(10)
            ''  Response.Flush()
            ''  Response.Flush()
            ''  System.Threading.Thread.Sleep(10)
            ''  Call scrape_for_controller(i, "3/jet-aircraft")
            ''  Response.Write("Turboprop, ")
            ''  System.Threading.Thread.Sleep(10)
            ''  Response.Flush()
            ''  Response.Flush()
            ''  System.Threading.Thread.Sleep(10)
            ''  Call scrape_for_controller(i, "8/turboprop-aircraft")
            ''  Response.Write("Piston Twin, ")
            ''  System.Threading.Thread.Sleep(10)
            ''  Response.Flush()
            ''  Response.Flush()
            ''  System.Threading.Thread.Sleep(10)
            ''  Call scrape_for_controller(i, "9/piston-twin-aircraft")
            ''  Response.Write("Piston Heli, ")
            ''  System.Threading.Thread.Sleep(10)
            ''  Response.Flush()
            ''  Response.Flush()
            ''  System.Threading.Thread.Sleep(10)
            ''  Call scrape_for_controller(i, "5/piston-helicopters")
            ''  Response.Write("Turbine Heli, ")
            ''  System.Threading.Thread.Sleep(10)
            ''  Response.Flush()
            ''  Response.Flush()
            ''  System.Threading.Thread.Sleep(10)
            ''  Call scrape_for_controller(i, "7/turbine-helicopters")
            ''Next


            ''System.Threading.Thread.Sleep(10)
            ''Response.Flush()
            ''Response.Flush()
            ''System.Threading.Thread.Sleep(10)

            ''For i = 53 To 59
            ''  Response.Write("<br/>Running General Page " & i)
            ''  System.Threading.Thread.Sleep(10)
            ''  Response.Flush()
            ''  Response.Flush()
            ''  System.Threading.Thread.Sleep(10)
            ''  Call scrape_for_controller(i, "last7")
            ''Next


            Response.Flush()
      'End If


      'If total_pages > 1 Then
      '  For i = 1 To total_pages
      '    Call scrape_for_controller(i, "3/jet-aircraft")

      '    'if we have found the general section start, then we already addd 5 to it (10+5) = 15 , so if i > 15, then stop it 
      '    If acpub_controller_general_start > 0 And i > acpub_controller_general_start Then
      '      i = total_pages
      '    End If
      '  Next
      'End If

      'If total_pages > 1 Then
      '  For i = 1 To total_pages
      '    Call scrape_for_controller(i, "8/turboprop-aircraft")

      '    'if we have found the general section start, then we already addd 5 to it (10+5) = 15 , so if i > 15, then stop it 
      '    If acpub_controller_general_start > 0 And i > acpub_controller_general_start Then
      '      i = total_pages
      '    End If
      '  Next
      'End If



      '  For i = 1 To 1
      'Call scrape_for_controller(i, "3/jet-aircraft")
      ' Call scrape_for_controller(i, "8/turboprop-aircraft")
      ' Call scrape_for_controller(i, "9/piston-twin-aircraft")
      ' Call scrape_for_controller(i, "5/piston-helicopters")
      ' Call scrape_for_controller(i, "7/turbine-helicopters")
      ' Next


      '   yt_table &= "</table>"

      yt_table &= "<table cellspacing='0' cellpadding='4' border='1'>"
      yt_table &= "<tr>"
      yt_table &= "<td align='left'>Pubs Searched: " & acpub_count & "</td>"
      yt_table &= "<td align='left'>Pubs Matched: " & acpub_match_count & "</td>"
      yt_table &= "<td align='left'>Pubs Inserted: " & acpub_insert_count & "</td>"
      yt_table &= "</tr>"
      yt_table &= "</table>"

      Response.Write(yt_table)

      ' Call get_MMSI_NUMBERS()

      If MySqlConn_JETNET.State = ConnectionState.Open Then
        MySqlConn_JETNET.Close()
      End If

      Call Insert_EMail_Queue_Record(yt_table) 

      Call insert_into_eventlog("Aircraft Pubs Finished", "Research Assistant")

      Response.Write("Done Pubs")

    Catch ex As Exception
    Finally

      MySqlConn_JETNET.Dispose()
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET = Nothing

    End Try


  End Sub
  Public Sub RUN_DATA_INTEGRITY_CHECKS()
    Dim results As String = ""
    Dim Temp_Name As String = ""
    Dim Query As String = ""
    Dim atemptable As New DataTable
    Dim atemptable2 As New DataTable
    Dim Temp_Query As String = ""
    Dim Temp_Count As Integer = 0
    Dim dalTable As New DataTable
    Dim dataSet As DataSet = New DataSet("dataSet")
    Dim temp_date As String = ""


    Try

      temp_date = Date.Now

      ' MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
      MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120
      ' REALLY SQL SERVER CONNECTIONS

      ypl_start_date = Date.Now

      Response.Write("<br/>Data Integrity Checks...")
      text_label.Text = ""


      Query = "  SELECT sqlrep_title, sqlrep_query FROM SQL_Report WITH(NOLOCK)"
      Query &= " WHERE sqlrep_level = 'JETNET' AND sqlrep_sub_id = 0 "
      ' Query &= " and sqlrep_title like '%Business Aircraft and Helicopter%' "

      Query &= " and sqlrep_type = 'Data Integrity' "

      MySqlCommand_JETNET.CommandText = Query.ToString
      MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

      Try
        atemptable.Load(MyAircraftReader_JETNET)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

      MyAircraftReader_JETNET.Close()

      If atemptable.Rows.Count > 0 Then
        For Each r As DataRow In atemptable.Rows

          Temp_Name = r.Item("sqlrep_title")
          Temp_Query = r.Item("sqlrep_query")
          Temp_Count = 0



          If Trim(Temp_Name) = "Sale Prices - Not Processed Out of Model Ranges" Then
            Temp_Query = Replace(Temp_Query, "order by amod_make_name, amod_model_name", " and (select COUNT(*) from Aircraft_Value av2 with (NOLOCK) where acval_type = 'CLEAR' and av2.acval_ac_id = Aircraft_Value.acval_ac_id and av2.acval_journ_id = Aircraft_Value.acval_journ_id) = 0  order by amod_make_name, amod_model_name")
          End If

          Response.Write("<br/>Running Select for " & Temp_Name & "")
          System.Threading.Thread.Sleep(10)
          Response.Flush()
          Response.Flush()
          System.Threading.Thread.Sleep(10)
          Response.Write(results)

          If InStr(Temp_Query, "OPENDATASOURCE") > 0 Then
          Else

            If InStr(Temp_Name, "Business Aircraft and Helicopter") > 0 Then
              Temp_Query = Temp_Query
            End If

            Try
              MySqlCommand_JETNET.CommandText = Temp_Query
              MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

              atemptable2 = New DataTable
              atemptable2.Clear()
              atemptable.PrimaryKey = Nothing
              atemptable2.Constraints.Clear()

              ' dalTable = atemptable2.Clone

              'For i = 0 To atemptable2.Columns.Count - 1
              '  If atemptable2.Columns(i).DataType.ToString.ToLower = "system.string" Then
              '    atemptable2.Columns(i).MaxLength = 1500
              '  End If
              'Next

              Try
                atemptable2.Load(MyAircraftReader_JETNET)
              Catch constrExc As System.Data.ConstraintException
                Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
              End Try
              MyAircraftReader_JETNET.Close()
              Temp_Count = atemptable2.Rows.Count
              atemptable2.Dispose()
              atemptable2 = Nothing


            Catch ex As Exception

            End Try
          End If

          Call insert_into_eventlog(Temp_Name & " (" & Temp_Count & ")", "Data Integrity", temp_date)
        Next
      End If





    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()
    End Try
  End Sub
  Public Sub YACHT_PUB(ByVal page As Integer)
    Dim results As String = ""
    Try




      ypl_start_date = Date.Now

      Response.Write("<br/>Running Yacht PUB...")

      If Trim(page) = 1 Then
        For i = 1 To 20
          ' results = get_yachts_for_sale_SUPERYACHTS(i)
          'results = get_yachts_for_sale_YATCO("http://www.yatco.com/search?lengthUnit=2&loa1=24&loa2=&year1=1988&year2=&currency=USD&price1=&price2=&builder=&typeID=0&countryID=0&stateID=0&vesselName=&pg=" & i & "&NR=3")
          results = get_yachts_for_sale_YATCO("http://www.yatco.com/search?pg=" & i & "&lengthUnit=1&loa1=75&loa2=&year1=&year2=&currency=USD&price1=&price2=&builder=&typeID=0&countryID=0&stateID=0&vesselName=&NR=3")
          text_label.Text = ""
          Response.Write("<br/>Running Page " & i)
          System.Threading.Thread.Sleep(10)
          Response.Flush()
          Response.Flush()
          System.Threading.Thread.Sleep(10)
        Next
      Else
        For i = 21 To 40
          ' results = get_yachts_for_sale_SUPERYACHTS(i)
          ' results = get_yachts_for_sale_YATCO("http://www.yatco.com/search?lengthUnit=2&loa1=24&loa2=&year1=1988&year2=&currency=USD&price1=&price2=&builder=&typeID=0&countryID=0&stateID=0&vesselName=&pg=" & i & "&NR=3")
          results = get_yachts_for_sale_YATCO("http://www.yatco.com/search?pg=" & i & "&lengthUnit=1&loa1=75&loa2=&year1=&year2=&currency=USD&price1=&price2=&builder=&typeID=0&countryID=0&stateID=0&vesselName=&NR=3")
          text_label.Text = ""
          Response.Write("<br/>Running Page " & i)
          System.Threading.Thread.Sleep(10)
          Response.Flush()
          Response.Flush()
          System.Threading.Thread.Sleep(10)
        Next
      End If




      results &= Chr(13) & Chr(10) & "YatCo Page - Yachts For Sale:" & Chr(13) & Chr(10)
      results &= "<br/>TOTAL ITEMS/MATCHES: " & CStr(found_yt) & " / " & CStr(Match) & Chr(13) & Chr(10)
      results &= "<br/>TOTAL NOT MATCHES: " & CStr(non_match) & "" & Chr(13) & Chr(10)
      results &= "<br/>------YACHT NOT FOUND: " & CStr(not_found) & "" & Chr(13) & Chr(10)
      results &= "<br/>------MORE THAN 1 YACHT FOUND: " & CStr(more1_found) & "" & Chr(13) & Chr(10)
      results &= "<br/>------YACHT NOT FOR SALE: " & CStr(not_fs) & "" & Chr(13) & Chr(10)
      results &= "<br/>------ASKING PRICE DIFFERENT: " & CStr(wrong_ask) & "" & Chr(13) & Chr(10)

      Response.Write("<br>Results: <br>")
      Response.Write(results)

      Call insert_into_eventlog("Yacht Pubs Finished", "Research Assistant")

    Catch ex As Exception

    End Try
  End Sub
  Public Sub YACHT_NEWS()
    Dim results As String = ""



    ypl_start_date = Date.Now

    text_label.Text = ""
    Response.Write("<br/>Running Yacht News...")
    System.Threading.Thread.Sleep(10)
    Response.Flush()
    Response.Flush()
    System.Threading.Thread.Sleep(10)

    yacht_news_name = ""
    results = results & get_yacht_news_super_yachts()


    Response.Write("<br/>" & results)
    System.Threading.Thread.Sleep(10)
    Response.Flush()
    Response.Flush()
    System.Threading.Thread.Sleep(10)
    results = ""

    TOTAL_NEWS = TOTAL_NEWS

    yacht_news_name = ""
    Try
      results = results & get_super_yacht_news2("http://www.superyachtnews.com/business")
      results = results & get_super_yacht_news2("http://www.superyachtnews.com/technology")
      results = results & get_super_yacht_news2("http://www.superyachtnews.com/owner")
      results = results & get_super_yacht_news2("http://www.superyachtnews.com/design")
      results = results & get_super_yacht_news2("http://www.superyachtnews.com/crew")

    Catch ex As Exception

    End Try 

    Response.Write("<br/>" & results)
    System.Threading.Thread.Sleep(10)
    Response.Flush()
    Response.Flush()
    System.Threading.Thread.Sleep(10)
    results = ""


    TOTAL_NEWS = TOTAL_NEWS

    yacht_news_name = ""
    Try
      results = results & Scrape_This_Page_Super_Yacht_Business()
    Catch ex As Exception

    End Try
    TOTAL_NEWS = TOTAL_NEWS


    Response.Write("<br/>" & results)
    System.Threading.Thread.Sleep(10)
    Response.Flush()
    Response.Flush()
    System.Threading.Thread.Sleep(10)
    results = ""


    yacht_news_name = ""
    results = results & get_yacht_news_super_yacht_times_NEW()
    TOTAL_NEWS = TOTAL_NEWS
    yacht_news_name = ""
    Try
      results = results & get_boat_international_news("https://www.boatinternational.com/yachts/news")
      results = results & get_boat_international_news("https://www.boatinternational.com/yacht-market-intelligence/brokerage-sales-news")
    Catch ex As Exception

    End Try
    TOTAL_NEWS = TOTAL_NEWS
    Response.Write("<br/>" & results)
    System.Threading.Thread.Sleep(10)
    Response.Flush()
    Response.Flush()
    System.Threading.Thread.Sleep(10)
    results = ""


    yacht_news_name = ""
    'new - is an archive, not sure how much will show up or not, currently 3 months old
    results = get_yacht_news_super_yacht_news()
    TOTAL_NEWS = TOTAL_NEWS

    Response.Write("<br/>" & results)
    System.Threading.Thread.Sleep(10)
    Response.Flush()
    Response.Flush()
    System.Threading.Thread.Sleep(10)
    results = ""


    results &= "<br/><br/>"

    results &= Chr(13) & Chr(10) & "TOTALS:" & Chr(13) & Chr(10)
    results &= "Total Companies Searched: " & CStr(TOTAL_COMPANIES) & "" & Chr(13) & Chr(10)
    results &= "News Entered Into Yacht-Spot: " & CStr(TOTAL_NEWS) & "" & Chr(13) & Chr(10)
    results &= "Total Companies With New News: " & CStr(TOTAL_COMPANIES_CONNECTED) & "" & Chr(13) & Chr(10)
    results &= "Yachts Connected to Articles: " & CStr(TOTAL_YACHTS_CONNECTED) & "" & Chr(13) & Chr(10)


    Call Insert_EMail_Queue_Record(results)


    Response.Write("<br>Results: <br><br>")
    Response.Write(results)
    System.Threading.Thread.Sleep(10)
    Response.Flush()
    Response.Flush()
    System.Threading.Thread.Sleep(10)


  End Sub
  Public Function get_yachts_for_sale_YATCO(ByVal page_link As String) As String
    get_yachts_for_sale_YATCO = ""

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(page_link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim temp_yacht_id As String = ""
    Dim company_string As String = ""
    Dim didnt_find_temp As Integer = 0
    Dim found_temp As Integer = 0
    Dim insert_strings As String = ""
    Dim results_table As New DataTable
    Dim Insert_Query As String = ""
    Dim original_string_text As String = ""
    Dim related_articles_text As String = ""
    Dim yacht_count As Integer = 0
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long
    Dim results As String = ""
    Dim array_split() As String
    Dim article_link As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim connected_yachts As Integer = 0
    Dim found_companies As Integer = 0

    Dim company_count As Integer = 0
    Dim found_yacht As Integer = 0
    Dim yacht_title As String = ""
    Dim article_date As String = ""
    Dim article_text As String = ""
    Dim every_three As Integer = 1
    Dim Str2 As System.IO.Stream
    Dim srRead2 As System.IO.StreamReader
    Dim req2 As System.Net.WebRequest
    Dim resp2 As System.Net.WebResponse
    Dim original_string_text2 As String = ""
    Dim string_text3 As String = ""
    Dim string_left As String = ""
    Dim string_right As String = ""
    Dim yacht_asking_price As String

    Dim yacht_builder As String = ""



    Try

      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
      ' MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120

      ' If Right(Trim(page_link), 1) = "0" Then
      '   Call insert_into_eventlog("Yacht Pubs Started", "Research Assistant")
      ' End If


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text


      spot_to_find = InStr(string_text, "<div class=""info""")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find)

        yt_table = "<table cellspacing='0' cellpadding='4' border='1'>"
        yt_table &= "<tr><td colspan='5'><b>www.YatCo.com Listings</b></td></tr>"
        yt_table &= "<tr><td><b>Yacht Name</b></td><td><b>YachtID</b></td><td><b>Price</b></td><td><b>Status</b></td></tr>"


        array_split = Split(string_text, "<div class=""info""")

        For i = 0 To array_split.Length - 2

          string_text = array_split(i)
          original_string_text = string_text

          article_date = ""
          article_link = ""
          yacht_title = ""
          article_text = ""
          yacht_asking_price = ""
          rows_match = "N"
          yacht_builder = ""
          ypl_details = ""

          spot_to_find = InStr(string_text, "<h2>")
          string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

          spot_to_find = InStr(string_text, "</h2>")
          yacht_title = Left(string_text, spot_to_find - 1)


          spot_to_find = InStr(string_text, "<a href='/builderinfo")
          string_text = Right(string_text, Len(string_text) - spot_to_find - 11)

          spot_to_find = InStr(string_text, ">")
          string_text = Right(string_text, Len(string_text) - spot_to_find)

          spot_to_find = InStr(string_text, "</a>")
          yacht_builder = Left(string_text, spot_to_find - 1)


          spot_to_find = InStr(string_text, "<li><a href=")
          string_text = Right(string_text, Len(string_text) - spot_to_find - 11)

          spot_to_find = InStr(string_text, "'>")
          ypl_link = Left(string_text, spot_to_find - 3)
          ypl_link = Replace(ypl_link, "'", "")
          ypl_link = "http://www.yatco.com" & ypl_link


          spot_to_find = InStr(string_text, "<h3>")
          string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

          spot_to_find = InStr(string_text, "</h3>")
          yacht_asking_price = Left(string_text, spot_to_find - 1)

          If InStr(Trim(yacht_asking_price), "Price on Application") > 0 Then
            yacht_asking_price = "POA"
          End If

          'spot_to_find = InStr(string_text, "href")
          'string_text = Right(string_text, Len(string_text) - spot_to_find)

          'spot_to_find = InStr(string_text, "title")
          'ypl_link = Left(string_text, spot_to_find - 3)
          'ypl_link = Right(ypl_link, Len(ypl_link) - 6) ' for the href= and the quote
          'spot_to_find = InStr(ypl_link, """")
          'ypl_link = Left(ypl_link, spot_to_find - 1)
          'ypl_link = "http://www.superyachts.com/" & ypl_link

          'spot_to_find = InStr(string_text, ">")
          'string_text = Right(string_text, Len(string_text) - spot_to_find)

          'spot_to_find = InStr(string_text, "</a>")
          'yacht_title = Left(string_text, spot_to_find - 1)
          'yacht_title = Replace(yacht_title, "Luxery Yacht", "")
          'yacht_title = Replace(yacht_title, "'", "''")
          'yacht_title = Replace(yacht_title, "&#x27;", "''")




          '' spot_to_find = InStr(original_string_text, "<span class=""label"">Price</span>")
          'spot_to_find = InStr(original_string_text, "<span class=""main_price"">")
          'string_text = Right(original_string_text, Len(original_string_text) - spot_to_find - 25)

          'spot_to_find = InStr(string_text, "</span>")
          'yacht_asking_price = Left(string_text, spot_to_find - 1)

          'spot_to_find = InStr(original_string_text, "itemprop=""manufacturer"">")
          'string_text = Right(original_string_text, Len(original_string_text) - spot_to_find - 23)

          'spot_to_find = InStr(string_text, "</a>")
          'yacht_builder = Left(string_text, spot_to_find - 1)

          'If Len(Trim(yacht_builder)) > 50 Then
          '  spot_to_find = InStr(original_string_text, "<span class=""label"">Builder</span>")
          '  string_text = Right(original_string_text, Len(original_string_text) - spot_to_find - 33)

          '  spot_to_find = InStr(string_text, ">")
          '  string_text = Right(string_text, Len(string_text) - spot_to_find - 1)

          '  spot_to_find = InStr(string_text, "</span>")
          '  yacht_builder = Left(string_text, spot_to_find - 1)
          '  yacht_builder = Replace(yacht_builder, "<br/>", "")
          'End If
          yacht_title = Replace(yacht_title, "88&#39;", "")

          yacht_id_sy = 0
          found_pub_match = False
          found_pub_id = 0
          temp_where = ""
          ' ADD IN A CHECK FOR THE PUB TABLE, FOR THAT SOURCE, FOR THAT URL
          ' GRAB THE YACHT ID ---------------
          If CHECK_IF_IN_PUB(ypl_link, pub_yacht_id, ypl_id, "") = True Then
            found_pub_id = pub_yacht_id

            If ypl_id = 179 Then
              ypl_id = ypl_id
            End If

            If InStr(Trim(yacht_asking_price), "Price on Application") > 0 Then
              temp_where = " and ypl_yacht_info = '" & Trim(yacht_title) & "' "
            Else
              yacht_asking_temp = Replace(yacht_asking_price, "$", "")
              yacht_asking_temp = Replace(yacht_asking_temp, ",", "")
              If Trim(yacht_asking_temp) <> "" Then
                If IsNumeric(yacht_asking_temp) Then
                  yacht_asking_temp = CLng(yacht_asking_temp)
                  yacht_asking_temp = "$" & FormatNumber(yacht_asking_temp, 0)
                End If
              End If
              temp_where = " and ypl_yacht_info = '" & Trim(yacht_title) & "' and (ypl_other_info like '%" & yacht_asking_temp & "%' or (ypl_other_info = '' and ypl_process_status = 'For Sale Found  – Exact Match')) "
            End If

            If CHECK_IF_IN_PUB(ypl_link, pub_yacht_id, ypl_id, temp_where) = True Then
              ' IF IT FINDS THAT PUB RECORD, and THE PUB MATCHES THE FIELDS ASKING PRICE, THEN WE ARE CORRECT and WE CAN SKIP
              found_pub_match = True
            Else
              ' IF THE RECORD DOESNT MATCH, THEN WE NEED TO RE-OPEN THE PUB RECORD AND RE-UPDATE IT WITH THE CORRECT INFO - 
              found_pub_match = False
            End If

          Else
            ' NO PUB RECORD FOUND
            found_pub_id = 0
          End If


          If ypl_id > 0 And found_pub_match = True Then
            ' IF IT FINDS THAT PUB RECORD, and THE PUB MATCHES THE FIELDS ASKING PRICE, THEN WE ARE CORRECT and WE CAN SKIP  
            Call update_yacht_ypl("")
          ElseIf found_pub_match = False Then

            ' set variable to true to update the pub record with the correct link to the correct items

            ' IF THE RECORD DOES MATCH, THEN WE NEED TO JUST UPDATE THE DATE, AND POSSIBLY STATUS


            rows_match = CHECK_IF_FOR_SALE(yacht_title, yacht_asking_price, yacht_builder)


            yt_table &= "<tr>"
            yt_table &= "<td align='left'>" & yacht_title & "</td>"
            yt_table &= "<td align='left'>" & yacht_id_sy & "</td>"
            yt_table &= "<td align='left'>" & yacht_asking_price & "</td>"

            If Trim(rows_match) = "" Then
              Match = Match + 1
              rows_temp = "For Sale Found  – Exact Match"
            Else
              non_match = non_match + 1
              If InStr(rows_match, "NO YACHT RECORD FOUND") > 0 Then
                not_found = not_found + 1
                rows_temp = "For Sale Not Found – No Yacht Match"
              End If

              If InStr(rows_match, "MORE THAN 1 YACHT RECORD") > 0 Then
                more1_found = more1_found + 1
                rows_temp = "For Sale Not Found – Dup Matches"
              End If


              If InStr(rows_match, "YACHT NOT CURRENTLY FOR SALE") > 0 Then
                not_fs = not_fs + 1
                rows_temp = "For Sale Not Found – Record Not for Sale"

                If InStr(rows_match, "$") > 0 Then
                  ypl_details = LCase(Replace(Trim(rows_match), "YACHT NOT CURRENTLY FOR SALE IN YACHT-SPOT", ""))
                  wrong_ask = wrong_ask + 1
                End If

              ElseIf InStr(rows_match, "$") > 0 Then
                rows_temp = "For Sale Found – Price Difference"
                ypl_details = LCase(rows_match)
                wrong_ask = wrong_ask + 1
              End If


            End If

            ' YACHT NOT CURRENTLY FOR SALE
            'Response.Write("<br/>" & yacht_title & ": " & yacht_asking_price & ", " & rows_match)
            found_yt = found_yt + 1

            yt_table &= "<td align='left'>" & rows_temp & "</td>"
            yt_table &= "</tr>"





            If found_pub_id > 0 Then
              ' then we have to do an update  
              Call CHECK_PUB_PRICE_RANGE(ypl_id, yacht_asking_price)
              Call Update_yacht_ypl_fields(9, yacht_title, rows_temp, yacht_id_sy, ypl_details, ypl_link, ypl_id)
            ElseIf Not CHECK_IF_PUB_EXISTS(yacht_title, yacht_id_sy, yacht_asking_price, 9) Then
              '  '---------------------------------------------------
              Call insert_into_yacht_ypl(9, yacht_title, rows_temp, yacht_id_sy, ypl_details, ypl_link)

              new_ypl_insert = new_ypl_insert + 1
              '  '---------------------------------------------------
            Else
              ' Call update_yacht_ypl(rows_temp)
            End If
          End If



        Next
      End If

      yt_table &= "</table>"

      ' Response.Write(yt_table)

      TOTAL_NEWS = TOTAL_NEWS + found_yt


      Me.text_label.Text = results
      get_yachts_for_sale_YATCO = results

      'If page_num = 9 Then
      ' Call update_yacht_ypl_change_date(2)
      ' End If

    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()
    End Try
  End Function

  Public Function get_yachts_for_sale_SUPERYACHTS(ByVal page_num As Integer) As String
    get_yachts_for_sale_SUPERYACHTS = ""

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("http://www.superyachts.com/search-results?property_profile_type_id=1&currencyId=1&min_price=0.0&max_price=1.65E8&perPage=48&priceOrder=desc&page=" & page_num)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim temp_yacht_id As String = ""
    Dim company_string As String = ""
    Dim didnt_find_temp As Integer = 0
    Dim found_temp As Integer = 0
    Dim insert_strings As String = ""
    Dim results_table As New DataTable
    Dim Insert_Query As String = ""
    Dim original_string_text As String = ""
    Dim related_articles_text As String = ""
    Dim yacht_count As Integer = 0
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long
    Dim results As String = ""
    Dim array_split() As String
    Dim article_link As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim connected_yachts As Integer = 0
    Dim found_companies As Integer = 0

    Dim company_count As Integer = 0
    Dim found_yacht As Integer = 0
    Dim yacht_title As String = ""
    Dim article_date As String = ""
    Dim article_text As String = ""
    Dim every_three As Integer = 1
    Dim Str2 As System.IO.Stream
    Dim srRead2 As System.IO.StreamReader
    Dim req2 As System.Net.WebRequest
    Dim resp2 As System.Net.WebResponse
    Dim original_string_text2 As String = ""
    Dim string_text3 As String = ""
    Dim string_left As String = ""
    Dim string_right As String = ""
    Dim yacht_asking_price As String

    Dim yacht_builder As String = ""



    Try


      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
      ' MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120

      If page_num = 1 Then
        Call insert_into_eventlog("Yacht Pubs Started", "Research Assistant")
      End If


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text


      spot_to_find = InStr(string_text, "photoContainer")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find)

        yt_table = "<table cellspacing='0' cellpadding='4' border='1'>"
        yt_table &= "<tr><td colspan='5'><b>www.Superyachts.com Listings</b></td></tr>"
        yt_table &= "<tr><td><b>Yacht Name</b></td><td><b>YachtID</b></td><td><b>Price</b></td><td><b>Status</b></td></tr>"


        array_split = Split(string_text, "photoContainer")

        For i = 0 To array_split.Length - 1

          string_text = array_split(i)
          original_string_text = string_text

          article_date = ""
          article_link = ""
          yacht_title = ""
          article_text = ""
          yacht_asking_price = ""
          rows_match = "N"
          yacht_builder = ""
          ypl_details = ""

          spot_to_find = InStr(string_text, "href")
          string_text = Right(string_text, Len(string_text) - spot_to_find)


          spot_to_find = InStr(string_text, "href")
          string_text = Right(string_text, Len(string_text) - spot_to_find)

          spot_to_find = InStr(string_text, "title")
          ypl_link = Left(string_text, spot_to_find - 3)
          ypl_link = Right(ypl_link, Len(ypl_link) - 6) ' for the href= and the quote
          spot_to_find = InStr(ypl_link, """")
          ypl_link = Left(ypl_link, spot_to_find - 1)
          ypl_link = "http://www.superyachts.com/" & ypl_link

          spot_to_find = InStr(string_text, ">")
          string_text = Right(string_text, Len(string_text) - spot_to_find)

          spot_to_find = InStr(string_text, "</a>")
          yacht_title = Left(string_text, spot_to_find - 1)
          yacht_title = Replace(yacht_title, "Luxery Yacht", "")
          yacht_title = Replace(yacht_title, "'", "''")
          yacht_title = Replace(yacht_title, "&#x27;", "''")




          ' spot_to_find = InStr(original_string_text, "<span class=""label"">Price</span>")
          spot_to_find = InStr(original_string_text, "<span class=""main_price"">")
          string_text = Right(original_string_text, Len(original_string_text) - spot_to_find - 25)

          spot_to_find = InStr(string_text, "</span>")
          yacht_asking_price = Left(string_text, spot_to_find - 1)

          spot_to_find = InStr(original_string_text, "itemprop=""manufacturer"">")
          string_text = Right(original_string_text, Len(original_string_text) - spot_to_find - 23)

          spot_to_find = InStr(string_text, "</a>")
          yacht_builder = Left(string_text, spot_to_find - 1)

          If Len(Trim(yacht_builder)) > 50 Then
            spot_to_find = InStr(original_string_text, "<span class=""label"">Builder</span>")
            string_text = Right(original_string_text, Len(original_string_text) - spot_to_find - 33)

            spot_to_find = InStr(string_text, ">")
            string_text = Right(string_text, Len(string_text) - spot_to_find - 1)

            spot_to_find = InStr(string_text, "</span>")
            yacht_builder = Left(string_text, spot_to_find - 1)
            yacht_builder = Replace(yacht_builder, "<br/>", "")
          End If


          yacht_id_sy = 0
          found_pub_match = False
          found_pub_id = 0
          temp_where = ""
          ' ADD IN A CHECK FOR THE PUB TABLE, FOR THAT SOURCE, FOR THAT URL
          ' GRAB THE YACHT ID ---------------
          If CHECK_IF_IN_PUB(ypl_link, pub_yacht_id, ypl_id, "") = True Then
            found_pub_id = pub_yacht_id

            If ypl_id = 179 Then
              ypl_id = ypl_id
            End If

            If InStr(Trim(yacht_asking_price), "POA") > 0 Then
              temp_where = " and ypl_yacht_info = '" & Trim(yacht_title) & "' "
            Else
              yacht_asking_temp = Replace(yacht_asking_price, "$", "")
              yacht_asking_temp = Replace(yacht_asking_temp, ",", "")
              If Trim(yacht_asking_temp) <> "" Then
                If IsNumeric(yacht_asking_temp) Then
                  yacht_asking_temp = CLng(yacht_asking_temp)
                  yacht_asking_temp = "$" & FormatNumber(yacht_asking_temp, 0)
                End If
              End If
              temp_where = " and ypl_yacht_info = '" & Trim(yacht_title) & "' and (ypl_other_info like '%" & yacht_asking_temp & "%' or (ypl_other_info = '' and ypl_process_status = 'For Sale Found  – Exact Match')) "
            End If

            If CHECK_IF_IN_PUB(ypl_link, pub_yacht_id, ypl_id, temp_where) = True Then
              ' IF IT FINDS THAT PUB RECORD, and THE PUB MATCHES THE FIELDS ASKING PRICE, THEN WE ARE CORRECT and WE CAN SKIP
              found_pub_match = True
            Else
              ' IF THE RECORD DOESNT MATCH, THEN WE NEED TO RE-OPEN THE PUB RECORD AND RE-UPDATE IT WITH THE CORRECT INFO - 
              found_pub_match = False
            End If

          Else
            ' NO PUB RECORD FOUND
            found_pub_id = 0
          End If


          If ypl_id > 0 And found_pub_match = True Then
            ' IF IT FINDS THAT PUB RECORD, and THE PUB MATCHES THE FIELDS ASKING PRICE, THEN WE ARE CORRECT and WE CAN SKIP  
            Call update_yacht_ypl("")
          ElseIf found_pub_match = False Then

            ' set variable to true to update the pub record with the correct link to the correct items

            ' IF THE RECORD DOES MATCH, THEN WE NEED TO JUST UPDATE THE DATE, AND POSSIBLY STATUS


            rows_match = CHECK_IF_FOR_SALE(yacht_title, yacht_asking_price, yacht_builder)


            yt_table &= "<tr>"
            yt_table &= "<td align='left'>" & yacht_title & "</td>"
            yt_table &= "<td align='left'>" & yacht_id_sy & "</td>"
            yt_table &= "<td align='left'>" & yacht_asking_price & "</td>"

            If Trim(rows_match) = "" Then
              Match = Match + 1
              rows_temp = "For Sale Found  – Exact Match"
            Else
              non_match = non_match + 1
              If InStr(rows_match, "NO YACHT RECORD FOUND") > 0 Then
                not_found = not_found + 1
                rows_temp = "For Sale Not Found – No Yacht Match"
              End If

              If InStr(rows_match, "MORE THAN 1 YACHT RECORD") > 0 Then
                more1_found = more1_found + 1
                rows_temp = "For Sale Not Found – Dup Matches"
              End If


              If InStr(rows_match, "YACHT NOT CURRENTLY FOR SALE") > 0 Then
                not_fs = not_fs + 1
                rows_temp = "For Sale Not Found – Record Not for Sale"

                If InStr(rows_match, "$") > 0 Then
                  ypl_details = LCase(Replace(Trim(rows_match), "YACHT NOT CURRENTLY FOR SALE IN YACHT-SPOT", ""))
                  wrong_ask = wrong_ask + 1
                End If

              ElseIf InStr(rows_match, "$") > 0 Then
                rows_temp = "For Sale Found – Price Difference"
                ypl_details = LCase(rows_match)
                wrong_ask = wrong_ask + 1
              End If


            End If

            ' YACHT NOT CURRENTLY FOR SALE
            'Response.Write("<br/>" & yacht_title & ": " & yacht_asking_price & ", " & rows_match)
            found_yt = found_yt + 1

            yt_table &= "<td align='left'>" & rows_temp & "</td>"
            yt_table &= "</tr>"
 

            If found_pub_id > 0 Then
              ' then we have to do an update  
              Call CHECK_PUB_PRICE_RANGE(ypl_id, yacht_asking_price)
              Call Update_yacht_ypl_fields(2, yacht_title, rows_temp, yacht_id_sy, ypl_details, ypl_link, ypl_id)
            ElseIf Not CHECK_IF_PUB_EXISTS(yacht_title, yacht_id_sy, yacht_asking_price, 2) Then
              '  '---------------------------------------------------
              Call insert_into_yacht_ypl(2, yacht_title, rows_temp, yacht_id_sy, ypl_details, ypl_link)

              new_ypl_insert = new_ypl_insert + 1
              '  '---------------------------------------------------
            Else
              Call update_yacht_ypl(rows_temp)
            End If
          End If



        Next
      End If

      yt_table &= "</table>"

      ' Response.Write(yt_table)

      TOTAL_NEWS = TOTAL_NEWS + found_yt


      Me.text_label.Text = results
      get_yachts_for_sale_SUPERYACHTS = results

      If page_num = 9 Then
        Call update_yacht_ypl_change_date(2)
      End If

    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()
    End Try
  End Function
  Public Function CHECK_IF_IN_PUB(ByVal link As String, ByRef pub_yacht_id As Long, ByRef yacht_pub_id As Long, ByVal where_clause As String) As Boolean
    CHECK_IF_IN_PUB = False
    Dim atemptable As New DataTable

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader
    Dim Query As String = ""
    Dim original_found As Integer = 0


    Try

      SqlConn.ConnectionString = MySqlConn_JETNET.ConnectionString
      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      Query = "select ypl_yacht_id, ypl_id from Yacht_Publication_Log with (NOLOCK) where ypl_source_url = '" & link & "'   "

      If Trim(where_clause) <> "" Then
        Query &= " " & where_clause
      End If

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        atemptable.Load(SqlReader)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

      If atemptable.Rows.Count > 0 Then
        For Each r As DataRow In atemptable.Rows
          pub_yacht_id = r.Item("ypl_yacht_id")
          yacht_pub_id = r.Item("ypl_id")
          CHECK_IF_IN_PUB = True
        Next
      End If

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try


  End Function
  Public Function CHECK_IF_FOR_SALE(ByVal yacht_title As String, ByVal yacht_asking As String, ByVal yacht_builder As String) As String
    CHECK_IF_FOR_SALE = ""
    Dim atemptable As New DataTable

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader
    Dim Query As String = ""
    Dim original_found As Integer = 0


    Try
      yacht_asking = Replace(yacht_asking, "$", "")
      yacht_asking = Replace(yacht_asking, ",", "")
      If Trim(yacht_asking) <> "" Then
        If IsNumeric(yacht_asking) Then
          yacht_asking = CLng(yacht_asking)
        End If
      End If

      SqlConn.ConnectionString = MySqlConn_JETNET.ConnectionString
      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      Query = "select yt_forsale_flag, yt_asking_price, yt_id from Yacht with (NOLOCK) where yt_yacht_name = '" & yacht_title & "' and yt_journ_id = 0 "

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader()

      Try
        atemptable.Load(SqlReader)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

      If atemptable.Rows.Count = 0 Then
        Query = "select yt_forsale_flag, yt_asking_price, yt_id from Yacht with (NOLOCK) where yt_yacht_name like '%" & yacht_title & "%' and yt_journ_id = 0 "

        SqlCommand.CommandText = Query.ToString
        SqlReader = SqlCommand.ExecuteReader()

        Try
          atemptable.Load(SqlReader)
        Catch constrExc As System.Data.ConstraintException
          Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
          ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
        End Try

        If atemptable.Rows.Count > 1 And Trim(yacht_builder) <> "" Then

          cutme(yacht_builder)

          atemptable.Clear()
          Query = "select yt_forsale_flag, yt_asking_price, yt_id from Yacht with (NOLOCK) "
          Query = Query & " inner join Yacht_Model with (NOLOCK) on yt_model_id = ym_model_id "
          Query = Query & " where yt_yacht_name like '%" & yacht_title & "%' and yt_journ_id = 0 "
          Query = Query & " and ( ym_brand_name = '" & Trim(yacht_builder) & "' "
          Query = Query & " or  ym_brand_name like '" & Trim(yacht_builder) & "%' "
          Query = Query & " or LEFT(ym_brand_name, 8) = LEFT('" & Trim(yacht_builder) & "', 8)   )"

          SqlCommand.CommandText = Query.ToString
          SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

          Try
            atemptable.Load(SqlReader)
          Catch constrExc As System.Data.ConstraintException
            Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
            ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
          End Try
        End If

      ElseIf atemptable.Rows.Count > 1 And Trim(yacht_builder) <> "" Then
        original_found = atemptable.Rows.Count
        cutme(yacht_builder)
        atemptable.Clear()
        Query = "select yt_forsale_flag, yt_asking_price, yt_id from Yacht with (NOLOCK) "
        Query = Query & " inner join Yacht_Model with (NOLOCK) on yt_model_id = ym_model_id "
        Query = Query & " where yt_yacht_name = '" & yacht_title & "' and yt_journ_id = 0 "
        Query = Query & " and ( ym_brand_name = '" & Trim(yacht_builder) & "' "
        Query = Query & " or  ym_brand_name like '" & Trim(yacht_builder) & "%' "
        Query = Query & " or LEFT(ym_brand_name, 8) = LEFT('" & Trim(yacht_builder) & "', 8)   )"

        SqlCommand.CommandText = Query.ToString
        SqlReader = SqlCommand.ExecuteReader()

        Try
          atemptable.Load(SqlReader)
        Catch constrExc As System.Data.ConstraintException
          Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
          ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
        End Try

      End If





      ' it is more than 1, do not do it 
      If atemptable.Rows.Count > 1 Then
        CHECK_IF_FOR_SALE = " MORE THAN 1 YACHT RECORD"
      ElseIf atemptable.Rows.Count = 1 Then
        For Each r As DataRow In atemptable.Rows
          If Trim(r.Item("yt_forsale_flag")) = "Y" Then
          Else
            CHECK_IF_FOR_SALE &= " YACHT NOT CURRENTLY FOR SALE IN YACHT-SPOT"
          End If

          If IsDBNull(r.Item("yt_asking_price")) Then
            If InStr(Trim(yacht_asking), "POA") > 0 Then
            Else
              CHECK_IF_FOR_SALE &= " $" & yacht_asking & " vs. NONE"
            End If
          ElseIf Trim(r.Item("yt_asking_price")) = Trim(yacht_asking) Then

          ElseIf IsNumeric(Trim(r.Item("yt_asking_price"))) = True And IsNumeric(Trim(yacht_asking)) = True Then ' if there are both numeric.. compare

            'if they are within 3 percent of each other
            If (CInt(CInt(Trim(r.Item("yt_asking_price")) * percent_off) + CInt(Trim(r.Item("yt_asking_price")))) > CInt(Trim(yacht_asking))) And CInt(CInt(Trim(yacht_asking) * percent_off) + CInt(Trim(yacht_asking))) > CInt(Trim(r.Item("yt_asking_price"))) Then
              ' within range 
              yacht_asking = yacht_asking
            ElseIf InStr(Trim(yacht_asking), "POA") > 0 Then
            Else
              CHECK_IF_FOR_SALE &= " $" & FormatNumber(yacht_asking, 0) & " vs. $" & FormatNumber(Trim(r.Item("yt_asking_price")), 0) & ""
            End If


          Else
            If InStr(Trim(yacht_asking), "POA") > 0 Then
            Else
              CHECK_IF_FOR_SALE &= " $" & FormatNumber(yacht_asking, 0) & " vs. $" & FormatNumber(Trim(r.Item("yt_asking_price")), 0) & ""
            End If
          End If

          yacht_id_sy = r.Item("yt_id")

        Next
      ElseIf original_found > 0 And atemptable.Rows.Count = 0 Then
        CHECK_IF_FOR_SALE = " MORE THAN 1 YACHT RECORD"
      ElseIf atemptable.Rows.Count = 0 Then
        CHECK_IF_FOR_SALE = " NO YACHT RECORD FOUND"
      End If

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try


  End Function
  Public Sub cutme(ByRef temp_val As String)

    If Trim(temp_val) <> "" Then
      For i = 0 To 100
        If Trim(temp_val) <> "" Then
          If Asc(Left(Trim(temp_val), 1)) < 32 Or Asc(Left(Trim(temp_val), 1)) > 255 Then
            temp_val = Right(Trim(temp_val), Len(Trim(temp_val)) - 1)
          Else
            i = 100
          End If
        End If
      Next


      For i = 0 To 100
        If Trim(temp_val) <> "" Then
          If Asc(Right(Trim(temp_val), 1)) < 32 Or Asc(Right(Trim(temp_val), 1)) > 255 Then
            temp_val = Left(Trim(temp_val), Len(Trim(temp_val)) - 1)
          Else
            i = 100
          End If
        End If
      Next
    End If


  End Sub

  Public Sub cutme_LF(ByRef temp_val As String)

    If Trim(temp_val) <> "" Then
      For i = 0 To 100
        If Trim(temp_val) <> "" Then
          If Asc(Left(Trim(temp_val), 1)) = 10 Or Asc(Left(Trim(temp_val), 1)) = 13 Then
            temp_val = Right(Trim(temp_val), Len(Trim(temp_val)) - 1)
          Else
            i = 100
          End If
        End If
      Next


      For i = 0 To 100
        If Trim(temp_val) <> "" Then
          If Asc(Left(Trim(temp_val), 1)) = 10 Or Asc(Left(Trim(temp_val), 1)) = 13 Then
            temp_val = Left(Trim(temp_val), Len(Trim(temp_val)) - 1)
          Else
            i = 100
          End If
        End If
      Next
    End If


  End Sub

  Public Function get_super_yacht_news2(ByVal link As String) As String
    get_super_yacht_news2 = ""

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim temp_yacht_id As String = ""
    Dim company_string As String = ""
    Dim didnt_find_temp As Integer = 0
    Dim found_temp As Integer = 0
    Dim insert_strings As String = ""
    Dim results_table As New DataTable
    Dim Insert_Query As String = ""
    Dim original_string_text As String = ""
    Dim related_articles_text As String = ""
    Dim yacht_count As Integer = 0
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long
    Dim results As String = ""
    Dim array_split() As String
    Dim article_link As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim connected_yachts As Integer = 0
    Dim found_companies As Integer = 0
    Dim found_news As Integer = 0
    Dim company_count As Integer = 0
    Dim found_yacht As Integer = 0
    Dim article_title As String = ""
    Dim article_date As String = ""
    Dim article_text As String = ""
    Dim every_three As Integer = 1
    Dim Str2 As System.IO.Stream
    Dim srRead2 As System.IO.StreamReader
    Dim req2 As System.Net.WebRequest
    Dim resp2 As System.Net.WebResponse
    Dim original_string_text2 As String = ""
    Dim string_text3 As String = ""
    Dim string_left As String = ""
    Dim string_right As String = ""


    Try


      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120



      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text


      spot_to_find = InStr(string_text, "<h3")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find)


        array_split = Split(string_text, "<h3")

        '  For i = 1 To array_split.Length - 1
        For i = 1 To array_split.Length - 1

          string_text = array_split(i)

          spot_to_find = InStr(string_text, "href")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 5)
          End If

          spot_to_find = InStr(string_text, "class")
          If spot_to_find > 0 Then
            article_link = Left(string_text, spot_to_find - 3)
            article_link = "http://www.superyachtnews.com" & article_link
            string_text = Right(string_text, Len(string_text) - spot_to_find - 5)
          End If

          spot_to_find = InStr(string_text, ">")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find)
          End If

          spot_to_find = InStr(string_text, "</a>")
          If spot_to_find > 0 Then
            article_title = Left(string_text, spot_to_find - 1)
            string_text = Right(string_text, Len(string_text) - spot_to_find)
          End If

          If InStr(article_title, "<div>") > 0 Then  ' mssed up, so skip it 
          Else

            Try


              req2 = System.Net.WebRequest.Create(article_link)
              resp2 = req2.GetResponse

              Str2 = resp2.GetResponseStream
              srRead2 = New System.IO.StreamReader(Str2)
              string_text3 = srRead2.ReadToEnd().ToString
              string_text2 = string_text3
              original_string_text2 = string_text3

              If InStr(Trim(string_text2), "dateModified") > 0 Then
                string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), "dateModified") - 17)

                If InStr(Trim(string_text2), ">") > 0 Then
                  string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), ">"))

                  If InStr(Trim(string_text2), "</span>") > 0 Then
                    string_text2 = Left(string_text2, InStr(Trim(string_text2), "</span>") - 1)
                    article_date = string_text2
                    article_date = fix_date(article_date, "")
                  End If
                End If
              End If

              string_text2 = original_string_text2
              If InStr(Trim(string_text2), "articleBody") > 0 Then
                string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), "articleBody") - 17)

                If InStr(Trim(string_text2), "<p>") > 0 Then
                  string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), "<p>") - 2)

                  If InStr(Trim(string_text2), "</p>") > 0 Then
                    string_text2 = Left(string_text2, InStr(Trim(string_text2), "</p>") - 1)
                    article_text = string_text2
                    article_text = Replace(article_text, "&ndash;", "-")
                  End If
                End If
              End If


              If Not CHECK_IF_NEWS_EXISTS(0, article_date, Replace(article_title, "'", "''"), True, "") Then
                '---------------------------------------------------
                Call insert_into_news(article_date, article_title, article_text, article_link, 0, 0, 0, True, "7")

                found_news = found_news + 1
                '---------------------------------------------------
              End If



            Catch ex As Exception

            End Try

            article_date = ""
            article_link = ""
            article_title = ""
            article_text = ""
          End If


        Next
      End If



      results = Chr(13) & Chr(10) & "Boat International - http://www.boatinternational.com/yachts/news:" & Chr(13) & Chr(10)
      results &= "News Entered Into Yacht-Spot: " & CStr(found_news) & "" & Chr(13) & Chr(10)


      TOTAL_NEWS = TOTAL_NEWS + found_news


      Me.text_label.Text = results
      get_super_yacht_news2 = results



    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()
    End Try
  End Function
  Public Function get_boat_international_news(ByVal temp_path As String) As String
    get_boat_international_news = ""

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(temp_path)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim temp_yacht_id As String = ""
    Dim company_string As String = ""
    Dim didnt_find_temp As Integer = 0
    Dim found_temp As Integer = 0
    Dim insert_strings As String = ""
    Dim results_table As New DataTable
    Dim Insert_Query As String = ""
    Dim original_string_text As String = ""
    Dim related_articles_text As String = ""
    Dim yacht_count As Integer = 0
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long
    Dim results As String = ""
    Dim array_split() As String
    Dim article_link As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim connected_yachts As Integer = 0
    Dim found_companies As Integer = 0
    Dim found_news As Integer = 0
    Dim company_count As Integer = 0
    Dim found_yacht As Integer = 0
    Dim article_title As String = ""
    Dim article_date As String = ""
    Dim article_text As String = ""
    Dim every_three As Integer = 1
    Dim Str2 As System.IO.Stream
    Dim srRead2 As System.IO.StreamReader
    Dim req2 As System.Net.WebRequest
    Dim resp2 As System.Net.WebResponse
    Dim original_string_text2 As String = ""
    Dim string_text3 As String = ""
    Dim string_left As String = ""
    Dim string_right As String = ""


    Try


      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
      ' MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120



      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text

      ' commented out since change 9/12/17
      spot_to_find = InStr(string_text, "itemprop=""name""")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find)


        ' array_split = Split(string_text, "itemprop=""headline"" itemprop=""name""") ' should be just last section
        ' array_split = Split(string_text, "itemprop=""name""")
        array_split = Split(string_text, "</h5></div></div>")

        '  For i = 1 To array_split.Length - 1
        For i = 0 To array_split.Length - 1

          '   If every_three = 1 Then
          string_text = array_split(i)

          If i >= 1 Then
            spot_to_find = InStr(string_text, "Yacht News")
            If spot_to_find > 0 Then
              string_text = Right(string_text, Len(string_text) - spot_to_find - 9)
            End If
          End If

          ' get rid of the first one 
          ' currently only the detail articles, have https on them, may need to look for soemthing differernt in the future 
          spot_to_find = InStr(string_text, "https://www.boatinternational.com/")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 33)
          End If

          'grab the secnod one 
          ' spot_to_find = InStr(string_text, "https://www.boatinternational.com/")
          article_link = ""
          spot_to_find = InStr(string_text, "class=")
          If spot_to_find > 0 Then
            article_link = Left(string_text, spot_to_find - 2)
            article_link = Replace(article_link, """", "")
            article_link = "https://www.boatinternational.com/" & Trim(article_link)
          End If

          ' Left(Trim(article_link), 42) = "https://www.boatinternational.com/archive/"
          If InStr(Trim(article_link), "/archive/") > 0 Or Trim(article_link) = "https://www.boatinternational.com/yachts-for-sale/results" Or Trim(article_text) = "https://www.boatinternational.com/archive/about-news" Then
            ' then re-do the link section 
            spot_to_find = InStr(string_text, "https://www.boatinternational.com/")
            If spot_to_find > 0 Then
              string_text = Right(string_text, Len(string_text) - spot_to_find - 33)
            End If

            'grab the secnod one 
            ' spot_to_find = InStr(string_text, "https://www.boatinternational.com/")
            article_link = ""
            spot_to_find = InStr(string_text, "class=")
            If spot_to_find > 0 Then
              article_link = Left(string_text, spot_to_find - 2)
              article_link = Replace(article_link, """", "")
              article_link = "https://www.boatinternational.com/" & Trim(article_link)
            End If

          End If



          article_title = ""
          spot_to_find = InStr(string_text, "itemprop=""name""")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 15)
            string_text = Replace(string_text, ">", "")
            article_title = Trim(string_text)

            If InStr(Trim(article_title), "</h5") > 0 Then
              spot_to_find = InStr(article_title, "</h5")
              If spot_to_find > 0 Then
                article_title = Left(article_title, spot_to_find - 1)
              End If
            End If
          End If

          'spot_to_find = InStr(string_text, "</h5>") ' h5
          'If spot_to_find > 0 Then
          '  article_title = Left(string_text, spot_to_find - 1)
          '  article_title = Replace(article_title, "> ", "")
          'End If

          'spot_to_find = InStr(string_text, " class=""js-gtm-track lis")
          '' spot_to_find = InStr(string_text, "js-gtm-track content-item__text-link")
          'If spot_to_find > 0 Then
          '  'string_text = Left(string_text, spot_to_find - 10)
          '  string_text = Left(string_text, spot_to_find - 1)
          'End If



          Try

            req2 = System.Net.WebRequest.Create(article_link)
            resp2 = req2.GetResponse

            Str2 = resp2.GetResponseStream
            srRead2 = New System.IO.StreamReader(Str2)
            string_text3 = srRead2.ReadToEnd().ToString
            string_text2 = string_text3
            original_string_text2 = string_text3

            If InStr(Trim(string_text2), "datePublished") > 0 Then
              string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), "datePublished") - 17)

              If InStr(Trim(string_text2), ">") > 0 Then
                string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), ">") - 1)

                If InStr(Trim(string_text2), "</time>") > 0 Then
                  string_text2 = Left(string_text2, InStr(Trim(string_text2), "</time>") - 1)
                  article_date = string_text2
                  article_date = fix_date(article_date, "")
                End If
              End If
            End If

            string_text2 = original_string_text2
            If InStr(Trim(string_text2), "articleBody") > 0 Then
              string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), "articleBody") - 17)

              If InStr(Trim(string_text2), "<p>") > 0 Then
                string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), "<p>") - 2)

                If InStr(Trim(string_text2), "</p>") > 0 Then
                  string_text2 = Left(string_text2, InStr(Trim(string_text2), "</p>") - 1)
                  article_text = string_text2


                  If InStr(article_text, "<a href=") > 0 Then
                    string_left = Left(Trim(article_text), InStr(article_text, "<a href=") - 1)
                    string_right = Right(Trim(article_text), Len(article_text) - InStr(article_text, "<a href=") - 8)

                    If InStr(string_right, ">") > 0 Then
                      string_right = Right(Trim(string_right), Len(string_right) - InStr(string_right, ">"))
                      string_right = Replace(string_right, "</a>", "")
                    End If
                    article_text = string_left & string_right
                  End If

                  ' do it a second time 
                  If InStr(article_text, "<a href=") > 0 Then
                    string_left = Left(Trim(article_text), InStr(article_text, "<a href=") - 1)
                    string_right = Right(Trim(article_text), Len(article_text) - InStr(article_text, "<a href=") - 8)

                    If InStr(string_right, ">") > 0 Then
                      string_right = Right(Trim(string_right), Len(string_right) - InStr(string_right, ">"))
                      string_right = Replace(string_right, "</a>", "")
                    End If
                    article_text = string_left & string_right
                  End If


                End If
              End If
            End If


            If article_date = "" And article_text = "" Then
              spot_to_find = InStr(Trim(string_text2), "gallery__text-description-short")
              If spot_to_find > 0 Then
                string_right = Right(Trim(string_text2), Len(string_text2) - spot_to_find - 33)

                spot_to_find = InStr(string_right, "<")
                If spot_to_find > 0 Then
                  string_right = Left(Trim(string_right), spot_to_find - 1)
                  article_text = string_right
                End If

              End If


            End If



            If Not CHECK_IF_NEWS_EXISTS(0, article_date, Replace(article_title, "'", "''"), True, "") And Trim(article_title) <> "" Then
              '---------------------------------------------------

              If article_date = "" Then
                article_date = FormatDateTime(Date.Now(), DateFormat.ShortDate)
              End If

              yacht_news_name = article_title
              Call insert_into_news(article_date, article_title, article_text, article_link, 0, 0, 0, True, "4")

              found_news = found_news + 1
              '---------------------------------------------------
            End If

          Catch ex As Exception

          End Try

          'item__text-link

          article_date = ""
          article_link = ""
          article_title = ""
          article_text = ""
          '  End If

          'every_three = every_three + 1

          '  If every_three = 4 Then
          'every_three = 1
          '   End If
        Next
      End If



      results = Chr(13) & Chr(10) & "Boat International - https://www.boatinternational.com/yachts/news:" & Chr(13) & Chr(10)
      results &= "News Entered Into Yacht-Spot: " & CStr(found_news) & "" & Chr(13) & Chr(10)


      TOTAL_NEWS = TOTAL_NEWS + found_news


      Me.text_label.Text = results
      get_boat_international_news = results



    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()
    End Try
  End Function

  Public Function Scrape_This_Page_Yacht_Paging() As String
    Scrape_This_Page_Yacht_Paging = ""

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("http://www.yachting-pages.com/superyacht_news/")
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim temp_yacht_id As String = ""
    Dim company_string As String = ""
    Dim didnt_find_temp As Integer = 0
    Dim found_temp As Integer = 0
    Dim insert_strings As String = ""
    Dim results_table As New DataTable
    Dim Insert_Query As String = ""
    Dim original_string_text As String = ""
    Dim related_articles_text As String = ""
    Dim yacht_count As Integer = 0
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long
    Dim results As String = ""
    Dim array_split() As String
    Dim article_link As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim connected_yachts As Integer = 0
    Dim found_companies As Integer = 0
    Dim found_news As Integer = 0
    Dim company_count As Integer = 0
    Dim found_yacht As Integer = 0
    Dim article_title As String = ""
    Dim article_date As String = ""
    Dim article_text As String = ""

    Try


      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120



      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text


      spot_to_find = InStr(string_text, "<div class=""listing-news"">")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find)


        array_split = Split(string_text, "<article class=""item""")

        '  For i = 1 To array_split.Length - 1
        For i = 1 To array_split.Length - 1
          string_text = array_split(i)

          spot_to_find = InStr(string_text, "<a href=")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 8)

            spot_to_find = InStr(string_text, "itemprop=")
            If spot_to_find > 0 Then
              article_link = Left(string_text, spot_to_find - 3)

              string_text = Right(string_text, Len(string_text) - spot_to_find - 1)

              spot_to_find = InStr(string_text, "</a>")
              If spot_to_find > 0 Then
                article_title = Left(string_text, spot_to_find - 1)

                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)

                spot_to_find = InStr(article_title, ">")
                If spot_to_find > 0 Then
                  article_title = Right(article_title, Len(article_title) - spot_to_find)
                End If


                spot_to_find = InStr(string_text, "datePublished")
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 16)

                  spot_to_find = InStr(string_text, ">")
                  If spot_to_find > 0 Then

                    string_text = Right(string_text, Len(string_text) - spot_to_find)

                    spot_to_find = InStr(string_text, "</a>")
                    If spot_to_find > 0 Then
                      article_date = Left(string_text, spot_to_find - 1)
                      article_date = fix_date(article_date, "yacht_paging")

                      spot_to_find = InStr(string_text, "articleBody")
                      If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 13)
                        spot_to_find = InStr(string_text, "<a href")
                        If spot_to_find > 0 Then
                          string_text = Left(string_text, spot_to_find - 1)
                          article_text = string_text


                          If Not CHECK_IF_NEWS_EXISTS(0, article_date, Replace(article_title, "'", "''"), True, "") Then
                            '---------------------------------------------------
                            Call insert_into_news(article_date, article_title, article_text, article_link, 0, 0, 0, True, "6")

                            found_news = found_news + 1
                            '---------------------------------------------------
                          End If



                        End If

                      End If
                    End If


                  End If


                End If

              End If

            End If
          End If

        Next
      End If


      results = Chr(13) & Chr(10) & "Super Yacht Business - www.superyachtbusiness.net:" & Chr(13) & Chr(10)
      results &= "News Entered Into Yacht-Spot: " & CStr(found_news) & "" & Chr(13) & Chr(10)


      TOTAL_NEWS = TOTAL_NEWS + found_news


      Me.text_label.Text = results
      Scrape_This_Page_Yacht_Paging = results

    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()
    End Try
  End Function

  Public Function Scrape_This_Page_Super_Yacht_Business() As String
    Scrape_This_Page_Super_Yacht_Business = ""

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("http://www.superyachtbusiness.net/news")
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim temp_yacht_id As String = ""
    Dim company_string As String = ""
    Dim didnt_find_temp As Integer = 0
    Dim found_temp As Integer = 0
    Dim insert_strings As String = ""
    Dim results_table As New DataTable
    Dim Insert_Query As String = ""
    Dim original_string_text As String = ""
    Dim related_articles_text As String = ""
    Dim yacht_count As Integer = 0
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long
    Dim results As String = ""
    Dim array_split() As String
    Dim article_link As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim connected_yachts As Integer = 0
    Dim found_companies As Integer = 0
    Dim found_news As Integer = 0
    Dim company_count As Integer = 0
    Dim found_yacht As Integer = 0
    Dim article_title As String = ""
    Dim article_date As String = ""
    Dim article_text As String = ""


    Try


      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120



      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text


      spot_to_find = InStr(string_text, "main-loop-wraper")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find)

        '<header class="entry-header"> 
        '			<h2 class="entry-title sub-heading"><a href="http://www.superyachtbusiness.net/new-products/ypi-introduces-virtual-reality-technology-to-new-build-projects-9532" rel="bookmark" itemprop="url"><span itemprop="name">YPI introduces virtual reality technology to new build projects</span></a></h2>

        '							<ul class="entry-meta">
        '																<li class="entry-date">
        '															12:03 pm - March 15, 2016								<meta itemprop="datePublished" content="">
        '													</li>
        '																			</ul> 
        '		</header> 
        '		<div class="entry-content"> 
        '			<p>New Oculus virtual reality technology enables owners to experience how the layouts feel in terms of space and volume before the yacht is built</p>
        '		</div> 
        '		<footer>

        array_split = Split(string_text, "entry-header")

        '  For i = 1 To array_split.Length - 1
        For i = 1 To array_split.Length - 1
          string_text = array_split(i)

          spot_to_find = InStr(string_text, "<a href=")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 8)

            spot_to_find = InStr(string_text, "rel=")
            If spot_to_find > 0 Then
              article_link = Left(string_text, spot_to_find - 3)

              string_text = Right(string_text, Len(string_text) - spot_to_find - 1)

              spot_to_find = InStr(string_text, "</span>")
              If spot_to_find > 0 Then
                article_title = Left(string_text, spot_to_find - 1)

                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)

                spot_to_find = InStr(article_title, "<span itemprop=""name"">")
                If spot_to_find > 0 Then
                  article_title = Right(article_title, Len(article_title) - spot_to_find - 21)
                End If


                spot_to_find = InStr(string_text, "<li class=""entry-date"">")
                string_text = Right(string_text, Len(string_text) - spot_to_find - 23)

                spot_to_find = InStr(string_text, "<meta")
                If spot_to_find > 0 Then
                  article_date = Left(string_text, spot_to_find - 1)
                  spot_to_find = InStr(article_date, " - ")
                  article_date = Right(article_date, Len(article_date) - spot_to_find - 2)
                  If IsDate(article_date) Then
                    article_date = CDate(article_date)
                  End If
                  article_date = Replace(article_date, " ", "")
                End If



                spot_to_find = InStr(string_text, "<p>")
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 2)

                  spot_to_find = InStr(string_text, "</p>")
                  If spot_to_find > 0 Then
                    article_text = Left(string_text, spot_to_find - 1)
                  End If

                  If Not CHECK_IF_NEWS_EXISTS(0, article_date, Replace(article_title, "'", "''"), True, "") Then
                    '  '---------------------------------------------------
                    Call insert_into_news(article_date, article_title, article_text, article_link, 0, 0, 0, True, "3")

                    found_news = found_news + 1
                    '  '---------------------------------------------------
                  End If


                End If

              End If
            End If


          End If
        Next
      End If

      'results = "<br><table cellspacing='0' cellpadding='0' border='0' valign='top'>"
      'results &= "<tr><Td align='left'><font color='black'>Super Yacht Business</font></td></tr>"
      'results &= "<tr><Td align='left'><font color='black'>News Entered Into Yacht-Spot: " & CStr(found_news) & "</font>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"
      'results &= "</table>"
      'results &= "</td>"

      'results &= "</td></tr></table>"



      results = Chr(13) & Chr(10) & "Super Yacht Business - www.superyachtbusiness.net:" & Chr(13) & Chr(10)
      results &= "News Entered Into Yacht-Spot: " & CStr(found_news) & "" & Chr(13) & Chr(10)


      TOTAL_NEWS = TOTAL_NEWS + found_news


      Me.text_label.Text = results
      Scrape_This_Page_Super_Yacht_Business = results

    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()
    End Try
  End Function



  Public Function get_yacht_news_super_yacht_news()
    get_yacht_news_super_yacht_news = ""
    Dim i As Integer = 0

    Try


      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120

      For i = 1 To 1
        Scrape_This_Page_Super_Yacht_News("http://www.superyachtnews.com/newsarchive.html?fleet", 1)
      Next

      Call insert_into_eventlog("Yacht News Finished", "Research Assistant")

    Catch ex As Exception
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()
    End Try
  End Function


  Public Function scrape_for_mmsi(ByVal page_num As Long)
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/ship_type:9/length_between:23%2C450/per_page:50")
    ' Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/ship_type:9/length_between:23%2C25/per_page:50")
    ' Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/ship_type:9/length_between:25%2C29/per_page:50")
    ' Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/ship_type:9/length_between:29%2C37/per_page:50")
    ' Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/ship_type:9/length_between:37%2C60/per_page:50")
    ' Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/ship_type:9/length_between:60%2C450/per_page:50")



    '  Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/length_between:25%2C450/ship_type:9/per_page:50/flag:MH")   '6
    '  Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/length_between:25%2C450/ship_type:9/per_page:50/flag:KY")    '20
    '  Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/length_between:25%2C450/ship_type:9/per_page:50/flag:GB")    '16
    '  Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/length_between:25%2C450/ship_type:9/per_page:50/flag:VG")  '3
    ' Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/length_between:25%2C450/ship_type:9/per_page:50/flag:IT")   '3
    ' Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/length_between:25%2C450/ship_type:9/per_page:50/flag:MT")  '9
    '  Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.marinetraffic.com/en/ais/index/ships/all/page:" & page_num & "/length_between:25%2C450/ship_type:9/per_page:50/flag:US")  '24



    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim original_string_text As String = ""
    Dim article_link As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim Str2 As System.IO.Stream
    Dim array_split() As String
    Dim array_split2() As String
    Dim k As Integer = 0
    Dim temp_yacht_name As String = ""
    Dim temp_yacht_flag As String = ""
    Dim temp_mssi As String = ""
    Dim temp_imo As String = ""
    Dim temp_exists As Boolean = False

    Dim temp_action As String = ""
    Dim temp_vessel_type As String = ""
    Dim skip_this As Boolean = False
    Dim extra_note As String = ""



    Try


      ' MySqlConn_JETNET.ConnectionString = "Data Source=www.jetnetsql2.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
      '' MySqlConn_JETNET2.ConnectionString = "Data Source=www.jetnetsql2.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
      '   MySqlConn_JETNET.ConnectionString = "Data Source=128.1.21.40;Initial Catalog=jetnet_ra_test;Persist Security Info=True;User ID=sa;Password=moejive"
      ' MySqlConn_JETNET.ConnectionString = "Data Source=128.1.21.200;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=sa;Password=moejive"

      ' MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
      'MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN

      'MySqlConn_JETNET.Open()
      'MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      'MySqlCommand_JETNET.CommandType = CommandType.Text
      'MySqlCommand_JETNET.CommandTimeout = 120




      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text


      spot_to_find = InStr(string_text, "My Fleet")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find)

        spot_to_find = InStr(string_text, "</tr>")
        If spot_to_find > 0 Then
          string_text = Right(string_text, Len(string_text) - spot_to_find - 2)


          array_split = Split(string_text, "<tr>")

          For i = 1 To array_split.Length - 1
            temp_yacht_name = ""
            temp_mssi = ""
            temp_imo = ""

            string_text = array_split(i)
            original_string_text = string_text

            array_split2 = Split(string_text, "</td>")

            ' For k = 0 To array_split2.Length - 1


            ' 0 is flag 
            temp_yacht_flag = array_split2(0)
            '" 					<td> 						<span><img title="Marshall Is" alt="Marshall Is" src="/img/flags/png40/MH.png"></span>					" 

            If InStr(Trim(temp_yacht_flag), "title=") > 0 Then
              temp_yacht_flag = Right(Trim(temp_yacht_flag), Len(Trim(temp_yacht_flag)) - InStr(temp_yacht_flag, "title=") - 6)
              If InStr(Trim(temp_yacht_flag), "alt=") > 0 Then
                temp_yacht_flag = Left(Trim(temp_yacht_flag), InStr(Trim(temp_yacht_flag), "alt=") - 3)
              End If
            End If



            ' 1 is id/ imo
            temp_imo = array_split2(1)
            temp_imo = Replace(temp_imo, "<td>", "")
            temp_imo = Replace(temp_imo, "IMO: ", "")
            cutme(temp_imo)

            ' 2 is mmsi
            temp_mssi = array_split2(2)
            temp_mssi = Replace(temp_mssi, "<td>", "")
            cutme(temp_mssi)

            ' 3 is vessal name
            temp_yacht_name = array_split2(3)
            If InStr(Trim(temp_yacht_name), "</a>") > 0 Then
              temp_yacht_name = Replace(Trim(temp_yacht_name), "<td>", "")
              temp_yacht_name = Left(Trim(temp_yacht_name), InStr(Trim(temp_yacht_name), "</a>") - 1)
              If InStr(Trim(temp_yacht_name), ">") > 0 Then
                temp_yacht_name = Right(Trim(temp_yacht_name), Len(Trim(temp_yacht_name)) - InStr(temp_yacht_name, ">"))
                temp_yacht_name = LTrim(RTrim((Trim(temp_yacht_name))))
              End If

              temp_yacht_name = Replace(temp_yacht_name, "'", "''")
            End If

            extra_note = ""
            skip_this = False
            ' 4 si photo
            temp_vessel_type = array_split2(4)
            If InStr(temp_vessel_type, "No photos for this ship") > 0 Then
              extra_note = "No Picture Found"
            End If

            ' 5 is type 
            temp_vessel_type = array_split2(5)
            If InStr(temp_vessel_type, "Exhibition Ship") > 0 Or InStr(temp_vessel_type, "Supply Vessel") > 0 Then
              skip_this = True
            End If


            ' 6 is latest postition
            ' 7 is port
            ' 8 is last port
            ' 9 is area
            ' 10 is destination
            ' 11 is my fleet
            'Next

            temp_exists = False
            match_has_mmsi = False
            mmsi_string = ""

            temp_exists = CHECK_IF_Yacht_EXISTS(temp_yacht_name, temp_imo, temp_mssi, 0)

            If temp_exists = False Then
              temp_exists = CHECK_IF_Yacht_EXISTS("", temp_imo, temp_mssi, 0)
              If temp_exists = False Then
                temp_exists = CHECK_IF_Yacht_EXISTS(temp_yacht_name, "", temp_mssi, 0)
                If temp_exists = False Then
                  temp_exists = CHECK_IF_Yacht_EXISTS("", "", temp_mssi, 0)
                  If temp_exists = False Then
                    mmsi_string = "No Matches Found"
                  Else
                    mmsi_string = "Found Yacht MMSI Match"
                  End If
                Else
                  mmsi_string = "Found Yacht Name Match"
                End If
              Else
                mmsi_string = "Found Yacht IMO Match"
              End If
            Else
              mmsi_string = "Found Yacht Name/IMO Match"
            End If

            If temp_exists = True And match_has_mmsi = True Then
              temp_action = "MMSI Match - Yacht Found"
            Else
              If temp_exists = True And match_has_mmsi = False Then
                ' then our curret yacht has an MMSI number 
                If CHECK_IF_Yacht_EXISTS("", "", "", yacht_id_sy) = True Then
                  temp_action = "MMSI Not a Match - Yacht Found"
                Else
                  temp_action = "MMSI Empty - Yacht Found"
                  blank_mmsi_on_yacht = blank_mmsi_on_yacht + 1
                End If
              ElseIf temp_exists = False Then
                temp_action = "MMSI Not Found - Yacht Not Found"
                new_yacht_mmsi = new_yacht_mmsi + 1
              Else
                temp_action = "Unknown"
              End If
            End If

            If match_has_mmsi = True Then
              match_mmsi_count = match_mmsi_count + 1
              mmsi_string &= ", MMSI Number Matches"
            Else
              mmsi_string &= ", MMSI Number Not Found"
            End If

            If yacht_id_sy > 0 Then
              mmsi_string &= ", Yacht ID: " & yacht_id_sy
              yt_count = yt_count + 1
            End If

            If Trim(temp_yacht_name) <> "" Then
              mmsi_string &= ", Yacht Name: " & temp_yacht_name
            End If

            If Trim(temp_imo) <> "" Then
              mmsi_string &= ", Imo Number: " & temp_imo
            End If

            If Trim(temp_yacht_name) <> "" Then
              mmsi_string &= ", MMSI: " & temp_mssi
            End If

            If Trim(ys_dups) <> "" Then
              mmsi_string &= ", " & ys_dups
            End If

            temp_imo = temp_imo
            temp_mssi = temp_mssi
            temp_yacht_name = temp_yacht_name
            ypl_link = ""
            rows_temp = ""
            ypl_details = ""

            If Trim(temp_action) <> "MMSI Match - Yacht Found" Then
              yt_table &= "<tr>"
              yt_table &= "<td align='left'>" & temp_yacht_name & "</td>"
              yt_table &= "<td align='left'>" & yacht_id_sy & "</td>"
              yt_table &= "<td align='left'>" & temp_mssi & " (" & ys_mmsi & ")</td>"

              If Trim(temp_mssi) <> Trim(ys_mmsi) Then
                ypl_details = "MMSI: " & temp_mssi & " vs (YS: " & ys_mmsi & ") "
              End If

              If Trim(temp_imo) <> Trim(ys_imo) Then
                If Trim(temp_imo) = "-" And Trim(ys_imo) = "" Then
                Else
                  If Trim(temp_mssi) <> Trim(ys_mmsi) Then
                    ypl_details &= ","
                  End If
                  ypl_details &= "IMO: " & temp_imo & " vs (YS: " & ys_imo & ") "
                End If
              End If

              If Trim(temp_yacht_flag) <> "" Then
                ypl_details &= " Yacht Flag: " & temp_yacht_flag
              End If

              If Trim(temp_imo) <> "" And Trim(temp_imo) <> "-" Then
                yt_table &= "<td align='left'>" & temp_imo & " (" & ys_imo & ")</td>"
              ElseIf Trim(temp_imo) = "" Or Trim(temp_imo) = "-" Then
                If Trim(ys_imo) <> "" Then
                  yt_table &= "<td align='left'>" & temp_imo & " (" & ys_imo & ")</td>"
                Else
                  yt_table &= "<td align='left'>&nbsp;</td>"
                End If
              End If

              If Trim(extra_note) <> "" Then
                ypl_details &= " " & extra_note
              End If

              yt_table &= "<td align='left'>" & mmsi_string & "</td>"
              yt_table &= "<td align='left'>" & temp_action & "</td>"
              yt_table &= "</tr>"

            End If


            ' if we suggested a skip and we didnt find it in our system, then skip it 
            If skip_this = True And temp_action = "MMSI Not Found - Yacht Not Found" Then
              bad_type = bad_type + 1
            Else
              ypl_id = 0
              If Not CHECK_IF_PUB_EXISTS(temp_yacht_name, yacht_id_sy, "", 8) Then
                '  '---------------------------------------------------
                ypl_link = "http://www.marinetraffic.com/en/ais/details/ships/mmsi:" & temp_mssi & "/"
                Call insert_into_yacht_ypl(8, temp_yacht_name, temp_action, yacht_id_sy, ypl_details, ypl_link) ' details is other info

                new_ypl_insert = new_ypl_insert + 1
                '  '---------------------------------------------------
              Else
                dup_count = dup_count + 1
                ''''Call update_yacht_ypl(temp_action)
              End If
            End If

            total_mmsi = total_mmsi + 1


          Next


        End If
      End If


      yt_table = ""


    Catch ex As Exception
    Finally
      'MySqlConn_JETNET.Close()
      'MySqlConn_JETNET.Dispose()

      'MySqlCommand_JETNET.Dispose()


    End Try

  End Function

  Private Sub PrintHelpPage()

    ' Create a WebBrowser instance. 
    Dim webBrowserForPrinting As New WebBrowser()

    ' Add an event handler that prints the document after it loads.
    AddHandler webBrowserForPrinting.DocumentCompleted, New  _
        WebBrowserDocumentCompletedEventHandler(AddressOf PrintDocument)

    ' Set the Url property to load the document.
    webBrowserForPrinting.Url = New Uri("\test\help.html")

  End Sub

  Private Sub PrintDocument(ByVal sender As Object, _
      ByVal e As WebBrowserDocumentCompletedEventArgs)

    Dim webBrowserForPrinting As WebBrowser = CType(sender, WebBrowser)

    ' Print the document now that it is fully loaded.
    webBrowserForPrinting.Print()
    MessageBox.Show("print")

    ' Dispose the WebBrowser now that the task is complete. 
    webBrowserForPrinting.Dispose()

  End Sub
    Public Function scrape_for_controller(ByVal str As StreamReader, ByVal page_num As Integer)
        ' Dim Str As System.IO.Stream
        ' Dim srRead As System.IO.StreamReader



        Dim string_text As String = ""
        Dim temp1_string As String = ""
        Dim string_text2 As String = ""
        Dim i As Integer = 0
        Dim final_string As String = ""
        Dim original_string_text As String = ""
        Dim article_link As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim k As Integer = 0
        Dim skip_this As Boolean = False
        Dim extra_note As String = ""

        Dim temp_ac_name As String = ""
        Dim temp_engine As String = ""
        Dim temp_eng As String = ""
        Dim temp_av As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_make As String = ""
        Dim temp_temp As String
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split_make() As String
        Dim pass_try As String = "1"
        Dim temp_seller_info_string As String = ""
        Dim temp_string2 As String = ""


        Try



            'Dim wb As New WebBrowser
            'wb.ScrollBarsEnabled = False
            'wb.ScriptErrorsSuppressed = True
            'wb.Navigate("https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft?mdlx=Contains&Cond=All&SortOrder=35&scf=False&LS=5")
            ' While wb.ReadyState

            'End While 
            'Dim webBrowserForPrinting As New WebBrowser()
            'AddHandler webBrowserForPrinting.DocumentCompleted, New  _
            '   WebBrowserDocumentCompletedEventHandler(AddressOf PrintDocument)
            'wb.Document.DomDocument.ToString()


            'Dim webBrowserForPrinting As New WebBrowser()
            'webBrowserForPrinting.Url = New Uri("\test\help.html")


            'Dim sw As StreamWriter
            'Dim poststring = " "

            'Try
            '  sw = New StreamWriter(req.GetRequestStream)
            '  sw.Write(poststring)
            '  sw.Close()
            'Catch ex As Exception

            'End Try

            'Call PrintHelpPage()




            'Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/")
            'Dim resp As System.Net.WebResponse = req.GetResponse


            'Str = resp.GetResponseStream
            'srRead = New System.IO.StreamReader(Str)
            '' read all the text 
            'string_text = srRead.ReadToEnd().ToString

            'req.Abort()
            'resp.Close()
            'Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft/manufacturer/cessna/model/citation-latitude/")



            '
            ' Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/manufacturer/cessna?sortorder=27&SCF=False")

            ' req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft?mdlx=Contains&Cond=All&SortOrder=35&scf=False&LS=5")
            ' req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft/manufacturer/cessna/model/citation-latitude")





            ' THIS IS THE LINK FOR ALL JETS   -- DOESNT WORK --https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft/
            ' ALL CESSNAS - https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft/manufacturer/cessna/model/citation-latitude/
            'https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft/?sortorder=27&SCF=False/

            '  Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft/?sortorder=27&SCF=False%2f&page=" & page_num & "/")
            ' '' '' ''System.Threading.Thread.Sleep(10)
            ' '' '' ''Response.Flush()
            ' '' '' ''System.Threading.Thread.Sleep(10)

            ' '' '' ''Dim req As System.Net.WebRequest



            ' '' '' ''If Trim(type_string) = "last7" Then
            ' '' '' ''  req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft?sortorder=27&SCF=False&page=" & page_num & "/")
            ' '' '' ''Else
            ' '' '' ''  req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/" & type_string & "/?sortorder=27&SCF=False%2f&page=" & page_num & "")
            ' '' '' ''End If


            ' '' '' ''Dim resp As System.Net.WebResponse = req.GetResponse

            'Dim temp_url As String = ""

            'temp_url = "https://www.controller.com/listings/aircraft/for-sale/list/category/" & type_string & "/?sortorder=27&SCF=False%2f"
            'Using client As New Net.WebClient
            '  Dim reqparm As New Specialized.NameValueCollection
            '  reqparm.Add("page", "1")
            '  Dim responsebytes = client.UploadValues(temp_url, "POST", reqparm)

            'End Using

            Response.Write("<br/>1")

            '  Str = resp.GetResponseStream
            '   srRead = New System.IO.StreamReader(Str)

            '   srRead = str.ReadToEnd()


            ' read all the text 
            string_text = str.ReadToEnd()

            Response.Write("<br/>2")
            '   resp.Close()
            '    resp = Nothing
            '   req = Nothing

            string_text = string_text
            original_string_text = string_text

            ' GET HOW MANY PAGES THERE ARE --------------
            If page_num = 1 Then
                spot_to_find = InStr(string_text, "listings-total-pages")
                If spot_to_find > 0 Then
                    string_text = Right(string_text, Len(string_text) - spot_to_find - 22)
                    spot_to_find = InStr(string_text, "</span>")
                    If spot_to_find > 0 Then
                        string_text = Left(Trim(string_text), spot_to_find - 1)
                        If IsNumeric(Trim(string_text)) = True Then
                            total_pages = CInt(Trim(string_text))
                        Else
                            total_pages = 10
                        End If
                    Else
                        total_pages = 10
                    End If
                Else
                    total_pages = 10
                End If
                string_text = original_string_text
            End If


            If InStr(string_text, "General Listings") > 0 And acpub_controller_general_start = 0 Then
                acpub_controller_general_start = (page_num + 5)
            End If



            '  spot_to_find = InStr(string_text, "listing-top-left")
            spot_to_find = InStr(string_text, "list-listing-title""") ' changed - 9/20 
            If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 2)

                ' array_split = Split(string_text, "listing-top-left")  ' changed - 9/20 
                array_split = Split(string_text, "list-listing-title""")

                For i = 0 To array_split.Length - 1
                    string_text = array_split(i)
                    original_string_text = string_text
                    acpub_count = acpub_count + 1

                    temp_ac_name = ""
                    temp_engine = ""
                    temp_eng = ""
                    temp_av = ""

                    pub_reg_no = ""
                    pub_ser_no = ""
                    pub_desc = ""
                    pub_price = ""
                    pub_aftt = ""
                    pub_seller_info = ""
                    pub_picture = ""
                    pub_status = ""
                    pub_landings = ""
                    pub_url = ""
                    has_pics = False
                    aftt_different = ""
                    landings_different = ""
                    pub_comp_id = 0
                    pub_seller_info_no_city = ""



                    If InStr(string_text, "FL-58") > 0 Then
                        string_text = string_text
                    End If

                    If InStr(string_text, "1998 BEECHCRAFT 58 BARON") > 0 Then
                        string_text = string_text
                    End If





                    'If InStr(string_text, "images/pictures.png") > 0 Then
                    '    has_pics = True
                    'Else
                    '    has_pics = False
                    'End If

                    ' changed - msw 
                    If InStr(string_text, "no-image-icon.svg") > 0 Then
                        has_pics = False
                    Else
                        has_pics = True
                    End If



                    spot_to_find = InStr(string_text, "listing-portion-title"">")
                    If spot_to_find > 0 Then
                        temp_ac_name = Right(string_text, Len(string_text) - spot_to_find - 22)

                        spot_to_find = InStr(temp_ac_name, "</h3>")
                        If spot_to_find > 0 Then
                            temp_ac_name = Left(Trim(temp_ac_name), spot_to_find - 1)
                        End If
                    End If

                    spot_to_find = InStr(temp_ac_name, "2012 ROBINSON R66")
                    If spot_to_find > 0 Then
                        temp_ac_name = temp_ac_name
                    End If


                    ' added MSW - 6/28/21 
                    spot_to_find = InStr(temp_ac_name, "</h2>")
                    If spot_to_find > 0 Then
                        temp_ac_name = Left(Trim(temp_ac_name), spot_to_find - 1)
                    End If


                    spot_to_find = InStr(string_text, "a href=")
                    If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 8)

                        spot_to_find = InStr(string_text, ">")
                        If spot_to_find > 0 Then
                            pub_url = Left(Trim(string_text), spot_to_find - 2)
                            pub_url = Replace(pub_url, """", "")

                            pub_url = "https://www.controller.com/" & pub_url
                            pub_url = RTrim(LTrim(pub_url))

                            string_text = Right(string_text, Len(string_text) - spot_to_find)

                            If Trim(temp_ac_name) <> "" Then
                            Else
                                spot_to_find = InStr(string_text, "<")
                                If spot_to_find > 0 Then
                                    temp_ac_name = Left(Trim(string_text), spot_to_find - 1)
                                End If
                            End If



                            spot_to_find = InStr(string_text, "listing-description-text")
                            If spot_to_find > 0 Then
                                string_text = Right(string_text, Len(string_text) - spot_to_find - 24)

                                spot_to_find = InStr(string_text, "</div>")
                                spot_to_find2 = InStr(Left(string_text, 20), "employee-category")
                                If spot_to_find > 0 And spot_to_find2 > 0 Then
                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 6)
                                End If

                                spot_to_find = InStr(string_text, "</div>")
                                If spot_to_find > 0 Then
                                    pub_desc = Left(Trim(string_text), spot_to_find - 1)

                                    spot_to_find = InStr(pub_desc, "<a class")
                                    If spot_to_find > 0 Then
                                        pub_desc = Left(Trim(pub_desc), spot_to_find - 1)
                                    End If
                                    If Trim(pub_desc) = "" Then
                                        pub_desc = pub_desc
                                    End If
                                    pub_desc = Replace(pub_desc, "’", "")
                                    pub_desc = Replace(pub_desc, "'", "")
                                    cutme(pub_desc)
                                    pub_desc = Left(Trim(pub_desc), 500)
                                    pub_desc = ""
                                End If

                            End If

                            If InStr(pub_desc, "Price Reduced,") > 0 Then
                                pub_url = pub_url
                            End If


                            If InStr(pub_url, "33267775") > 0 Then
                                pub_url = pub_url
                            End If

                            pass_try = "1"


                            If InStr(string_text, "N353KM") > 0 Then
                                pub_url = pub_url
                            End If



                            string_text = Replace(string_text, "Serial Number<!-- -->:</span>", "Serial Number</span>")
                            string_text = Replace(string_text, "Registration #<!-- -->:</span>", "Registration #</span>")
                            string_text = Replace(string_text, "Registration #:</span>", "Registration #</span>")
                            string_text = Replace(string_text, "Total Time<!-- -->: </span>", "Total Time</span>")
                            string_text = Replace(string_text, "Serial Number</div>", "Serial Number</span>")





                            spot_to_find = InStr(Trim(string_text), "Serial #</span>")
                            If spot_to_find > 0 Then
                                pass_try = "1"
                            ElseIf InStr(Trim(string_text), "Serial #:</span>") > 0 Then
                                pass_try = "2"
                            ElseIf InStr(Trim(string_text), "Serial Number:</span>") > 0 Then
                                pass_try = "3"
                            ElseIf InStr(Trim(string_text), "Serial Number</span>") > 0 Then
                                pass_try = "4"
                            Else
                                ' if we dont find anything, then try again withthe full string 
                                string_text = original_string_text   ' reset it -- most likely the serial number is before the description

                                string_text = Replace(string_text, "Serial Number<!-- -->:</span>", "Serial Number</span>")
                                string_text = Replace(string_text, "Registration #<!-- -->:</span>", "Registration #</span>")
                                string_text = Replace(string_text, "Registration #:</span>", "Registration #</span>")
                                string_text = Replace(string_text, "Total Time<!-- -->: </span>", "Total Time</span>")
                                string_text = Replace(string_text, "Serial Number</div>", "Serial Number</span>")
                            End If







                            spot_to_find = InStr(Trim(string_text), "Serial #</span>")
                            If spot_to_find > 0 Then
                                pass_try = "1"
                            ElseIf InStr(Trim(string_text), "Serial #:</span>") > 0 Then
                                pass_try = "2"
                            ElseIf InStr(Trim(string_text), "Serial Number:</span>") > 0 Then
                                pass_try = "3"
                            ElseIf InStr(Trim(string_text), "Serial Number</span>") > 0 Then
                                pass_try = "4"
                            Else

                                spot_to_find = InStr(Trim(string_text), "Registration #</span>")
                                If spot_to_find = 0 Then
                                    spot_to_find = InStr(Trim(string_text), "Registration #:</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Replace(string_text, "Registration #:</span>", "Registration #</span>")
                                        pass_try = "4"   ' changed from = "2" - passing all thro 4 - no need to replace 
                                    Else
                                        ' if it doesnt have serial or reg, look for total time 
                                        spot_to_find = InStr(Trim(string_text), "Total Time</span>")
                                        If spot_to_find = 0 Then
                                            spot_to_find = InStr(Trim(string_text), "Total Time:</span>")
                                            If spot_to_find > 0 Then
                                                pass_try = "2"
                                            Else
                                                pass_try = "4"
                                            End If
                                        Else
                                            pass_try = "4"
                                        End If

                                    End If
                                Else
                                    pass_try = "4"
                                End If



                            End If





                            If pass_try = "1" Then   ' THEN WE ARE ASSUMING IT IS ON THE TYPE OF PAGE WITH NO : , this type has for sale after - PAGE 1 -----------
                                '---------------------------------------------------------------------------------------------------------------------------

                                spot_to_find = InStr(Trim(string_text), "Serial #</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 14)

                                    spot_to_find = InStr(Trim(string_text), "</div>")
                                    If spot_to_find > 0 Then
                                        pub_ser_no = Replace(pub_ser_no, "n>", "")
                                        pub_ser_no = Replace(pub_ser_no, ">", "")
                                        pub_ser_no = Left(Trim(string_text), spot_to_find - 1)
                                        pub_ser_no = LTrim(RTrim(pub_ser_no))
                                        cutme(pub_ser_no)
                                    End If
                                End If

                                spot_to_find = InStr(Trim(string_text), "Registration #</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 22)

                                    spot_to_find = InStr(Trim(string_text), "</div>")
                                    If spot_to_find > 0 Then
                                        pub_reg_no = Left(Trim(Trim(string_text)), spot_to_find - 1)
                                        pub_reg_no = LTrim(RTrim(pub_reg_no))
                                        cutme_LF(pub_reg_no)
                                        pub_reg_no = LTrim(RTrim(pub_reg_no))
                                    End If
                                End If

                                spot_to_find = InStr(Trim(string_text), "Total Time</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 18)

                                    spot_to_find = InStr(Trim(string_text), "</div>")
                                    If spot_to_find > 0 Then
                                        pub_aftt = Left(Trim(string_text), spot_to_find - 1)

                                        pub_aftt = Replace(pub_aftt, " ", "")
                                        pub_aftt = LTrim(pub_aftt)
                                        pub_aftt = RTrim(pub_aftt)
                                    End If
                                End If

                                Try
                                    spot_to_find = InStr(string_text, "For Sale Price:</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 21)

                                        spot_to_find = InStr(string_text, "</div>")
                                        If spot_to_find > 0 Then
                                            pub_price = Left(Trim(string_text), spot_to_find - 2)
                                            pub_price = Replace(pub_price, "<span class=""nobr"">", "")
                                            pub_price = Replace(pub_price, "</span>", "")
                                            pub_price = Replace(pub_price, "USD $", "")
                                            spot_to_find = InStr(pub_price, "<div")
                                            If spot_to_find > 0 Then
                                                pub_price = Left(Trim(pub_price), spot_to_find - 1)
                                            End If
                                            cutme(pub_price)
                                        End If
                                    End If

                                Catch ex As Exception
                                    Response.Write(ex.ToString)
                                End Try

                            ElseIf pass_try = "2" Then  '--------- THIS WILL BE THE SETUP OF THE PAGE 2's - THE ONES WITHOUT THE : and where for sale is first 
                                '---------------------------------------------------------------------------------------------------------------------------
                                '---------------------------------------------------------------------------------------------------------------------------


                                Try
                                    spot_to_find = InStr(string_text, "For Sale Price:</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 21)

                                        spot_to_find = InStr(string_text, "</div>")
                                        If spot_to_find > 0 Then
                                            pub_price = Left(Trim(string_text), spot_to_find - 2)
                                            pub_price = Replace(pub_price, "<span class=""nobr"">", "")
                                            pub_price = Replace(pub_price, "</span>", "")
                                            pub_price = Replace(pub_price, "USD $", "")
                                            spot_to_find = InStr(pub_price, "<div")
                                            If spot_to_find > 0 Then
                                                pub_price = Left(Trim(pub_price), spot_to_find - 1)
                                            End If
                                            cutme(pub_price)
                                        End If
                                    End If

                                Catch ex As Exception
                                    Response.Write(ex.ToString)
                                End Try


                                spot_to_find = InStr(Trim(string_text), "Serial #:</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 15)   ' changed from 16 to 15 - MSW - 3/3/20 

                                    spot_to_find = InStr(Trim(string_text), "</div>")
                                    If spot_to_find > 0 Then
                                        pub_ser_no = Left(Trim(string_text), spot_to_find - 1)

                                        If InStr(string_text, ">") > 0 Then
                                            string_text = Replace(string_text, ">", "")
                                        End If

                                        pub_ser_no = LTrim(RTrim(pub_ser_no))
                                        cutme(pub_ser_no)
                                    End If
                                End If

                                spot_to_find = InStr(Trim(string_text), "Registration #:</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 22)

                                    spot_to_find = InStr(Trim(string_text), "</div>")
                                    If spot_to_find > 0 Then
                                        pub_reg_no = Left(Trim(Trim(string_text)), spot_to_find - 1)
                                        pub_reg_no = LTrim(RTrim(pub_reg_no))
                                        cutme_LF(pub_reg_no)
                                        pub_reg_no = LTrim(RTrim(pub_reg_no))
                                    End If
                                End If

                                spot_to_find = InStr(Trim(string_text), "Total Time:</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 18)

                                    spot_to_find = InStr(Trim(string_text), "</div>")
                                    If spot_to_find > 0 Then
                                        pub_aftt = Left(Trim(string_text), spot_to_find - 1)

                                        pub_aftt = Replace(pub_aftt, " ", "")
                                        pub_aftt = LTrim(pub_aftt)
                                        pub_aftt = RTrim(pub_aftt)
                                    End If
                                End If
                            ElseIf pass_try = "3" Then  '--------- THIS WILL BE THE SETUP OF THE PAGE 2's - THE ONES WITHOUT THE : and where for sale is first 
                                '---------------------------------------------------------------------------------------------------------------------------
                                '---------------------------------------------------------------------------------------------------------------------------

                                Try
                                    spot_to_find = InStr(string_text, "For Sale Price:</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 21)

                                        spot_to_find = InStr(string_text, "</div>")
                                        If spot_to_find > 0 Then
                                            pub_price = Left(Trim(string_text), spot_to_find - 2)
                                            pub_price = Replace(pub_price, "<span class=""nobr"">", "")
                                            pub_price = Replace(pub_price, "</span>", "")
                                            pub_price = Replace(pub_price, "USD $", "")
                                            spot_to_find = InStr(pub_price, "<div")
                                            If spot_to_find > 0 Then
                                                pub_price = Left(Trim(pub_price), spot_to_find - 1)
                                            End If
                                            cutme(pub_price)
                                        End If
                                    End If

                                Catch ex As Exception
                                    Response.Write(ex.ToString)
                                End Try

                                spot_to_find = InStr(Trim(string_text), "Serial Number:</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 20)

                                    spot_to_find = InStr(Trim(string_text), "</div>")
                                    If spot_to_find > 0 Then
                                        pub_ser_no = Left(Trim(string_text), spot_to_find - 1)


                                        If InStr(string_text, ">") > 0 Then
                                            string_text = Replace(string_text, ">", "")
                                        End If
                                        pub_ser_no = LTrim(RTrim(pub_ser_no))
                                        cutme(pub_ser_no)
                                    End If
                                End If

                                spot_to_find = InStr(Trim(string_text), "Registration #:</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 22)

                                    spot_to_find = InStr(Trim(string_text), "</div>")
                                    If spot_to_find > 0 Then
                                        pub_reg_no = Left(Trim(Trim(string_text)), spot_to_find - 1)
                                        pub_reg_no = LTrim(RTrim(pub_reg_no))
                                        cutme_LF(pub_reg_no)
                                        pub_reg_no = LTrim(RTrim(pub_reg_no))
                                    End If
                                End If

                                spot_to_find = InStr(Trim(string_text), "Total Time:</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 18)

                                    spot_to_find = InStr(Trim(string_text), "</div>")
                                    If spot_to_find > 0 Then
                                        pub_aftt = Left(Trim(string_text), spot_to_find - 1)

                                        pub_aftt = Replace(pub_aftt, " ", "")
                                        pub_aftt = LTrim(pub_aftt)
                                        pub_aftt = RTrim(pub_aftt)
                                    End If
                                End If

                            ElseIf pass_try = "4" Then  '--------- THIS WILL BE THE SETUP OF THE PAGE 2's - THE ONES WITHOUT THE : and where for sale is first 
                                '---------------------------------------------------------------------------------------------------------------------------
                                '---------------------------------------------------------------------------------------------------------------------------



                                spot_to_find = InStr(Trim(string_text), "Serial Number</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 20)

                                    spot_to_find = InStr(Trim(string_text), "</span>")

                                    ' added in MSW - 8/24/2020 
                                    If spot_to_find = 1 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 7)
                                    End If


                                    spot_to_find = InStr(Trim(string_text), "</span>")
                                    If spot_to_find > 0 Then
                                        pub_ser_no = Left(Trim(string_text), spot_to_find - 1)
                                        pub_ser_no = LTrim(RTrim(pub_ser_no))

                                        spot_to_find = InStr(Trim(pub_ser_no), ">")
                                        If spot_to_find > 0 Then
                                            pub_ser_no = Right(Trim(pub_ser_no), Len(Trim(pub_ser_no)) - spot_to_find)
                                        End If
                                        pub_ser_no = Replace(pub_ser_no, "</span>", "")
                                        cutme(pub_ser_no)
                                    End If
                                End If

                                spot_to_find = InStr(Trim(string_text), "Registration #</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 21)

                                    ' added in MSW - 8/24/2020 
                                    If spot_to_find = 1 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 7)
                                    End If

                                    spot_to_find = InStr(Trim(string_text), "</span>")
                                    If spot_to_find > 0 Then
                                        pub_reg_no = Left(Trim(Trim(string_text)), spot_to_find - 1)
                                        pub_reg_no = LTrim(RTrim(pub_reg_no))



                                        spot_to_find = InStr(Trim(pub_reg_no), ">")
                                        If spot_to_find > 0 Then
                                            pub_reg_no = Right(Trim(pub_reg_no), Len(Trim(pub_reg_no)) - spot_to_find)
                                        End If

                                        cutme_LF(pub_reg_no)
                                        pub_reg_no = LTrim(RTrim(pub_reg_no))
                                    End If
                                End If

                                spot_to_find = InStr(Trim(string_text), "Total Time</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 18)

                                    spot_to_find = InStr(Trim(string_text), "</div>")
                                    If spot_to_find > 0 Then
                                        pub_aftt = Left(Trim(string_text), spot_to_find - 1)

                                        pub_aftt = Replace(pub_aftt, "</span>", "")

                                        spot_to_find = InStr(Trim(pub_aftt), ">")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Right(Trim(pub_aftt), Len(Trim(pub_aftt)) - spot_to_find)
                                        End If

                                        '' added in 2 more in - MSW - aftt 
                                        spot_to_find = InStr(Trim(pub_aftt), ">")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Right(Trim(pub_aftt), Len(Trim(pub_aftt)) - spot_to_find)
                                        End If

                                        spot_to_find = InStr(Trim(pub_aftt), ">")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Right(Trim(pub_aftt), Len(Trim(pub_aftt)) - spot_to_find)
                                        End If

                                        spot_to_find = InStr(Trim(pub_aftt), ">")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Right(Trim(pub_aftt), Len(Trim(pub_aftt)) - spot_to_find)
                                        End If

                                        spot_to_find = InStr(Trim(pub_aftt), ">")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Right(Trim(pub_aftt), Len(Trim(pub_aftt)) - spot_to_find)
                                        End If

                                        spot_to_find = InStr(Trim(pub_aftt), ">")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Right(Trim(pub_aftt), Len(Trim(pub_aftt)) - spot_to_find)
                                        End If


                                        pub_aftt = Replace(pub_aftt, " ", "")
                                        pub_aftt = LTrim(pub_aftt)
                                        pub_aftt = RTrim(pub_aftt)
                                    End If
                                End If

                                If Trim(pub_aftt) = "" Then
                                    spot_to_find = InStr(Trim(string_text), "Total Time Since New: ")
                                    If spot_to_find > 0 Then
                                        pub_aftt = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 21)

                                        spot_to_find = InStr(Trim(pub_aftt), " Hrs")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Left(Trim(pub_aftt), spot_to_find - 1)
                                        End If

                                        spot_to_find = InStr(Trim(LCase(pub_aftt)), " hours")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Left(Trim(pub_aftt), spot_to_find - 1)
                                        End If


                                        spot_to_find = InStr(Trim(pub_aftt), ">")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Right(Trim(pub_aftt), Len(Trim(pub_aftt)) - spot_to_find)
                                        End If

                                        If Len(Trim(pub_aftt)) > 15 Then
                                            pub_aftt = ""
                                        End If
                                    End If
                                End If




                                spot_to_find = InStr(string_text, "Total Landings</span>", CompareMethod.Text)
                                If spot_to_find > 0 Then
                                    spot_to_find = InStr(string_text, "</span>", CompareMethod.Text)
                                    pub_landings = Left(string_text, spot_to_find - 1)

                                    spot_to_find = InStr(Trim(pub_landings), ">")
                                    If spot_to_find > 0 Then
                                        pub_landings = Right(Trim(pub_landings), Len(Trim(pub_landings)) - spot_to_find)
                                    End If



                                    cutme(pub_landings)
                                    pub_landings = Replace(pub_landings, ",", "")
                                    pub_landings = RTrim(LTrim(pub_landings))
                                ElseIf InStr(string_text, "Total Landings", CompareMethod.Text) > 0 Then

                                    spot_to_find = InStr(string_text, "Total Landings", CompareMethod.Text)
                                    If spot_to_find > 0 Then
                                        spot_to_find = InStr(string_text, "</span>", CompareMethod.Text)
                                        pub_landings = Left(string_text, spot_to_find - 1)

                                        spot_to_find = InStr(Trim(pub_landings), ">")
                                        If spot_to_find > 0 Then
                                            pub_landings = Right(Trim(pub_landings), Len(Trim(pub_landings)) - spot_to_find)
                                        End If


                                        spot_to_find = InStr(Trim(pub_landings), ">")
                                        If spot_to_find > 0 Then
                                            pub_landings = Right(Trim(pub_landings), Len(Trim(pub_landings)) - spot_to_find)
                                        End If


                                        cutme(pub_landings)
                                        pub_landings = Replace(pub_landings, ",", "")
                                        pub_landings = RTrim(LTrim(pub_landings))
                                    End If



                                End If


                                string_text = Replace(string_text, "", "Registration #</span>")

                                spot_to_find = InStr(string_text, "isting-dealer-info")
                                If spot_to_find > 0 Then
                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 18)

                                    spot_to_find = InStr(string_text, ">")
                                    If spot_to_find > 0 Then
                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 1)


                                        spot_to_find = InStr(string_text, "</h5>") ' get whatever is after the name, could be span or href
                                        If spot_to_find > 0 Then
                                            pub_seller_info = Left(Trim(string_text), spot_to_find - 1)
                                        End If

                                        spot_to_find = InStr(pub_seller_info, ">")
                                        If spot_to_find > 0 Then
                                            pub_seller_info = Right(pub_seller_info, Len(pub_seller_info) - spot_to_find)
                                        End If



                                        spot_to_find = InStr(Trim(pub_seller_info), ">")
                                        If spot_to_find > 0 Then
                                            pub_seller_info = Right(Trim(pub_seller_info), Len(Trim(pub_seller_info)) - spot_to_find)
                                        End If

                                    End If



                                End If

                                Try
                                    spot_to_find = InStr(string_text, "retail-price-container")
                                    If spot_to_find > 0 Then
                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 22)

                                        spot_to_find = InStr(string_text, "</div>")
                                        If spot_to_find > 0 Then
                                            pub_price = Left(Trim(string_text), spot_to_find - 2)
                                            pub_price = Replace(pub_price, "<span class=""price"">", "")
                                            pub_price = Replace(pub_price, "</span>", "")
                                            pub_price = Replace(pub_price, "</span", "")
                                            pub_price = Replace(pub_price, "USD $", "")
                                            pub_price = Replace(pub_price, "><span class=""spec-label"">", "")
                                            pub_price = Replace(pub_price, "Price:", "")
                                            pub_price = Replace(pub_price, "USD $", "")
                                            spot_to_find = InStr(pub_price, "<div")
                                            If spot_to_find > 0 Then
                                                pub_price = Left(Trim(pub_price), spot_to_find - 1)
                                            End If
                                            cutme(pub_price)
                                        End If
                                    Else
                                        ' MSW - go look at original , see if its in there 
                                        spot_to_find = InStr(original_string_text, "retail-price-container")

                                        If spot_to_find > 0 Then
                                            temp_string = original_string_text
                                            temp_string = Right(temp_string, Len(temp_string) - spot_to_find - 22)

                                            spot_to_find = InStr(temp_string, "</div>")
                                            If spot_to_find > 0 Then
                                                pub_price = Left(Trim(temp_string), spot_to_find - 2)
                                                pub_price = Replace(pub_price, "<span class=""price"">", "")
                                                pub_price = Replace(pub_price, "</span>", "")
                                                pub_price = Replace(pub_price, "</span", "")
                                                pub_price = Replace(pub_price, "USD $", "")
                                                pub_price = Replace(pub_price, "><span class=""spec-label"">", "")
                                                pub_price = Replace(pub_price, "Price:", "")
                                                pub_price = Replace(pub_price, "USD $", "")
                                                spot_to_find = InStr(pub_price, "<div")
                                                If spot_to_find > 0 Then
                                                    pub_price = Left(Trim(pub_price), spot_to_find - 1)
                                                End If
                                                cutme(pub_price)
                                            End If
                                        End If

                                    End If

                                Catch ex As Exception
                                    Response.Write(ex.ToString)
                                End Try

                            Else ' MAYBE THERE IS NO SERIAL NUMBER - so check 
                                '---------------------------------------------------------------------------------------------------------------------------


                            End If



                            If InStr(acpub_original_name, "1991 BELL 212") > 0 Then
                                acpub_original_name = acpub_original_name
                            End If

                            If Trim(pub_reg_no) <> "" Then
                                If InStr(pub_reg_no, Chr(13) & Chr(10)) > 0 Then
                                    pub_reg_no = Left(Trim(pub_reg_no), InStr(pub_reg_no, Chr(13) & Chr(10)) - 1)
                                End If

                                If InStr(pub_reg_no, Chr(10)) > 0 Then
                                    pub_reg_no = Left(Trim(pub_reg_no), InStr(pub_reg_no, Chr(10)) - 1)
                                End If
                            End If


                            If InStr(pub_url, " target=") > 0 Then
                                pub_url = Replace(pub_url, " target=", "")
                            End If



                            spot_to_find = InStr(string_text, "Engine(s):</span>")
                            If spot_to_find > 0 Then
                                string_text = Right(string_text, Len(string_text) - spot_to_find - 17)

                                spot_to_find = InStr(string_text, "</div>")
                                If spot_to_find > 0 Then
                                    temp_eng = Left(Trim(string_text), spot_to_find - 1)
                                End If
                            End If


                            If Trim(pub_aftt) = "" Then
                                spot_to_find = InStr(string_text, "Airframe:</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 16)

                                    spot_to_find = InStr(string_text, "</div>")
                                    If spot_to_find > 0 Then
                                        pub_aftt = Left(Trim(string_text), spot_to_find - 1)
                                        spot_to_find = InStr(pub_aftt, ":")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Right(pub_aftt, Len(pub_aftt) - spot_to_find)
                                            cutme(pub_aftt)
                                            spot_to_find = InStr(LCase(pub_aftt), "hours")
                                            If spot_to_find > 0 Then
                                                pub_aftt = Left(Trim(pub_aftt), spot_to_find + 4)
                                            Else
                                                spot_to_find = InStr(pub_aftt, "Total")
                                                If spot_to_find > 0 Then
                                                    pub_aftt = Left(Trim(pub_aftt), spot_to_find - 2)
                                                ElseIf InStr(pub_aftt, "Landings") > 0 Then
                                                    pub_aftt = Left(Trim(pub_aftt), InStr(pub_aftt, "Landings") - 2)
                                                Else
                                                    spot_to_find = spot_to_find
                                                End If
                                            End If
                                        End If
                                    End If

                                    pub_aftt = Replace(pub_aftt, "Airframe", "")
                                    pub_aftt = Replace(pub_aftt, "Total", "")
                                    pub_aftt = Replace(pub_aftt, "Time", "")

                                    If IsNumeric(Left(Trim(pub_aftt), 6)) = True Then
                                        pub_aftt = Left(Trim(pub_aftt), 6)
                                    ElseIf IsNumeric(Left(Trim(pub_aftt), 5)) = True Then
                                        pub_aftt = Left(Trim(pub_aftt), 5)
                                    ElseIf IsNumeric(Left(Trim(pub_aftt), 4)) = True Then
                                        pub_aftt = Left(Trim(pub_aftt), 4)
                                    ElseIf IsNumeric(Left(Trim(pub_aftt), 3)) = True Then
                                        pub_aftt = Left(Trim(pub_aftt), 3)
                                    Else
                                        pub_aftt = ""
                                    End If
                                End If
                            End If


                            '--------- ADDED MSW - 9/20/21 ------------------------
                            '----------------------------------------------------------------------------------------------

                            If IsNothing(pub_price) = True Or Trim(pub_price) = "" Then
                                spot_to_find = InStr(original_string_text, "listing-image-price")
                                If spot_to_find > 0 Then
                                    pub_price = Right(original_string_text, Len(original_string_text) - spot_to_find - 20)

                                    spot_to_find = InStr(pub_price, "</div>")
                                    If spot_to_find > 0 Then
                                        pub_price = Left(Trim(pub_price), spot_to_find - 1)
                                    End If

                                    pub_price = Replace(pub_price, "USD ", "")

                                End If
                            End If


                            temp1_string = ""
                            ' added MSW - for new AFTT process
                            If Trim(pub_aftt) = "" Then
                                temp1_string = string_text
                                spot_to_find = InStr(string_text, "Total Time<!-- -->:</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 26)


                                    spot_to_find = InStr(string_text, "</span>")
                                    If spot_to_find > 0 Then
                                        pub_aftt = Left(Trim(string_text), spot_to_find - 1)

                                        spot_to_find = InStr(pub_aftt, ">")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Right(pub_aftt, Len(pub_aftt) - spot_to_find)
                                        End If
                                    End If

                                End If
                            End If




                            If Trim(pub_landings) = "" Then
                                If Trim(temp1_string) <> "" Then
                                    string_text = temp1_string
                                End If
                                spot_to_find = InStr(string_text, "Total Landings<!-- -->:</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 30)


                                    spot_to_find = InStr(string_text, "</span>")
                                    If spot_to_find > 0 Then
                                        pub_landings = Left(Trim(string_text), spot_to_find - 1)

                                        spot_to_find = InStr(pub_landings, ">")
                                        If spot_to_find > 0 Then
                                            pub_landings = Right(pub_landings, Len(pub_landings) - spot_to_find - 1)
                                        End If
                                    End If

                                End If
                            End If
                            '----------------------------------------------------------------------------------------------


                            spot_to_find = InStr(string_text, "Avionics/Radios:</span>")
                            If spot_to_find > 0 Then
                                string_text = Right(string_text, Len(string_text) - spot_to_find - 24)

                                spot_to_find = InStr(string_text, "</div>")
                                If spot_to_find > 0 Then
                                    temp_av = Left(Trim(string_text), spot_to_find - 1)
                                End If
                            End If

                            If Trim(pub_seller_info) = "" Then
                                spot_to_find = InStr(original_string_text, "col dealer-info")
                                If spot_to_find > 0 Then
                                    string_text = Right(original_string_text, Len(original_string_text) - spot_to_find - 20)

                                    spot_to_find = InStr(string_text, "div Class=""bold"">")
                                    If spot_to_find > 0 Then
                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 19)

                                        temp_seller_info_string = string_text

                                        spot_to_find = InStr(string_text, ">")
                                        If spot_to_find > 0 Then
                                            string_text = Right(string_text, Len(string_text) - spot_to_find)

                                            spot_to_find = InStr(string_text, "</") ' get whatever is after the name, could be span or href
                                            If spot_to_find > 0 Then
                                                pub_seller_info = Left(Trim(string_text), spot_to_find - 1)
                                            End If

                                            If Len(Trim(pub_seller_info)) > 40 Then

                                                spot_to_find = InStr(string_text, "<div>")
                                                If spot_to_find > 0 Then
                                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 4)

                                                    spot_to_find = InStr(string_text, "</div>")
                                                    If spot_to_find > 0 Then
                                                        pub_seller_info &= " " & Left(Trim(string_text), spot_to_find - 1)
                                                    End If
                                                End If
                                            Else
                                                pub_seller_info = pub_seller_info
                                            End If

                                            spot_to_find = InStr(temp_seller_info_string, pub_seller_info)
                                            If spot_to_find > 0 Then
                                                string_text = Right(temp_seller_info_string, Len(temp_seller_info_string) - spot_to_find - 5)

                                                spot_to_find = InStr(string_text, "</div>")
                                                If spot_to_find > 0 Then
                                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 6)

                                                    spot_to_find = InStr(string_text, "</div>")
                                                    If spot_to_find > 0 Then
                                                        string_text = Left(Trim(string_text), spot_to_find - 1)
                                                        string_text = Replace(string_text, "<div>", "")
                                                        string_text = Replace(string_text, "div>", "")
                                                        string_text = Replace(string_text, "", "")
                                                        pub_seller_info &= ", " & string_text
                                                    End If


                                                End If

                                            End If

                                            spot_to_find = InStr(temp_seller_info_string, "Phone:</span>")
                                            If spot_to_find > 0 Then
                                                string_text = Right(temp_seller_info_string, Len(temp_seller_info_string) - spot_to_find - 14)


                                                spot_to_find = InStr(Trim(string_text), "phonetype")
                                                If spot_to_find > 0 Then
                                                    string_text = " " & Left(Trim(string_text), spot_to_find - 2)

                                                    spot_to_find = InStr(string_text, "href=")
                                                    If spot_to_find > 0 Then
                                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 6)
                                                        string_text = Replace(string_text, "tel:+", "")
                                                        string_text = Replace(string_text, "el:+", "")
                                                        string_text = Replace(string_text, "el:", "")
                                                        string_text = Replace(string_text, """", "")
                                                        pub_seller_info &= ", " & string_text
                                                    End If


                                                End If


                                            End If


                                        End If
                                    End If
                                End If
                            Else
                                ' WE HAVE SELLER INFO 
                                pub_seller_info = pub_seller_info
                            End If




                        End If

                    End If








                    If InStr(pub_seller_info, "class=""dealer-name"">") > 0 Then
                        spot_to_find = InStr(Trim(pub_seller_info), "class=""dealer-name"">")
                        If spot_to_find > 0 Then
                            pub_seller_info = Right(Trim(pub_seller_info), Len(Trim(pub_seller_info)) - spot_to_find - 19)
                        End If
                        'class="dealer-name">
                    End If


                    ' NOT SURE IF THE ABOVE SECTION DOES ANYTHING---- 9/21/2020
                    If Trim(pub_seller_info) = "" Then
                        temp_string2 = original_string_text

                        If InStr(temp_string2, "class=""dealer-name"">") > 0 Then

                            spot_to_find = InStr(Trim(temp_string2), "class=""dealer-name"">")
                            If spot_to_find > 0 Then
                                pub_seller_info = Right(Trim(temp_string2), Len(Trim(temp_string2)) - spot_to_find - 19)
                            End If

                            spot_to_find = InStr(Trim(pub_seller_info), "</h5>")
                            If spot_to_find > 0 Then
                                pub_seller_info = Left(Trim(pub_seller_info), spot_to_find - 1)
                            End If

                            pub_seller_info = Replace(pub_seller_info, ">", "")
                        End If
                    End If




                    If Trim(pub_ser_no) = "FL-123" Or Trim(pub_ser_no) = "FL-22" Or Trim(pub_ser_no) = "BE-21" Then
                        pub_ser_no = pub_ser_no
                    End If


                    If Trim(pub_ser_no) = "" And Trim(pub_reg_no) = "" Then
                        pub_ser_no = pub_ser_no
                    End If

                    If Trim(pub_reg_no) <> "" Then
                        If InStr(Trim(pub_reg_no), "  ") > 0 Then
                            pub_reg_no = " " & Left(Trim(pub_reg_no), InStr(Trim(pub_reg_no), "  ") - 1)
                        End If
                    End If

                    pub_seller_info = Replace(pub_seller_info, "'", "''")
                    pub_seller_info = replace_fm(pub_seller_info)


                    ' ADDED MSW - 9/20/21 ----
                    If Trim(pub_seller_info) = "" Then
                        spot_to_find = InStr(temp_string2, "<span>Seller</span>:</span>")
                        If spot_to_find > 0 Then   '<span class="spec-value">Mesinger Jet Sales</span> 
                            pub_comp_id = 0
                            temp_string2 = Right(Trim(temp_string2), Len(Trim(temp_string2)) - spot_to_find - 28)

                            spot_to_find = InStr(Trim(temp_string2), "</span>")
                            If spot_to_find > 0 Then
                                temp_string2 = Left(Trim(temp_string2), spot_to_find - 1)

                                spot_to_find = InStr(Trim(temp_string2), ">")
                                If spot_to_find > 0 Then
                                    pub_seller_info = Right(Trim(temp_string2), Len(Trim(temp_string2)) - spot_to_find)

                                    pub_seller_info_no_city = pub_seller_info


                                    temp_string2 = original_string_text

                                    spot_to_find = InStr(temp_string2, "Location:</span>")
                                    If spot_to_find > 0 Then
                                        temp_string2 = Right(Trim(temp_string2), Len(Trim(temp_string2)) - spot_to_find - 17)

                                        spot_to_find = InStr(Trim(temp_string2), ">")
                                        If spot_to_find > 0 Then
                                            temp_string2 = Right(Trim(temp_string2), Len(Trim(temp_string2)) - spot_to_find)

                                            spot_to_find = InStr(Trim(temp_string2), "</span>")
                                            If spot_to_find > 0 Then
                                                temp_string2 = Left(Trim(temp_string2), spot_to_find - 1)

                                                spot_to_find = InStr(Trim(temp_string2), ",")
                                                If spot_to_find > 0 Then
                                                    temp_string2 = Left(Trim(temp_string2), spot_to_find - 1)
                                                    ' should get us the city 

                                                    pub_city = temp_string2
                                                    pub_seller_info = pub_seller_info_no_city & " " & pub_city

                                                End If

                                            End If


                                        End If

                                    End If
                                End If





                                End If


                        End If
                    End If



                    ' added in MSW - 6/282/21 ---------
                    If Trim(pub_seller_info) <> "" Then
                        If InStr(pub_seller_info, "</div<div") > 0 Then
                            spot_to_find = InStr(Trim(pub_seller_info), "</div<div")
                            If spot_to_find > 0 Then
                                pub_seller_info = Left(Trim(pub_seller_info), spot_to_find - 1)
                            End If
                        End If

                        If InStr(pub_seller_info, "</h3<div class=dealer-data") > 0 Then
                            pub_seller_info_no_city = Left(Trim(pub_seller_info), InStr(pub_seller_info, "</h3<div class=dealer-data") - 1)
                            pub_seller_info = Replace(pub_seller_info, "</h3<div class=dealer-data", "  ")
                        End If


                        If InStr(pub_seller_info, "</h3</a<div class=dealer-data") > 0 Then
                            pub_seller_info_no_city = Left(Trim(pub_seller_info), InStr(pub_seller_info, "</h3</a<div class=dealer-data") - 1)   ' make the no city one show up correctly - before cutting - msw - 6/30/21 
                            pub_seller_info = Replace(pub_seller_info, "</h3</a<div class=dealer-data", "  ")
                        End If



                    End If




                    ' added MSW -  3/21/22 
                    If Trim(pub_seller_info) <> "" Then

                        spot_to_find = InStr(Trim(pub_seller_info), "<!--")

                        spot_to_find2 = InStr(Trim(pub_seller_info), "</i>")
                        temp1_string = Trim(pub_seller_info)


                        If spot_to_find2 > 0 Then
                            pub_seller_info = Left(Trim(pub_seller_info), spot_to_find - 1)
                        End If

                        If spot_to_find2 > 0 Then
                            pub_seller_info = pub_seller_info & " " & Right(Trim(temp1_string), Len(Trim(temp1_string)) - spot_to_find2 - 3)
                        End If


                        spot_to_find = InStr(Trim(pub_seller_info), ">")

                        If InStr(pub_seller_info, "span") > 0 Then
                            If spot_to_find > 0 Then
                                pub_seller_info = Right(Trim(pub_seller_info), Len(Trim(pub_seller_info)) - spot_to_find) 
                            End If
                        End If
                    End If


















                    acpub_original_name = temp_ac_name & " " & pub_ser_no


                    temp_make = ""
                    temp_model = ""
                    temp_ac_name = replace_fm(temp_ac_name)
                    If Trim(temp_ac_name) <> "" Then
                        temp_year = Left(Trim(temp_ac_name), 4)
                        temp_temp = Right(Trim(temp_ac_name), Len(Trim(temp_ac_name)) - 5)
                        array_split_make = Split(Trim(temp_temp), " ")

                        If array_split_make.Length = 2 Then
                            temp_make = array_split_make(0)
                            temp_model = array_split_make(1)
                        ElseIf array_split_make.Length = 3 Then
                            temp_make = array_split_make(1)
                            temp_model = array_split_make(2)
                        ElseIf array_split_make.Length = 4 Then

                            If array_split_make(1) = "KING" And array_split_make(2) = "AIR" Then
                                temp_make = array_split_make(1) & " " & array_split_make(2) ' then make it king air 
                                temp_model = array_split_make(3)
                            Else
                                temp_make = array_split_make(2)
                                temp_model = array_split_make(3)
                            End If


                        ElseIf array_split_make.Length = 5 Then
                            temp_make = array_split_make(3)
                            temp_model = array_split_make(4)
                        Else
                            temp_temp = temp_temp
                        End If
                    End If


                    '  Response.Write("<br/>________________<br/>")
                    ' Response.Write("<br>" & pub_url)
                    '  Response.Write("<br>" & temp_ac_name)
                    '  Response.Write("<br>" & pub_ser_no)
                    '  Response.Write("<br>" & pub_desc)
                    ' Response.Write("<br>" & pub_aftt)
                    ' Response.Write("<br>" & temp_engine)
                    '  Response.Write("<br>" & pub_price)
                    '  Response.Write("<br>" & temp_eng)
                    '  Response.Write("<br>" & temp_av)
                    '  Response.Write("<br>" & pub_aftt)
                    '  Response.Write("<br>" & pub_seller_info)
                    '  Response.Write("<br/>________________<br/>")

                    If InStr(LCase(Trim(temp_make)), "diamond") > 0 Then
                        pub_ser_no = Replace(pub_ser_no, ".C", ".")  ' LOOKS LIKE they display the serials with a C while we dont
                    End If

                    If InStr(acpub_original_name, "ROBINSON R44 CLIPPER II") > 0 Then
                        acpub_original_name = acpub_original_name
                    End If

                    If InStr(pub_ser_no, "-") > 0 Then
                        pub_ser_no = pub_ser_no
                    End If

                    If InStr(pub_ser_no, "072") > 0 Then
                        pub_ser_no = pub_ser_no
                    End If

                    ' MSW - 3/21/22 
                    If InStr(LCase(pub_reg_no), "upon") > 0 Then
                        pub_reg_no = ""
                    End If


                    na_skip = False
                    If InStr(LCase(Trim(pub_ser_no)), "n/a") > 0 Then
                        na_skip = True
                    End If

                    If na_skip = False Then
                        temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
                        If temp_ac_id = 0 Then
                            temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
                            If temp_ac_id = 0 Then
                                temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                                If temp_ac_id = 0 Then
                                    temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                                    If temp_ac_id = 0 Then
                                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                                        If temp_ac_id = 0 Then
                                            If Trim(pub_reg_no) <> "" Then
                                                temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                                                If temp_ac_id = 0 Then
                                                    ' one last try 
                                                    temp_ac_id = find_ac_global_search("", temp_make, temp_model, pub_reg_no)
                                                    If temp_ac_id = 0 Then
                                                        temp_ac_id = temp_ac_id
                                                    End If
                                                End If
                                            End If


                                            If temp_ac_id = 0 Then
                                                If Trim(pub_ser_no) <> "" Then
                                                    temp_ac_id = find_ac_ac_search(pub_ser_no, temp_make, temp_model, "")
                                                End If

                                                If temp_ac_id = 0 Then
                                                    If Trim(pub_reg_no) <> "" Then
                                                        temp_ac_id = find_ac_ac_search("", temp_make, temp_model, pub_reg_no)
                                                    End If

                                                    If temp_ac_id = 0 Then
                                                        If Trim(pub_ser_no) <> "" And Left(Trim(pub_ser_no), 1) = "0" Then
                                                            temp_ac_id = find_ac_ac_search(Right(Trim(pub_ser_no), Len(Trim(pub_ser_no)) - 1), temp_make, temp_model, "")
                                                        End If
                                                    End If

                                                End If
                                            End If
                                            temp_ac_id = temp_ac_id



                                        End If

                                    End If
                                End If
                            End If

                        End If
                    End If

                    ' added in .. in case 
                    If Trim(pub_seller_info) <> "" And Trim(pub_seller_info_no_city) = "" Then
                        pub_seller_info_no_city = pub_seller_info
                    End If


                    If pub_comp_id = 0 Then
                        If Not IsNothing(pub_seller_info_no_city) Then
                            If Trim(pub_seller_info_no_city) <> "" Then
                                'if there is a line feed or break, then try to get just the company name 
                                If InStr(Trim(pub_seller_info_no_city), Asc(10)) > 0 Then
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info_no_city), InStr(Trim(pub_seller_info_no_city), Asc(10)) - 1))
                                ElseIf InStr(Trim(pub_seller_info_no_city), Asc(13)) > 0 Then
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info_no_city), InStr(Trim(pub_seller_info_no_city), Asc(13)) - 1))
                                Else
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info_no_city), 17))
                                End If
                            End If
                        End If
                    End If

                    If pub_comp_id = 0 And Trim(pub_seller_info_no_city <> "") Then
                        pub_comp_id = find_comp_id_previous_pub(pub_seller_info_no_city, 2)
                    End If

                    ' if still dont find - try again - MSW - 9 /27/21 - if they r different, try to get more specific 
                    If pub_comp_id = 0 And Trim(pub_seller_info_no_city) <> Trim(pub_seller_info) Then
                        If Trim(pub_city) <> "" Then
                            pub_comp_id = find_comp_id_global_search(Trim(pub_seller_info_no_city), pub_city)
                        End If
                    End If



                    pub_desc = ""
                    acpub_status = "N" ' default to N in case we find nothing - it stays N 
                    acpub_price_details = ""
                    If On_Naughty_List(temp_ac_name) = True Then
                        ' if its on naughtly list then excldue
                    Else
                        If temp_ac_id > 0 Then
                            Call find_ac_data(temp_ac_id)

                            If Trim(aftt_different) <> "" Then
                                pub_desc = pub_desc & aftt_different
                            End If

                            If Trim(landings_different) <> "" Then
                                pub_desc = pub_desc & landings_different
                            End If


                            If Trim(acpub_price_details) <> "" Then
                                pub_desc = pub_desc & " " & acpub_price_details
                            End If
                        Else
                            acpub_process_status = "For Sale Not Found – No AC Match"
                            acpub_status = "O"

                            If Trim(pub_aftt) <> "" Then
                                pub_desc = "Pub AFTT: " & pub_aftt
                            End If

                            If Trim(pub_price) <> "" Then
                                pub_desc = "Pub Price: " & pub_price
                            End If
                        End If



                        pub_ser_no = find_last_symbol(pub_ser_no, ">")
                        pub_reg_no = find_last_symbol(pub_reg_no, ">")
                        pub_price = find_last_symbol(pub_price, ">")
                        pub_aftt = find_last_symbol(pub_aftt, ">")
                        pub_seller_info = find_last_symbol(pub_seller_info, ">")
                        pub_comp_id = find_last_symbol(pub_comp_id, ">")
                        temp_ac_id = find_last_symbol(temp_ac_id, ">")
                        temp_make = temp_make
                        temp_model = temp_model
                        pub_url = pub_url


                        pub_seller_info = pub_seller_info
                        '  If Trim(pub_ser_no) = "" And Trim(pub_reg_no) = "" Then
                        '      pub_ser_no = pub_ser_no
                        '      original_string_text = original_string_text
                        '  ElseIf temp_ac_id = 0 Then
                        '      temp_ac_id = temp_ac_id
                        '  Else
                        '''''''   Call check_insert_ac_pub(temp_ac_id, 2)
                        '''
                        '  If InStr(string_text, "N353KM") > 0 Then 


                        Call check_insert_ac_pub(temp_ac_id, 2)



                        ' End If
                    End If




                Next


            End If



            Response.Write("<br/>7")


        Catch ex As Exception
            Response.Write("<br/>Error Controller: ")
            Response.Write(ex)
        Finally

        End Try

    End Function
    Public Function find_last_symbol(ByVal string_to_Replace As String, ByVal find_symbol As String)

        find_last_symbol = ""

        If InStr(Trim(string_to_Replace), find_symbol) > 0 Then
            string_to_Replace = Right(string_to_Replace, Len(Trim(string_to_Replace)) - InStr(string_to_Replace, find_symbol))
        End If

        string_to_Replace = Replace(string_to_Replace, "</span>", "")

        find_last_symbol = string_to_Replace

    End Function



    Public Function scrape_for_controller_8_2020(ByVal str As StreamReader, ByVal page_num As Integer)
        ' Dim Str As System.IO.Stream
        ' Dim srRead As System.IO.StreamReader



        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim i As Integer = 0
        Dim final_string As String = ""
        Dim original_string_text As String = ""
        Dim article_link As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim k As Integer = 0
        Dim skip_this As Boolean = False
        Dim extra_note As String = ""

        Dim temp_ac_name As String = ""
        Dim temp_engine As String = ""
        Dim temp_eng As String = ""
        Dim temp_av As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_make As String = ""
        Dim temp_temp As String
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split_make() As String
        Dim pass_try As String = "1"
        Dim temp_seller_info_string As String = ""


        Try



            'Dim wb As New WebBrowser
            'wb.ScrollBarsEnabled = False
            'wb.ScriptErrorsSuppressed = True
            'wb.Navigate("https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft?mdlx=Contains&Cond=All&SortOrder=35&scf=False&LS=5")
            ' While wb.ReadyState

            'End While 
            'Dim webBrowserForPrinting As New WebBrowser()
            'AddHandler webBrowserForPrinting.DocumentCompleted, New  _
            '   WebBrowserDocumentCompletedEventHandler(AddressOf PrintDocument)
            'wb.Document.DomDocument.ToString()


            'Dim webBrowserForPrinting As New WebBrowser()
            'webBrowserForPrinting.Url = New Uri("\test\help.html")


            'Dim sw As StreamWriter
            'Dim poststring = " "

            'Try
            '  sw = New StreamWriter(req.GetRequestStream)
            '  sw.Write(poststring)
            '  sw.Close()
            'Catch ex As Exception

            'End Try

            'Call PrintHelpPage()




            'Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/")
            'Dim resp As System.Net.WebResponse = req.GetResponse


            'Str = resp.GetResponseStream
            'srRead = New System.IO.StreamReader(Str)
            '' read all the text 
            'string_text = srRead.ReadToEnd().ToString

            'req.Abort()
            'resp.Close()
            'Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft/manufacturer/cessna/model/citation-latitude/")



            '
            ' Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/manufacturer/cessna?sortorder=27&SCF=False")

            ' req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft?mdlx=Contains&Cond=All&SortOrder=35&scf=False&LS=5")
            ' req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft/manufacturer/cessna/model/citation-latitude")





            ' THIS IS THE LINK FOR ALL JETS   -- DOESNT WORK --https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft/
            ' ALL CESSNAS - https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft/manufacturer/cessna/model/citation-latitude/
            'https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft/?sortorder=27&SCF=False/

            '  Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft/?sortorder=27&SCF=False%2f&page=" & page_num & "/")
            ' '' '' ''System.Threading.Thread.Sleep(10)
            ' '' '' ''Response.Flush()
            ' '' '' ''System.Threading.Thread.Sleep(10)

            ' '' '' ''Dim req As System.Net.WebRequest



            ' '' '' ''If Trim(type_string) = "last7" Then
            ' '' '' ''  req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft?sortorder=27&SCF=False&page=" & page_num & "/")
            ' '' '' ''Else
            ' '' '' ''  req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/" & type_string & "/?sortorder=27&SCF=False%2f&page=" & page_num & "")
            ' '' '' ''End If


            ' '' '' ''Dim resp As System.Net.WebResponse = req.GetResponse

            'Dim temp_url As String = ""

            'temp_url = "https://www.controller.com/listings/aircraft/for-sale/list/category/" & type_string & "/?sortorder=27&SCF=False%2f"
            'Using client As New Net.WebClient
            '  Dim reqparm As New Specialized.NameValueCollection
            '  reqparm.Add("page", "1")
            '  Dim responsebytes = client.UploadValues(temp_url, "POST", reqparm)

            'End Using

            Response.Write("<br/>1")

            '  Str = resp.GetResponseStream
            '   srRead = New System.IO.StreamReader(Str)

            '   srRead = str.ReadToEnd()


            ' read all the text 
            string_text = str.ReadToEnd()

            Response.Write("<br/>2")
            '   resp.Close()
            '    resp = Nothing
            '   req = Nothing

            string_text = string_text
            original_string_text = string_text

            ' GET HOW MANY PAGES THERE ARE --------------
            If page_num = 1 Then
                spot_to_find = InStr(string_text, "listings-total-pages")
                If spot_to_find > 0 Then
                    string_text = Right(string_text, Len(string_text) - spot_to_find - 22)
                    spot_to_find = InStr(string_text, "</span>")
                    If spot_to_find > 0 Then
                        string_text = Left(Trim(string_text), spot_to_find - 1)
                        If IsNumeric(Trim(string_text)) = True Then
                            total_pages = CInt(Trim(string_text))
                        Else
                            total_pages = 10
                        End If
                    Else
                        total_pages = 10
                    End If
                Else
                    total_pages = 10
                End If
                string_text = original_string_text
            End If


            If InStr(string_text, "General Listings") > 0 And acpub_controller_general_start = 0 Then
                acpub_controller_general_start = (page_num + 5)
            End If


            spot_to_find = InStr(string_text, "listings-list")



            If spot_to_find > 0 Then
                zero_Count += 1
                string_text = Right(string_text, Len(string_text) - spot_to_find)

                spot_to_find = InStr(string_text, "col listing-info cf")
                If spot_to_find > 0 Then
                    string_text = Right(string_text, Len(string_text) - spot_to_find - 2)

                    array_split = Split(string_text, "col listing-info cf")

                    For i = 1 To array_split.Length - 1
                        string_text = array_split(i)
                        original_string_text = string_text
                        acpub_count = acpub_count + 1

                        temp_ac_name = ""
                        temp_engine = ""
                        temp_eng = ""
                        temp_av = ""

                        pub_reg_no = ""
                        pub_ser_no = ""
                        pub_desc = ""
                        pub_price = ""
                        pub_aftt = ""
                        pub_seller_info = ""
                        pub_picture = ""
                        pub_status = ""
                        pub_landings = ""
                        pub_url = ""
                        has_pics = False
                        aftt_different = ""
                        landings_different = ""
                        pub_comp_id = 0


                        If InStr(string_text, "images/pictures.png") > 0 Then
                            has_pics = True
                        Else
                            has_pics = False
                        End If

                        spot_to_find = InStr(string_text, "a href=")
                        If spot_to_find > 0 Then
                            string_text = Right(string_text, Len(string_text) - spot_to_find - 8)

                            spot_to_find = InStr(string_text, ">")
                            If spot_to_find > 0 Then
                                pub_url = Left(Trim(string_text), spot_to_find - 2)
                                pub_url = Replace(pub_url, """", "")

                                pub_url = "https://www.controller.com/" & pub_url
                                pub_url = RTrim(LTrim(pub_url))

                                string_text = Right(string_text, Len(string_text) - spot_to_find)
                                spot_to_find = InStr(string_text, "<")
                                If spot_to_find > 0 Then
                                    temp_ac_name = Left(Trim(string_text), spot_to_find - 1)
                                End If







                                spot_to_find = InStr(string_text, "cf m-bottom-5")
                                If spot_to_find > 0 Then
                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 14)

                                    spot_to_find = InStr(string_text, "</div>")
                                    spot_to_find2 = InStr(Left(string_text, 20), "employee-category")
                                    If spot_to_find > 0 And spot_to_find2 > 0 Then
                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 6)
                                    End If

                                    spot_to_find = InStr(string_text, "</div>")
                                    If spot_to_find > 0 Then
                                        pub_desc = Left(Trim(string_text), spot_to_find - 1)

                                        spot_to_find = InStr(pub_desc, "<a class")
                                        If spot_to_find > 0 Then
                                            pub_desc = Left(Trim(pub_desc), spot_to_find - 1)
                                        End If
                                        If Trim(pub_desc) = "" Then
                                            pub_desc = pub_desc
                                        End If
                                        pub_desc = Replace(pub_desc, "’", "")
                                        pub_desc = Replace(pub_desc, "'", "")
                                        cutme(pub_desc)
                                        pub_desc = Left(Trim(pub_desc), 500)
                                        pub_desc = ""
                                    End If

                                End If

                                If InStr(pub_desc, "Price Reduced,") > 0 Then
                                    pub_url = pub_url
                                End If


                                If InStr(pub_url, "33267775") > 0 Then
                                    pub_url = pub_url
                                End If

                                pass_try = "1"

                                spot_to_find = InStr(Trim(string_text), "Serial #</span>")
                                If spot_to_find > 0 Then
                                    pass_try = "1"
                                ElseIf InStr(Trim(string_text), "Serial #:</span>") > 0 Then
                                    pass_try = "2"
                                ElseIf InStr(Trim(string_text), "Serial Number:</span>") > 0 Then
                                    pass_try = "3"
                                ElseIf InStr(Trim(string_text), "Serial Number</span>") > 0 Then
                                    pass_try = "4"
                                Else

                                    spot_to_find = InStr(Trim(string_text), "Registration #</span>")
                                    If spot_to_find = 0 Then
                                        spot_to_find = InStr(Trim(string_text), "Registration #:</span>")
                                        If spot_to_find > 0 Then
                                            pass_try = "2"
                                        Else
                                            ' if it doesnt have serial or reg, look for total time 
                                            spot_to_find = InStr(Trim(string_text), "Total Time</span>")
                                            If spot_to_find = 0 Then
                                                spot_to_find = InStr(Trim(string_text), "Total Time:</span>")
                                                If spot_to_find > 0 Then
                                                    pass_try = "2"
                                                Else
                                                    pass_try = pass_try
                                                End If
                                            Else
                                                pass_try = "1"
                                            End If

                                        End If
                                    Else
                                        pass_try = "1"
                                    End If



                                End If



                                ' added MSW - 6/28/21 
                                spot_to_find = InStr(temp_ac_name, "</h2>")
                                If spot_to_find > 0 Then
                                    temp_ac_name = Left(Trim(string_text), spot_to_find - 1)
                                End If


                                If pass_try = "1" Then   ' THEN WE ARE ASSUMING IT IS ON THE TYPE OF PAGE WITH NO : , this type has for sale after - PAGE 1 -----------
                                    '---------------------------------------------------------------------------------------------------------------------------

                                    spot_to_find = InStr(Trim(string_text), "Serial #</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 14)

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_ser_no = Replace(pub_ser_no, "n>", "")
                                            pub_ser_no = Replace(pub_ser_no, ">", "")
                                            pub_ser_no = Left(Trim(string_text), spot_to_find - 1)
                                            pub_ser_no = LTrim(RTrim(pub_ser_no))
                                            cutme(pub_ser_no)
                                        End If
                                    End If

                                    spot_to_find = InStr(Trim(string_text), "Registration #</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 22)

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_reg_no = Left(Trim(Trim(string_text)), spot_to_find - 1)
                                            pub_reg_no = LTrim(RTrim(pub_reg_no))
                                            cutme_LF(pub_reg_no)
                                            pub_reg_no = LTrim(RTrim(pub_reg_no))
                                        End If
                                    End If

                                    spot_to_find = InStr(Trim(string_text), "Total Time</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 18)

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Left(Trim(string_text), spot_to_find - 1)

                                            pub_aftt = Replace(pub_aftt, " ", "")
                                            pub_aftt = LTrim(pub_aftt)
                                            pub_aftt = RTrim(pub_aftt)
                                        End If
                                    End If

                                    Try
                                        spot_to_find = InStr(string_text, "For Sale Price:</span>")
                                        If spot_to_find > 0 Then
                                            string_text = Right(string_text, Len(string_text) - spot_to_find - 21)

                                            spot_to_find = InStr(string_text, "</div>")
                                            If spot_to_find > 0 Then
                                                pub_price = Left(Trim(string_text), spot_to_find - 2)
                                                pub_price = Replace(pub_price, "<span class=""nobr"">", "")
                                                pub_price = Replace(pub_price, "</span>", "")
                                                pub_price = Replace(pub_price, "USD $", "")
                                                spot_to_find = InStr(pub_price, "<div")
                                                If spot_to_find > 0 Then
                                                    pub_price = Left(Trim(pub_price), spot_to_find - 1)
                                                End If
                                                cutme(pub_price)
                                            End If
                                        End If

                                    Catch ex As Exception
                                        Response.Write(ex.ToString)
                                    End Try

                                ElseIf pass_try = "2" Then  '--------- THIS WILL BE THE SETUP OF THE PAGE 2's - THE ONES WITHOUT THE : and where for sale is first 
                                    '---------------------------------------------------------------------------------------------------------------------------
                                    '---------------------------------------------------------------------------------------------------------------------------


                                    Try
                                        spot_to_find = InStr(string_text, "For Sale Price:</span>")
                                        If spot_to_find > 0 Then
                                            string_text = Right(string_text, Len(string_text) - spot_to_find - 21)

                                            spot_to_find = InStr(string_text, "</div>")
                                            If spot_to_find > 0 Then
                                                pub_price = Left(Trim(string_text), spot_to_find - 2)
                                                pub_price = Replace(pub_price, "<span class=""nobr"">", "")
                                                pub_price = Replace(pub_price, "</span>", "")
                                                pub_price = Replace(pub_price, "USD $", "")
                                                spot_to_find = InStr(pub_price, "<div")
                                                If spot_to_find > 0 Then
                                                    pub_price = Left(Trim(pub_price), spot_to_find - 1)
                                                End If
                                                cutme(pub_price)
                                            End If
                                        End If

                                    Catch ex As Exception
                                        Response.Write(ex.ToString)
                                    End Try


                                    spot_to_find = InStr(Trim(string_text), "Serial #:</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 15)   ' changed from 16 to 15 - MSW - 3/3/20 

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_ser_no = Left(Trim(string_text), spot_to_find - 1)

                                            If InStr(string_text, ">") > 0 Then
                                                string_text = Replace(string_text, ">", "")
                                            End If

                                            pub_ser_no = LTrim(RTrim(pub_ser_no))
                                            cutme(pub_ser_no)
                                        End If
                                    End If

                                    spot_to_find = InStr(Trim(string_text), "Registration #:</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 22)

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_reg_no = Left(Trim(Trim(string_text)), spot_to_find - 1)
                                            pub_reg_no = LTrim(RTrim(pub_reg_no))
                                            cutme_LF(pub_reg_no)
                                            pub_reg_no = LTrim(RTrim(pub_reg_no))
                                        End If
                                    End If

                                    spot_to_find = InStr(Trim(string_text), "Total Time:</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 18)

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Left(Trim(string_text), spot_to_find - 1)

                                            pub_aftt = Replace(pub_aftt, " ", "")
                                            pub_aftt = LTrim(pub_aftt)
                                            pub_aftt = RTrim(pub_aftt)
                                        End If
                                    End If
                                ElseIf pass_try = "3" Then  '--------- THIS WILL BE THE SETUP OF THE PAGE 2's - THE ONES WITHOUT THE : and where for sale is first 
                                    '---------------------------------------------------------------------------------------------------------------------------
                                    '---------------------------------------------------------------------------------------------------------------------------

                                    Try
                                        spot_to_find = InStr(string_text, "For Sale Price:</span>")
                                        If spot_to_find > 0 Then
                                            string_text = Right(string_text, Len(string_text) - spot_to_find - 21)

                                            spot_to_find = InStr(string_text, "</div>")
                                            If spot_to_find > 0 Then
                                                pub_price = Left(Trim(string_text), spot_to_find - 2)
                                                pub_price = Replace(pub_price, "<span class=""nobr"">", "")
                                                pub_price = Replace(pub_price, "</span>", "")
                                                pub_price = Replace(pub_price, "USD $", "")
                                                spot_to_find = InStr(pub_price, "<div")
                                                If spot_to_find > 0 Then
                                                    pub_price = Left(Trim(pub_price), spot_to_find - 1)
                                                End If
                                                cutme(pub_price)
                                            End If
                                        End If

                                    Catch ex As Exception
                                        Response.Write(ex.ToString)
                                    End Try

                                    spot_to_find = InStr(Trim(string_text), "Serial Number:</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 20)

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_ser_no = Left(Trim(string_text), spot_to_find - 1)


                                            If InStr(string_text, ">") > 0 Then
                                                string_text = Replace(string_text, ">", "")
                                            End If
                                            pub_ser_no = LTrim(RTrim(pub_ser_no))
                                            cutme(pub_ser_no)
                                        End If
                                    End If

                                    spot_to_find = InStr(Trim(string_text), "Registration #:</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 22)

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_reg_no = Left(Trim(Trim(string_text)), spot_to_find - 1)
                                            pub_reg_no = LTrim(RTrim(pub_reg_no))
                                            cutme_LF(pub_reg_no)
                                            pub_reg_no = LTrim(RTrim(pub_reg_no))
                                        End If
                                    End If

                                    spot_to_find = InStr(Trim(string_text), "Total Time:</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 18)

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Left(Trim(string_text), spot_to_find - 1)

                                            pub_aftt = Replace(pub_aftt, " ", "")
                                            pub_aftt = LTrim(pub_aftt)
                                            pub_aftt = RTrim(pub_aftt)
                                        End If
                                    End If

                                ElseIf pass_try = "4" Then  '--------- THIS WILL BE THE SETUP OF THE PAGE 2's - THE ONES WITHOUT THE : and where for sale is first 
                                    '---------------------------------------------------------------------------------------------------------------------------
                                    '---------------------------------------------------------------------------------------------------------------------------



                                    spot_to_find = InStr(Trim(string_text), "Serial Number</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 21)

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_ser_no = Left(Trim(string_text), spot_to_find - 1)
                                            pub_ser_no = LTrim(RTrim(pub_ser_no))
                                            cutme(pub_ser_no)
                                        End If
                                    End If

                                    spot_to_find = InStr(Trim(string_text), "Registration #</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 22)

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_reg_no = Left(Trim(Trim(string_text)), spot_to_find - 1)
                                            pub_reg_no = LTrim(RTrim(pub_reg_no))
                                            cutme_LF(pub_reg_no)
                                            pub_reg_no = LTrim(RTrim(pub_reg_no))
                                        End If
                                    End If

                                    spot_to_find = InStr(Trim(string_text), "Total Time</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(Trim(string_text), Len(Trim(string_text)) - spot_to_find - 18)

                                        spot_to_find = InStr(Trim(string_text), "</div>")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Left(Trim(string_text), spot_to_find - 1)

                                            pub_aftt = Replace(pub_aftt, " ", "")
                                            pub_aftt = LTrim(pub_aftt)
                                            pub_aftt = RTrim(pub_aftt)
                                        End If
                                    End If


                                    spot_to_find = InStr(string_text, "col dealer-info")
                                    If spot_to_find > 0 Then
                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 20)

                                        spot_to_find = InStr(string_text, ">")
                                        If spot_to_find > 0 Then
                                            string_text = Right(string_text, Len(string_text) - spot_to_find - 1)

                                            spot_to_find = InStr(string_text, ">")
                                            If spot_to_find > 0 Then
                                                string_text = Right(string_text, Len(string_text) - spot_to_find)

                                                spot_to_find = InStr(string_text, "</") ' get whatever is after the name, could be span or href
                                                If spot_to_find > 0 Then
                                                    pub_seller_info = Left(Trim(string_text), spot_to_find - 1)
                                                End If

                                                spot_to_find = InStr(pub_seller_info, ">")
                                                If spot_to_find > 0 Then
                                                    pub_seller_info = Right(pub_seller_info, Len(pub_seller_info) - spot_to_find)
                                                End If

                                            End If
                                        End If



                                    End If

                                    Try
                                        spot_to_find = InStr(string_text, "For Sale Price:</span>")
                                        If spot_to_find > 0 Then
                                            string_text = Right(string_text, Len(string_text) - spot_to_find - 21)

                                            spot_to_find = InStr(string_text, "</div>")
                                            If spot_to_find > 0 Then
                                                pub_price = Left(Trim(string_text), spot_to_find - 2)
                                                pub_price = Replace(pub_price, "<span class=""nobr"">", "")
                                                pub_price = Replace(pub_price, "</span>", "")
                                                pub_price = Replace(pub_price, "USD $", "")
                                                spot_to_find = InStr(pub_price, "<div")
                                                If spot_to_find > 0 Then
                                                    pub_price = Left(Trim(pub_price), spot_to_find - 1)
                                                End If
                                                cutme(pub_price)
                                            End If
                                        End If

                                    Catch ex As Exception
                                        Response.Write(ex.ToString)
                                    End Try

                                Else ' MAYBE THERE IS NO SERIAL NUMBER - so check 
                                    '---------------------------------------------------------------------------------------------------------------------------


                                End If



                                If InStr(acpub_original_name, "1991 BELL 212") > 0 Then
                                    acpub_original_name = acpub_original_name
                                End If

                                If Trim(pub_reg_no) <> "" Then
                                    If InStr(pub_reg_no, Chr(13) & Chr(10)) > 0 Then
                                        pub_reg_no = Left(Trim(pub_reg_no), InStr(pub_reg_no, Chr(13) & Chr(10)) - 1)
                                    End If

                                    If InStr(pub_reg_no, Chr(10)) > 0 Then
                                        pub_reg_no = Left(Trim(pub_reg_no), InStr(pub_reg_no, Chr(10)) - 1)
                                    End If
                                End If





                                spot_to_find = InStr(string_text, "Engine(s):</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 17)

                                    spot_to_find = InStr(string_text, "</div>")
                                    If spot_to_find > 0 Then
                                        temp_eng = Left(Trim(string_text), spot_to_find - 1)
                                    End If
                                End If


                                If Trim(pub_aftt) = "" Then
                                    spot_to_find = InStr(string_text, "Airframe:</span>")
                                    If spot_to_find > 0 Then
                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 16)

                                        spot_to_find = InStr(string_text, "</div>")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Left(Trim(string_text), spot_to_find - 1)
                                            spot_to_find = InStr(pub_aftt, ":")
                                            If spot_to_find > 0 Then
                                                pub_aftt = Right(pub_aftt, Len(pub_aftt) - spot_to_find)
                                                cutme(pub_aftt)
                                                spot_to_find = InStr(LCase(pub_aftt), "hours")
                                                If spot_to_find > 0 Then
                                                    pub_aftt = Left(Trim(pub_aftt), spot_to_find + 4)
                                                Else
                                                    spot_to_find = InStr(pub_aftt, "Total")
                                                    If spot_to_find > 0 Then
                                                        pub_aftt = Left(Trim(pub_aftt), spot_to_find - 2)
                                                    ElseIf InStr(pub_aftt, "Landings") > 0 Then
                                                        pub_aftt = Left(Trim(pub_aftt), InStr(pub_aftt, "Landings") - 2)
                                                    Else
                                                        spot_to_find = spot_to_find
                                                    End If
                                                End If
                                            End If
                                        End If

                                        pub_aftt = Replace(pub_aftt, "Airframe", "")
                                        pub_aftt = Replace(pub_aftt, "Total", "")
                                        pub_aftt = Replace(pub_aftt, "Time", "")

                                        If IsNumeric(Left(Trim(pub_aftt), 6)) = True Then
                                            pub_aftt = Left(Trim(pub_aftt), 6)
                                        ElseIf IsNumeric(Left(Trim(pub_aftt), 5)) = True Then
                                            pub_aftt = Left(Trim(pub_aftt), 5)
                                        ElseIf IsNumeric(Left(Trim(pub_aftt), 4)) = True Then
                                            pub_aftt = Left(Trim(pub_aftt), 4)
                                        ElseIf IsNumeric(Left(Trim(pub_aftt), 3)) = True Then
                                            pub_aftt = Left(Trim(pub_aftt), 3)
                                        Else
                                            pub_aftt = ""
                                        End If
                                    End If
                                End If


                                spot_to_find = InStr(string_text, "Avionics/Radios:</span>")
                                If spot_to_find > 0 Then
                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 24)

                                    spot_to_find = InStr(string_text, "</div>")
                                    If spot_to_find > 0 Then
                                        temp_av = Left(Trim(string_text), spot_to_find - 1)
                                    End If
                                End If

                                If Trim(pub_seller_info) = "" Then
                                    spot_to_find = InStr(original_string_text, "col dealer-info")
                                    If spot_to_find > 0 Then
                                        string_text = Right(original_string_text, Len(original_string_text) - spot_to_find - 20)

                                        spot_to_find = InStr(string_text, "div Class=""bold"">")
                                        If spot_to_find > 0 Then
                                            string_text = Right(string_text, Len(string_text) - spot_to_find - 19)

                                            temp_seller_info_string = string_text

                                            spot_to_find = InStr(string_text, ">")
                                            If spot_to_find > 0 Then
                                                string_text = Right(string_text, Len(string_text) - spot_to_find)

                                                spot_to_find = InStr(string_text, "</") ' get whatever is after the name, could be span or href
                                                If spot_to_find > 0 Then
                                                    pub_seller_info = Left(Trim(string_text), spot_to_find - 1)
                                                End If

                                                If Len(Trim(pub_seller_info)) > 40 Then

                                                    spot_to_find = InStr(string_text, "<div>")
                                                    If spot_to_find > 0 Then
                                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 4)

                                                        spot_to_find = InStr(string_text, "</div>")
                                                        If spot_to_find > 0 Then
                                                            pub_seller_info &= " " & Left(Trim(string_text), spot_to_find - 1)
                                                        End If
                                                    End If
                                                Else
                                                    pub_seller_info = pub_seller_info
                                                End If

                                                spot_to_find = InStr(temp_seller_info_string, pub_seller_info)
                                                If spot_to_find > 0 Then
                                                    string_text = Right(temp_seller_info_string, Len(temp_seller_info_string) - spot_to_find - 5)

                                                    spot_to_find = InStr(string_text, "</div>")
                                                    If spot_to_find > 0 Then
                                                        string_text = Right(string_text, Len(string_text) - spot_to_find - 6)

                                                        spot_to_find = InStr(string_text, "</div>")
                                                        If spot_to_find > 0 Then
                                                            string_text = Left(Trim(string_text), spot_to_find - 1)
                                                            string_text = Replace(string_text, "<div>", "")
                                                            string_text = Replace(string_text, "div>", "")
                                                            string_text = Replace(string_text, "", "")
                                                            pub_seller_info &= ", " & string_text
                                                        End If


                                                    End If

                                                End If

                                                spot_to_find = InStr(temp_seller_info_string, "Phone:</span>")
                                                If spot_to_find > 0 Then
                                                    string_text = Right(temp_seller_info_string, Len(temp_seller_info_string) - spot_to_find - 14)


                                                    spot_to_find = InStr(Trim(string_text), "phonetype")
                                                    If spot_to_find > 0 Then
                                                        string_text = " " & Left(Trim(string_text), spot_to_find - 2)

                                                        spot_to_find = InStr(string_text, "href=")
                                                        If spot_to_find > 0 Then
                                                            string_text = Right(string_text, Len(string_text) - spot_to_find - 6)
                                                            string_text = Replace(string_text, "tel:+", "")
                                                            string_text = Replace(string_text, "el:+", "")
                                                            string_text = Replace(string_text, "el:", "")
                                                            string_text = Replace(string_text, """", "")
                                                            pub_seller_info &= ", " & string_text
                                                        End If


                                                    End If


                                                End If


                                            End If
                                        End If
                                    End If
                                Else
                                    ' WE HAVE SELLER INFO 
                                    pub_seller_info = pub_seller_info
                                End If




                            End If

                        End If

                        pub_seller_info = Replace(pub_seller_info, "'", "''")
                        pub_seller_info = replace_fm(pub_seller_info)

                        acpub_original_name = temp_ac_name & " " & pub_ser_no


                        temp_make = ""
                        temp_model = ""
                        temp_ac_name = replace_fm(temp_ac_name)
                        If Trim(temp_ac_name) <> "" Then
                            temp_year = Left(Trim(temp_ac_name), 4)
                            temp_temp = Right(Trim(temp_ac_name), Len(Trim(temp_ac_name)) - 5)
                            array_split_make = Split(Trim(temp_temp), " ")

                            If array_split_make.Length = 2 Then
                                temp_make = array_split_make(0)
                                temp_model = array_split_make(1)
                            ElseIf array_split_make.Length = 3 Then
                                temp_make = array_split_make(1)
                                temp_model = array_split_make(2)
                            ElseIf array_split_make.Length = 4 Then
                                temp_make = array_split_make(2)
                                temp_model = array_split_make(3)
                            ElseIf array_split_make.Length = 5 Then
                                temp_make = array_split_make(3)
                                temp_model = array_split_make(4)
                            Else
                                temp_temp = temp_temp
                            End If
                        End If


                        '  Response.Write("<br/>________________<br/>")
                        ' Response.Write("<br>" & pub_url)
                        '  Response.Write("<br>" & temp_ac_name)
                        '  Response.Write("<br>" & pub_ser_no)
                        '  Response.Write("<br>" & pub_desc)
                        ' Response.Write("<br>" & pub_aftt)
                        ' Response.Write("<br>" & temp_engine)
                        '  Response.Write("<br>" & pub_price)
                        '  Response.Write("<br>" & temp_eng)
                        '  Response.Write("<br>" & temp_av)
                        '  Response.Write("<br>" & pub_aftt)
                        '  Response.Write("<br>" & pub_seller_info)
                        '  Response.Write("<br/>________________<br/>")

                        If InStr(LCase(Trim(temp_make)), "diamond") > 0 Then
                            pub_ser_no = Replace(pub_ser_no, ".C", ".")  ' LOOKS LIKE they display the serials with a C while we dont
                        End If

                        If InStr(acpub_original_name, "ROBINSON R44 CLIPPER II") > 0 Then
                            acpub_original_name = acpub_original_name
                        End If

                        If InStr(pub_ser_no, "-") > 0 Then
                            pub_ser_no = pub_ser_no
                        End If

                        If InStr(pub_ser_no, "072") > 0 Then
                            pub_ser_no = pub_ser_no
                        End If


                        temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
                        If temp_ac_id = 0 Then
                            temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
                            If temp_ac_id = 0 Then
                                temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                                If temp_ac_id = 0 Then
                                    temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                                    If temp_ac_id = 0 Then
                                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                                        If temp_ac_id = 0 Then
                                            If Trim(pub_reg_no) <> "" Then
                                                temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                                                If temp_ac_id = 0 Then
                                                    ' one last try 
                                                    temp_ac_id = find_ac_global_search("", temp_make, temp_model, pub_reg_no)
                                                    If temp_ac_id = 0 Then
                                                        temp_ac_id = temp_ac_id
                                                    End If
                                                End If
                                            End If
                                        End If

                                    End If
                                End If
                            End If

                        End If



                        If pub_comp_id = 0 Then
                            If Not IsNothing(pub_seller_info) Then
                                If Trim(pub_seller_info) <> "" Then
                                    'if there is a line feed or break, then try to get just the company name 
                                    If InStr(Trim(pub_seller_info), Asc(10)) > 0 Then
                                        pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), InStr(Trim(pub_seller_info), Asc(10)) - 1))
                                    ElseIf InStr(Trim(pub_seller_info), Asc(13)) > 0 Then
                                        pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), InStr(Trim(pub_seller_info), Asc(13)) - 1))
                                    Else
                                        pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), 17))
                                    End If
                                End If
                            End If
                        End If

                        If pub_comp_id = 0 And Trim(pub_seller_info <> "") Then
                            pub_comp_id = find_comp_id_previous_pub(pub_seller_info, 2)
                        End If


                        pub_desc = ""
                        acpub_status = "N" ' default to N in case we find nothing - it stays N 
                        acpub_price_details = ""
                        If On_Naughty_List(temp_ac_name) = True Then
                            ' if its on naughtly list then excldue
                        Else
                            If temp_ac_id > 0 Then
                                Call find_ac_data(temp_ac_id)

                                If Trim(aftt_different) <> "" Then
                                    pub_desc = pub_desc & aftt_different
                                End If

                                If Trim(acpub_price_details) <> "" Then
                                    pub_desc = pub_desc & " " & acpub_price_details
                                End If
                            Else
                                acpub_process_status = "For Sale Not Found – No AC Match"
                                acpub_status = "O"

                                If Trim(pub_aftt) <> "" Then
                                    pub_desc = "Pub AFTT: " & pub_aftt
                                End If

                                If Trim(pub_price) <> "" Then
                                    pub_desc = "Pub Price: " & pub_price
                                End If
                            End If


                            pub_seller_info = pub_seller_info
                            Call check_insert_ac_pub(temp_ac_id, 2)

                        End If




                    Next


                End If
            Else
                temp_ac_id = temp_ac_id ' NO PAGE ITEMS FOUND - 
                non_zero_Count += 1
            End If



            Response.Write("<br/>7")


        Catch ex As Exception
            Response.Write("<br/>Error Controller: ")
            Response.Write(ex)
        Finally

        End Try

    End Function

    Public Function scrape_for_controller_original(ByVal page_num As Long, ByVal type_string As String)
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader



    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim original_string_text As String = ""
    Dim article_link As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim k As Integer = 0
    Dim skip_this As Boolean = False
    Dim extra_note As String = ""

    Dim temp_ac_name As String = ""
    Dim temp_engine As String = ""
    Dim temp_eng As String = ""
    Dim temp_av As String = ""
    Dim temp_ac_id As Long = 0
    Dim temp_make As String = ""
    Dim temp_temp As String
    Dim temp_model As String = ""
    Dim temp_year As String = ""
    Dim array_split_make() As String




    Try



      'Dim wb As New WebBrowser
      'wb.ScrollBarsEnabled = False
      'wb.ScriptErrorsSuppressed = True
      'wb.Navigate("https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft?mdlx=Contains&Cond=All&SortOrder=35&scf=False&LS=5")
      ' While wb.ReadyState

      'End While 
      'Dim webBrowserForPrinting As New WebBrowser()
      'AddHandler webBrowserForPrinting.DocumentCompleted, New  _
      '   WebBrowserDocumentCompletedEventHandler(AddressOf PrintDocument)
      'wb.Document.DomDocument.ToString()


      'Dim webBrowserForPrinting As New WebBrowser()
      'webBrowserForPrinting.Url = New Uri("\test\help.html")


      'Dim sw As StreamWriter
      'Dim poststring = " "

      'Try
      '  sw = New StreamWriter(req.GetRequestStream)
      '  sw.Write(poststring)
      '  sw.Close()
      'Catch ex As Exception

      'End Try

      'Call PrintHelpPage()




      'Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/")
      'Dim resp As System.Net.WebResponse = req.GetResponse


      'Str = resp.GetResponseStream
      'srRead = New System.IO.StreamReader(Str)
      '' read all the text 
      'string_text = srRead.ReadToEnd().ToString

      'req.Abort()
      'resp.Close()
      'Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft/manufacturer/cessna/model/citation-latitude/")



      '
      ' Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/manufacturer/cessna?sortorder=27&SCF=False")

      ' req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft?mdlx=Contains&Cond=All&SortOrder=35&scf=False&LS=5")
      ' req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft/manufacturer/cessna/model/citation-latitude")





      ' THIS IS THE LINK FOR ALL JETS   -- DOESNT WORK --https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft/
      ' ALL CESSNAS - https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft/manufacturer/cessna/model/citation-latitude/
      'https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft/?sortorder=27&SCF=False/

      '  Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft/?sortorder=27&SCF=False%2f&page=" & page_num & "/")
      System.Threading.Thread.Sleep(10)
      Response.Flush()
      System.Threading.Thread.Sleep(10)

      Dim req As System.Net.WebRequest



      If Trim(type_string) = "last7" Then
        req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/13/aircraft?sortorder=27&SCF=False&page=" & page_num & "/")
      Else
        req = System.Net.WebRequest.Create("https://www.controller.com/listings/aircraft/for-sale/list/category/" & type_string & "/?sortorder=27&SCF=False%2f&page=" & page_num & "")
      End If


      Dim resp As System.Net.WebResponse = req.GetResponse

      'Dim temp_url As String = ""

      'temp_url = "https://www.controller.com/listings/aircraft/for-sale/list/category/" & type_string & "/?sortorder=27&SCF=False%2f"
      'Using client As New Net.WebClient
      '  Dim reqparm As New Specialized.NameValueCollection
      '  reqparm.Add("page", "1")
      '  Dim responsebytes = client.UploadValues(temp_url, "POST", reqparm)

      'End Using



      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)




      ' read all the text 
      string_text = srRead.ReadToEnd().ToString

      resp.Close()
      resp = Nothing
      req = Nothing

      string_text = string_text
      original_string_text = string_text

      ' GET HOW MANY PAGES THERE ARE --------------
      If page_num = 1 Then
        spot_to_find = InStr(string_text, "listings-total-pages")
        If spot_to_find > 0 Then
          string_text = Right(string_text, Len(string_text) - spot_to_find - 22)
          spot_to_find = InStr(string_text, "</span>")
          If spot_to_find > 0 Then
            string_text = Left(Trim(string_text), spot_to_find - 1)
            If IsNumeric(Trim(string_text)) = True Then
              total_pages = CInt(Trim(string_text))
            Else
              total_pages = 10
            End If
          Else
            total_pages = 10
          End If
        Else
          total_pages = 10
        End If
        string_text = original_string_text
      End If


      If InStr(string_text, "General Listings") > 0 And acpub_controller_general_start = 0 Then
        acpub_controller_general_start = (page_num + 5)
      End If


      spot_to_find = InStr(string_text, "listings-list")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find)

        spot_to_find = InStr(string_text, "col listing-info cf")
        If spot_to_find > 0 Then
          string_text = Right(string_text, Len(string_text) - spot_to_find - 2)

          array_split = Split(string_text, "col listing-info cf")

          For i = 1 To array_split.Length - 1
            string_text = array_split(i)
            original_string_text = string_text
            acpub_count = acpub_count + 1

            temp_ac_name = ""
            temp_engine = ""
            temp_eng = ""
            temp_av = ""

            pub_reg_no = ""
            pub_ser_no = ""
            pub_desc = ""
            pub_price = ""
            pub_aftt = ""
            pub_seller_info = ""
            pub_picture = ""
            pub_status = ""
            pub_url = ""
            has_pics = False
            aftt_different = ""


            If InStr(string_text, "images/pictures.png") > 0 Then
              has_pics = True
            Else
              has_pics = False
            End If

            spot_to_find = InStr(string_text, "a href=")
            If spot_to_find > 0 Then
              string_text = Right(string_text, Len(string_text) - spot_to_find - 8)

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                pub_url = Left(Trim(string_text), spot_to_find - 3)
                pub_url = "https://www.controller.com/" & pub_url

                string_text = Right(string_text, Len(string_text) - spot_to_find)
                spot_to_find = InStr(string_text, "<")
                If spot_to_find > 0 Then
                  temp_ac_name = Left(Trim(string_text), spot_to_find - 1)
                End If


                spot_to_find = InStr(string_text, "cf m-bottom-5")
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 14)

                  spot_to_find = InStr(string_text, "</div>")
                  If spot_to_find > 0 Then
                    pub_desc = Left(Trim(string_text), spot_to_find - 1)

                    spot_to_find = InStr(pub_desc, "<a class")
                    If spot_to_find > 0 Then
                      pub_desc = Left(Trim(pub_desc), spot_to_find - 1)
                    End If
                    If Trim(pub_desc) = "" Then
                      pub_desc = pub_desc
                    End If
                    pub_desc = Replace(pub_desc, "’", "")
                    pub_desc = Replace(pub_desc, "'", "")
                    cutme(pub_desc)
                  End If

                End If

                Try
                  spot_to_find = InStr(string_text, "For Sale Price:</span>")
                  If spot_to_find > 0 Then
                    string_text = Right(string_text, Len(string_text) - spot_to_find - 21)

                    spot_to_find = InStr(string_text, "</div>")
                    If spot_to_find > 0 Then
                      pub_price = Left(Trim(string_text), spot_to_find - 2)
                      pub_price = Replace(pub_price, "<span class=""nobr"">", "")
                      pub_price = Replace(pub_price, "</span>", "")
                      pub_price = Replace(pub_price, "USD $", "")
                      spot_to_find = InStr(pub_price, "<div")
                      If spot_to_find > 0 Then
                        pub_price = Left(Trim(pub_price), spot_to_find - 1)
                      End If
                      cutme(pub_price)
                    End If
                  End If

                Catch ex As Exception
                  Response.Write(ex.ToString)
                End Try

                spot_to_find = InStr(string_text, "Serial #:</span>")
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 16)

                  spot_to_find = InStr(string_text, "</div>")
                  If spot_to_find > 0 Then
                    pub_ser_no = Left(Trim(string_text), spot_to_find - 1)
                  End If
                End If


                spot_to_find = InStr(string_text, "Total Time:</span>")
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 18)

                  spot_to_find = InStr(string_text, "</div>")
                  If spot_to_find > 0 Then
                    pub_aftt = Left(Trim(string_text), spot_to_find - 1)
                  End If
                End If

                spot_to_find = InStr(string_text, "Registration #:</span>")
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 22)

                  spot_to_find = InStr(string_text, "</div>")
                  If spot_to_find > 0 Then
                    pub_reg_no = Left(Trim(string_text), spot_to_find - 1)
                  End If
                End If



                spot_to_find = InStr(string_text, "Engine(s):</span>")
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 17)

                  spot_to_find = InStr(string_text, "</div>")
                  If spot_to_find > 0 Then
                    temp_eng = Left(Trim(string_text), spot_to_find - 1)
                  End If
                End If


                If Trim(pub_aftt) = "" Then
                  spot_to_find = InStr(string_text, "Airframe:</span>")
                  If spot_to_find > 0 Then
                    string_text = Right(string_text, Len(string_text) - spot_to_find - 16)

                    spot_to_find = InStr(string_text, "</div>")
                    If spot_to_find > 0 Then
                      pub_aftt = Left(Trim(string_text), spot_to_find - 1)
                      spot_to_find = InStr(pub_aftt, ":")
                      If spot_to_find > 0 Then
                        pub_aftt = Right(pub_aftt, Len(pub_aftt) - spot_to_find)
                        cutme(pub_aftt)
                        spot_to_find = InStr(LCase(pub_aftt), "hours")
                        If spot_to_find > 0 Then
                          pub_aftt = Left(Trim(pub_aftt), spot_to_find + 4)
                        Else
                          spot_to_find = InStr(pub_aftt, "Total")
                          If spot_to_find > 0 Then
                            pub_aftt = Left(Trim(pub_aftt), spot_to_find - 2)
                          ElseIf InStr(pub_aftt, "Landings") > 0 Then
                            pub_aftt = Left(Trim(pub_aftt), InStr(pub_aftt, "Landings") - 2)
                          Else
                            spot_to_find = spot_to_find
                          End If
                        End If
                      End If
                    End If

                    pub_aftt = Replace(pub_aftt, "Airframe", "")
                    pub_aftt = Replace(pub_aftt, "Total", "")
                    pub_aftt = Replace(pub_aftt, "Time", "")

                    If IsNumeric(Left(Trim(pub_aftt), 6)) = True Then
                      pub_aftt = Left(Trim(pub_aftt), 6)
                    ElseIf IsNumeric(Left(Trim(pub_aftt), 5)) = True Then
                      pub_aftt = Left(Trim(pub_aftt), 5)
                    ElseIf IsNumeric(Left(Trim(pub_aftt), 4)) = True Then
                      pub_aftt = Left(Trim(pub_aftt), 4)
                    ElseIf IsNumeric(Left(Trim(pub_aftt), 3)) = True Then
                      pub_aftt = Left(Trim(pub_aftt), 3)
                    Else
                      pub_aftt = ""
                    End If
                  End If
                End If


                spot_to_find = InStr(string_text, "Avionics/Radios:</span>")
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 24)

                  spot_to_find = InStr(string_text, "</div>")
                  If spot_to_find > 0 Then
                    temp_av = Left(Trim(string_text), spot_to_find - 1)
                  End If
                End If

                spot_to_find = InStr(string_text, "col dealer-info")
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 20)

                  spot_to_find = InStr(string_text, ">")
                  If spot_to_find > 0 Then
                    string_text = Right(string_text, Len(string_text) - spot_to_find - 1)

                    spot_to_find = InStr(string_text, ">")
                    If spot_to_find > 0 Then
                      string_text = Right(string_text, Len(string_text) - spot_to_find)

                      spot_to_find = InStr(string_text, "</") ' get whatever is after the name, could be span or href
                      If spot_to_find > 0 Then
                        pub_seller_info = Left(Trim(string_text), spot_to_find - 1)
                      End If


                      spot_to_find = InStr(string_text, "<div>")
                      If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 4)

                        spot_to_find = InStr(string_text, "</div>")
                        If spot_to_find > 0 Then
                          pub_seller_info &= " " & Left(Trim(string_text), spot_to_find - 1)
                        End If
                      End If


                      spot_to_find = InStr(string_text, "Phone: </span>")
                      If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 14)
                        spot_to_find = InStr(string_text, "</span>")
                        If spot_to_find > 0 Then
                          string_text = Right(string_text, Len(string_text) - spot_to_find - 7)
                          spot_to_find = InStr(string_text, "</a>")
                          If spot_to_find > 0 Then
                            pub_seller_info &= " " & Left(Trim(string_text), spot_to_find - 1)
                          End If

                        End If


                      End If


                    End If
                  End If
                End If


              End If

            End If

            pub_seller_info = Replace(pub_seller_info, "'", "''")

            acpub_original_name = temp_ac_name & " " & pub_ser_no

            temp_year = Left(Trim(temp_ac_name), 4)
            temp_temp = Right(Trim(temp_ac_name), Len(Trim(temp_ac_name)) - 5)
            array_split_make = Split(Trim(temp_temp), " ")

            If array_split_make.Length = 2 Then
              temp_make = array_split_make(0)
              temp_model = array_split_make(1)
            ElseIf array_split_make.Length = 3 Then
              temp_make = array_split_make(1)
              temp_model = array_split_make(2)
            ElseIf array_split_make.Length = 4 Then
              temp_make = array_split_make(2)
              temp_model = array_split_make(3)
            ElseIf array_split_make.Length = 5 Then
              temp_make = array_split_make(3)
              temp_model = array_split_make(4)
            Else
              temp_temp = temp_temp
            End If


            '  Response.Write("<br/>________________<br/>")
            ' Response.Write("<br>" & pub_url)
            '  Response.Write("<br>" & temp_ac_name)
            '  Response.Write("<br>" & pub_ser_no)
            '  Response.Write("<br>" & pub_desc)
            ' Response.Write("<br>" & pub_aftt)
            ' Response.Write("<br>" & temp_engine)
            '  Response.Write("<br>" & pub_price)
            '  Response.Write("<br>" & temp_eng)
            '  Response.Write("<br>" & temp_av)
            '  Response.Write("<br>" & pub_aftt)
            '  Response.Write("<br>" & pub_seller_info)
            '  Response.Write("<br/>________________<br/>")

            If InStr(LCase(Trim(temp_make)), "diamond") > 0 Then
              pub_ser_no = Replace(pub_ser_no, ".C", ".")  ' LOOKS LIKE they display the serials with a C while we dont
            End If


            temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
            If temp_ac_id = 0 Then
              temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
              If temp_ac_id = 0 Then
                temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                If temp_ac_id = 0 Then
                  temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                  If temp_ac_id = 0 Then
                    temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                    If temp_ac_id = 0 Then
                      If Trim(pub_reg_no) <> "" Then
                        temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                        If temp_ac_id = 0 Then
                          temp_ac_id = temp_ac_id
                        End If
                      End If
                    End If

                  End If
                End If
              End If

            End If

            acpub_price_details = ""
                        If On_Naughty_List(temp_ac_name) = True Then
                            ' if its on naughtly list then excldue
                        Else
                            If temp_ac_id > 0 Then
                Call find_ac_data(temp_ac_id)
              Else
                acpub_process_status = "For Sale Not Found – No AC Match"
                acpub_status = "O"
              End If

              If Trim(aftt_different) <> "" Then
                pub_desc = pub_desc & aftt_different
              End If

              If Trim(acpub_price_details) <> "" Then
                pub_desc = pub_desc & " " & acpub_price_details
              End If

              Call check_insert_ac_pub(temp_ac_id, 2)
            End If




          Next


        End If
      Else
        temp_ac_id = temp_ac_id ' NO PAGE ITEMS FOUND - 
      End If





    Catch ex As Exception
      Response.Write("<br/>Error Controller: ")
      Response.Write(ex)
    Finally

    End Try

  End Function
  Public Function On_Naughty_List(ByVal temp_ac_name As String) As Boolean
    On_Naughty_List = False

        For i = 0 To Naughty_List_Size - 1
            If InStr(Trim(UCase(temp_ac_name)), Trim(UCase(Naughty_List_Of_Models(i)))) > 0 Then
                On_Naughty_List = True
                i = Naughty_List_Size
            End If
        Next

    End Function

  Public Function scrape_for_ASO()
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader



    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim original_string_text As String = ""
    Dim article_link As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim k As Integer = 0
    Dim skip_this As Boolean = False
    Dim extra_note As String = ""

    Dim temp_ac_name As String = ""
    Dim temp_engine As String = ""
    Dim temp_eng As String = ""
    Dim temp_av As String = ""
    Dim temp_ac_id As Long = 0
    Dim temp_make As String = ""
    Dim temp_temp As String
    Dim temp_model As String = ""
    Dim temp_year As String = ""
    Dim array_split_make() As String
    Dim tcount As Integer = 0




    Try


      System.Threading.Thread.Sleep(10)
      Response.Flush()
      System.Threading.Thread.Sleep(10)

      Dim req As System.Net.WebRequest

      '  req = System.Net.WebRequest.Create("https://www.trade-a-plane.com/search?s-type=aircraft&s-advanced=yes&s-custom=&sale_status=For+Sale&category_level1=Jets&days_old-max=10/")

      req = System.Net.WebRequest.Create("https://www.aso.com/listings/AircraftListings.aspx?ll=tru")
 
      Dim resp As System.Net.WebResponse = req.GetResponse


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString

      resp.Close()
      resp = Nothing
      req = Nothing

      string_text = string_text
      original_string_text = string_text


      spot_to_find = InStr(string_text, "photoListingsDescription")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find - 2)

        array_split = Split(string_text, "photoListingsDescription")

        For i = 1 To array_split.Length - 1
          tcount = tcount + 1  ' do every third one 
          string_text = array_split(i)
          original_string_text = string_text

          acpub_count = acpub_count + 1
         

          If tcount = 2 Then
            temp_ac_name = ""
            temp_engine = ""
            temp_eng = ""
            temp_av = ""

            pub_reg_no = ""
            pub_ser_no = ""
            pub_desc = ""
            pub_price = ""
            pub_aftt = ""
            pub_seller_info = ""
            pub_picture = ""
            pub_status = ""
            pub_url = ""
            has_pics = False
            aftt_different = ""
            temp_year = ""
                            pub_comp_id = 0




            spot_to_find = InStr(string_text, "href=")
            If spot_to_find > 0 Then
              string_text = Right(string_text, Len(string_text) - spot_to_find - 5)

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                pub_url = Left(Trim(string_text), spot_to_find - 12)

                pub_url = "https://www.aso.com/listings/" & pub_url
                string_text = Right(string_text, Len(string_text) - spot_to_find)
              End If
            End If


            spot_to_find = InStr(string_text, "</a>")
            If spot_to_find > 0 Then
              temp_ac_name = Left(Trim(string_text), spot_to_find - 1)

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
              End If
            End If

            spot_to_find = InStr(temp_ac_name, """")
            If spot_to_find > 0 Then
              temp_ac_name = Left(Trim(temp_ac_name), spot_to_find - 1) 
            End If

            spot_to_find = InStr(temp_ac_name, "(")
            If spot_to_find > 0 Then
              temp_ac_name = Left(Trim(temp_ac_name), spot_to_find - 1)
            End If
            cutme(temp_ac_name)



            spot_to_find = InStr(string_text, "<td  style")
            If spot_to_find > 0 Then
              string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
              End If

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
              End If

              spot_to_find = InStr(string_text, "</span>")
              If spot_to_find > 0 Then
                pub_reg_no = Left(Trim(string_text), spot_to_find - 1)
              End If
              pub_reg_no = Replace(pub_reg_no, "Reg#:", "")
              cutme(pub_reg_no)
            Else
              spot_to_find = InStr(string_text, "<td style")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

                spot_to_find = InStr(string_text, ">")
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
                End If

                spot_to_find = InStr(string_text, ">")
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
                End If

                spot_to_find = InStr(string_text, "</span>")
                If spot_to_find > 0 Then
                  pub_reg_no = Left(Trim(string_text), spot_to_find - 1)
                End If
                pub_reg_no = Replace(pub_reg_no, "Reg#:", "")
                cutme(pub_reg_no) 
              End If
            End If

            spot_to_find = InStr(string_text, "<td style")
            If spot_to_find > 0 Then
              string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
              End If

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
              End If

              spot_to_find = InStr(string_text, "</span>")
              If spot_to_find > 0 Then
                pub_ser_no = Left(Trim(string_text), spot_to_find - 1)
              End If
              pub_ser_no = Replace(pub_ser_no, "S/N:", "")
              cutme(pub_ser_no)
            End If

            spot_to_find = InStr(string_text, "<td style=""width")
            If spot_to_find > 0 Then
              string_text = Right(string_text, Len(string_text) - spot_to_find - 10)


              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
              End If

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
              End If

              spot_to_find = InStr(string_text, ">")
              spot_to_find2 = InStr(string_text, "</span>")
              ' if we are already at the end, then just use it, otherwise, go 1 more for the bold ones
              If spot_to_find2 < spot_to_find Then

              Else
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find + 1)
                End If
              End If 

              spot_to_find = InStr(string_text, "</span>")
              If spot_to_find > 0 Then
                pub_price = Left(Trim(string_text), spot_to_find - 1)
              End If
              pub_price = Replace(pub_price, ">", "") '' incase 
              pub_price = Replace(pub_price, "$", "") '' incase 
              pub_price = Replace(pub_price, ":", "") '' incase 
              cutme(pub_price)
            End If


            spot_to_find = InStr(string_text, "<td style")
            If spot_to_find > 0 Then
              string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
              End If

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
              End If

              spot_to_find = InStr(string_text, "</span>")
              If spot_to_find > 0 Then
                pub_aftt = Left(Trim(string_text), spot_to_find - 1)
              End If
              pub_aftt = Replace(pub_aftt, "TTAF:", "")
              cutme(pub_aftt)
            End If

            'skip location stuff
            spot_to_find = InStr(string_text, "<td style")
            If spot_to_find > 0 Then
              string_text = Right(string_text, Len(string_text) - spot_to_find - 10)
            End If

            spot_to_find = InStr(string_text, "<td style")
            If spot_to_find > 0 Then
              string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
              End If

              spot_to_find = InStr(string_text, ">")
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find)
              End If

              spot_to_find = InStr(string_text, "</span>")
              If spot_to_find > 0 Then
                pub_seller_info = Left(Trim(string_text), spot_to_find - 1)
              End If
              pub_seller_info = Replace(pub_seller_info, "'", "''")
              cutme(pub_seller_info)
            End If


            If IsNumeric(Left(Trim(temp_ac_name), "4")) Then
              temp_year = Left(Trim(temp_ac_name), "4")
            Else
              temp_year = ""
            End If

                        acpub_original_name = temp_ac_name & " " & pub_ser_no

                        ' added in MSW -
                        If Trim(pub_url) <> "" Then
                            pub_url = Replace(pub_url, "https://www.aso.com/listings/https://www.aso", "https://www.aso")
                        End If



                        Response.Write("<Br>")
            Response.Write("<Br>" & temp_ac_name)
            Response.Write("<Br>" & pub_price)
            Response.Write("<Br>" & pub_url)
            Response.Write("<Br>" & pub_seller_info)
            'Response.Write("<Br>" & pub_desc)
            Response.Write("<Br>" & temp_year)
            Response.Write("<Br>" & pub_ser_no)
            Response.Write("<Br>" & pub_reg_no)
            Response.Write("<Br>" & pub_aftt)

            If Trim(temp_ac_name) <> "" Then

              array_split_make = Split(Trim(temp_ac_name), " ")

              If array_split_make.Length = 2 Then
                temp_make = array_split_make(0)
                temp_model = array_split_make(1)
              ElseIf array_split_make.Length = 3 Then
                temp_make = array_split_make(1)
                temp_model = array_split_make(2)
              ElseIf array_split_make.Length = 4 Then
                temp_make = array_split_make(2)
                temp_model = array_split_make(3)
              ElseIf array_split_make.Length = 5 Then
                temp_make = array_split_make(3)
                temp_model = array_split_make(4)
              Else
                temp_temp = ""
              End If



              temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
                            If temp_ac_id = 0 Then
                                temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                                If temp_ac_id = 0 Then
                                    temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                                    If temp_ac_id = 0 Then

                                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                                        If temp_ac_id = 0 Then
                                            temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                                            If temp_ac_id = 0 Then
                                                temp_ac_id = temp_ac_id
                                            End If
                                        End If

                                    End If
                                End If

                            End If


                            acpub_price_details = ""
                            If On_Naughty_List(temp_ac_name) = True Then
                                ' if its on naughtly list then exclude 
                                temp_ac_id = temp_ac_id
                            Else
                                If temp_ac_id > 0 Then
                                    Call find_ac_data(temp_ac_id)
                                Else
                                    acpub_process_status = "For Sale Not Found – No AC Match"
                  acpub_status = "O"
                End If

                If Trim(aftt_different) <> "" Then
                  pub_desc = pub_desc & " " & aftt_different
                End If

                If Trim(acpub_price_details) <> "" Then
                  pub_desc = pub_desc & " " & acpub_price_details
                End If

                Call check_insert_ac_pub(temp_ac_id, 1)
              End If

              Response.Write("<Br>AC ID:" & temp_ac_id)
            End If
          End If


          If tcount = 2 Then
            tcount = 0
          End If

        Next


      End If




    Catch ex As Exception
    Finally

    End Try

  End Function
    Public Function scrape_for_GlobalAIR(ByVal link As String)
        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader


        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim i As Integer = 0
        Dim final_string As String = ""
        Dim original_string_text As String = ""
        Dim article_link As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim k As Integer = 0
        Dim skip_this As Boolean = False
        Dim extra_note As String = ""

        Dim temp_ac_name As String = ""
        Dim temp_engine As String = ""
        Dim temp_eng As String = ""
        Dim temp_av As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_make As String = ""
        Dim temp_temp As String
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split_make() As String
        Dim tcount As Integer = 0




        Try
            System.Net.ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType)

            System.Threading.Thread.Sleep(10)
            Response.Flush()
            System.Threading.Thread.Sleep(10)

            Dim req As System.Net.WebRequest

            req = System.Net.WebRequest.Create(link)


            Dim resp As System.Net.WebResponse = req.GetResponse


            Str = resp.GetResponseStream
            srRead = New System.IO.StreamReader(Str)
            ' read all the text 
            string_text = srRead.ReadToEnd().ToString

            resp.Close()
            resp = Nothing
            req = Nothing

            string_text = string_text
            original_string_text = string_text


            spot_to_find = InStr(string_text, "result-container")
            If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find + 5)

                array_split = Split(string_text, "result-container")

                For i = 1 To array_split.Length - 1
                    tcount = tcount + 1  ' do every third one 
                    string_text = array_split(i)
                    original_string_text = string_text

                    acpub_count = acpub_count + 1

                    temp_ac_name = ""
                    temp_engine = ""
                    temp_eng = ""
                    temp_av = ""

                    pub_reg_no = ""
                    pub_ser_no = ""
                    pub_desc = ""
                    pub_price = ""
                    pub_aftt = ""
                    pub_seller_info = ""
                    pub_picture = ""
                    pub_status = ""
                    pub_url = ""
                    has_pics = False
                    aftt_different = ""
                    temp_year = ""
                    pub_comp_id = 0

                    spot_to_find = InStr(string_text, " href=")
                    If spot_to_find > 0 Then
                        pub_url = Right(string_text, Len(string_text) - spot_to_find - 6)
                        spot_to_find = InStr(pub_url, ">")
                        If spot_to_find > 0 Then
                            pub_url = Left(Trim(pub_url), spot_to_find - 1)
                            pub_url = Replace(pub_url, """", "")


                            If InStr(pub_url, "w.globalair.com") = 0 Then
                                pub_url = "https://www.globalair.com" & pub_url
                            End If

                            If InStr(pub_url, "email-protection") > 0 Then
                                pub_url = pub_url
                            End If


                        End If
                    End If




                    spot_to_find = InStr(string_text, "<div>TT: <span>")
                    If spot_to_find > 0 Then
                        pub_aftt = Right(string_text, Len(string_text) - spot_to_find - 14)

                        spot_to_find = InStr(pub_aftt, "</span></div>")
                        If spot_to_find > 0 Then
                            pub_aftt = Left(Trim(pub_aftt), spot_to_find - 1)
                        End If

                        pub_aftt = Replace(pub_aftt, " hrs", "")
                        pub_aftt = Replace(pub_aftt, "hrs", "")
                    End If

                    spot_to_find = InStr(string_text, "<div>SN: <span>")
                    If spot_to_find > 0 Then
                        pub_ser_no = Right(string_text, Len(string_text) - spot_to_find - 14)

                        spot_to_find = InStr(pub_ser_no, "</span></div>")
                        If spot_to_find > 0 Then
                            pub_ser_no = Left(Trim(pub_ser_no), spot_to_find - 1)
                        End If
                    End If


                    If Trim(pub_ser_no) = "" Then
                        spot_to_find = InStr(string_text, "<div>SN:")
                        If spot_to_find > 0 Then
                            pub_ser_no = Right(string_text, Len(string_text) - spot_to_find - 8)

                            spot_to_find = InStr(pub_ser_no, "</div>")
                            If spot_to_find > 0 Then
                                pub_ser_no = Left(Trim(pub_ser_no), spot_to_find - 1)
                            End If
                        End If
                    End If

                    spot_to_find = InStr(string_text, "<div>RN: <span>")
                    If spot_to_find > 0 Then
                        pub_reg_no = Right(string_text, Len(string_text) - spot_to_find - 14)

                        spot_to_find = InStr(pub_reg_no, "</span></div>")
                        If spot_to_find > 0 Then
                            pub_reg_no = Left(Trim(pub_reg_no), spot_to_find - 1)
                        End If

                        pub_reg_no = Replace(pub_reg_no, "<br/>", "")
                        pub_reg_no = Replace(pub_reg_no, "<br>", "")
                    End If


                    If Trim(pub_reg_no) = "" Then
                        spot_to_find = InStr(string_text, "<div>RN:")
                        If spot_to_find > 0 Then
                            pub_reg_no = Right(string_text, Len(string_text) - spot_to_find - 8)

                            spot_to_find = InStr(pub_reg_no, "</div>")
                            If spot_to_find > 0 Then
                                pub_reg_no = Left(Trim(pub_reg_no), spot_to_find - 1)
                            End If

                            pub_reg_no = Replace(pub_reg_no, "<br/>", "")
                            pub_reg_no = Replace(pub_reg_no, "<br>", "")
                        End If
                    End If

                    spot_to_find = InStr(string_text, "Price: <span>")
                    If spot_to_find > 0 Then
                        pub_price = Right(string_text, Len(string_text) - spot_to_find - 12)

                        spot_to_find = InStr(pub_price, "</span></div>")
                        If spot_to_find > 0 Then
                            pub_price = Left(Trim(pub_price), spot_to_find - 1)
                        End If
                    End If



                    temp_year = ""
                    temp_ac_name = ""
                    spot_to_find = InStr(string_text, "result-title")
                    If spot_to_find > 0 Then
                        temp_ac_name = Right(string_text, Len(string_text) - spot_to_find - 13)

                        spot_to_find = InStr(temp_ac_name, "</a></div>")
                        If spot_to_find > 0 Then
                            temp_ac_name = Left(Trim(temp_ac_name), spot_to_find - 1)
                        End If

                        spot_to_find = InStr(temp_ac_name, ">")
                        If spot_to_find > 0 Then
                            temp_ac_name = Right(temp_ac_name, Len(temp_ac_name) - spot_to_find)
                        End If

                        ' added MSW - 7/20/2020 
                        spot_to_find = InStr(temp_ac_name, "</a>")
                        If spot_to_find > 0 Then
                            temp_ac_name = Left(Trim(temp_ac_name), spot_to_find - 1)
                        End If

                        If IsNumeric(Left(Trim(temp_ac_name), 4)) = True Then
                            temp_year = Left(Trim(temp_ac_name), 4)
                        End If


                        spot_to_find = InStr(Trim(temp_ac_name), ">")
                        If spot_to_find > 0 Then
                            temp_ac_name = Right(temp_ac_name, Len(Trim(temp_ac_name)) - spot_to_find)
                        End If

                        temp_ac_name = Replace(temp_ac_name, "</h3>", "")
                    End If

                    If InStr(temp_ac_name, "013 Bell 412-EPi") > 0 Then
                        temp_ac_name = temp_ac_name
                    End If


                    spot_to_find = InStr(string_text, "result-seller")
                    If spot_to_find > 0 Then
                        pub_seller_info = Right(string_text, Len(string_text) - spot_to_find - 15)

                        spot_to_find = InStr(pub_seller_info, "</a></div>")
                        If spot_to_find > 0 Then
                            pub_seller_info = Left(Trim(pub_seller_info), spot_to_find - 1)
                        End If

                        spot_to_find = InStr(pub_seller_info, ">")
                        If spot_to_find > 0 Then
                            pub_seller_info = Right(pub_seller_info, Len(pub_seller_info) - spot_to_find)
                        End If

                        pub_seller_info = Replace(pub_seller_info, "<Br>", "")
                        pub_seller_info = Replace(pub_seller_info, "<Br/>", "")

                        ' added MSW - 7/20/2020 
                        spot_to_find = InStr(pub_seller_info, "</a>")
                        If spot_to_find > 0 Then
                            pub_seller_info = Left(Trim(pub_seller_info), spot_to_find - 1)
                        End If
                    End If







                    If pub_comp_id = 0 Then
                        If Not IsNothing(pub_seller_info) Then
                            If Trim(pub_seller_info) <> "" Then
                                'if there is a line feed or break, then try to get just the company name 
                                If InStr(Trim(pub_seller_info), Asc(10)) > 0 Then
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), InStr(Trim(pub_seller_info), Asc(10)) - 1))
                                ElseIf InStr(Trim(pub_seller_info), Asc(13)) > 0 Then
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), InStr(Trim(pub_seller_info), Asc(13)) - 1))
                                Else
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), 17))

                                    If pub_comp_id = 0 Then
                                        pub_comp_id = find_comp_id_global_search(Trim(pub_seller_info))
                                    End If
                                End If
                            End If
                        End If
                    End If

                    If pub_comp_id = 0 And Trim(pub_seller_info <> "") Then
                        pub_comp_id = find_comp_id_previous_pub(pub_seller_info, 8)
                    End If



                    Response.Write("<Br>")
                    Response.Write("<Br>" & temp_ac_name)
                    Response.Write("<Br>" & pub_price)
                    Response.Write("<Br>" & pub_url)
                    Response.Write("<Br>" & pub_seller_info)
                    'Response.Write("<Br>" & pub_desc)
                    Response.Write("<Br>" & temp_year)
                    Response.Write("<Br>" & pub_ser_no)
                    Response.Write("<Br>" & pub_reg_no)
                    Response.Write("<Br>" & pub_aftt)


                    acpub_original_name = temp_ac_name & " " & pub_ser_no

                    If Trim(pub_ser_no) = "" And Trim(pub_reg_no) = "" Then
                        pub_ser_no = pub_ser_no
                    End If


                    If Trim(temp_ac_name) <> "" Then

                        array_split_make = Split(Trim(temp_ac_name), " ")

                        If array_split_make.Length = 2 Then
                            temp_make = array_split_make(0)
                            temp_model = array_split_make(1)
                        ElseIf array_split_make.Length = 3 Then
                            temp_make = array_split_make(1)
                            temp_model = array_split_make(2)
                        ElseIf array_split_make.Length = 4 Then
                            temp_make = array_split_make(2)
                            temp_model = array_split_make(3)
                        ElseIf array_split_make.Length = 5 Then
                            temp_make = array_split_make(3)
                            temp_model = array_split_make(4)
                        Else
                            temp_temp = ""
                        End If



                        pub_url = Replace(Trim(pub_url), "TMAKE", Trim(temp_make))
                        pub_url = Replace(Trim(pub_url), "TMODEL", Trim(temp_model))


                        temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
                        If temp_ac_id = 0 Then
                            temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                            If temp_ac_id = 0 Then
                                temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                                If temp_ac_id = 0 Then

                                    temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                                    If temp_ac_id = 0 Then
                                        temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                                        If temp_ac_id = 0 Then
                                            temp_ac_id = temp_ac_id
                                        End If
                                    End If

                                End If
                            End If

                        End If

                        ' added in MSW then we didnt find it - re check with an N reg number 
                        If temp_ac_id = 0 Then
                            If Trim(pub_reg_no) <> "" Then
                                If Left(Trim(pub_reg_no), 1) <> "N" Then
                                    temp_ac_id = find_ac_global_search("", "", "", "N" & pub_reg_no)
                                    If temp_ac_id = 0 Then
                                        temp_ac_id = temp_ac_id
                                    End If
                                End If
                            End If
                        End If



                            pub_desc = ""
                        acpub_price_details = ""
                        acpub_process_status = ""
                        acpub_status = ""
                        If On_Naughty_List(temp_ac_name) = True Then
                            ' if its on naughtly list then exclude 
                            temp_ac_id = temp_ac_id
                        Else
                            If temp_ac_id > 0 Then
                                Call find_ac_data(temp_ac_id)
                            Else
                                acpub_process_status = "For Sale Not Found – No AC Match"
                                acpub_status = "O"
                            End If

                            If Trim(aftt_different) <> "" Then
                                pub_desc = pub_desc & aftt_different
                            End If


                            If Trim(acpub_price_details) <> "" Then
                                If Trim(aftt_different) <> "" Then
                                    pub_desc = pub_desc & ", "
                                End If
                                pub_desc = pub_desc & " " & acpub_price_details
                            End If

                            ' find bad ? 
                            If temp_ac_id = 0 And InStr(acpub_process_status, "Found") > 0 Then
                                temp_ac_id = temp_ac_id
                            ElseIf temp_ac_id = 0 Then
                                temp_ac_id = temp_ac_id
                            End If

                            Call check_insert_ac_pub(temp_ac_id, 8)
                        End If

                        Response.Write("<Br>AC ID:" & temp_ac_id)
                    End If


                    'If tcount = 2 Then
                    '  tcount = 0
                    'End If

                Next


            End If




        Catch ex As Exception
        Finally

        End Try

    End Function
    Public Function scrape_for_TradeAPlane(ByVal page_no As Integer)
        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader



        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim i As Integer = 0
        Dim final_string As String = ""
        Dim original_string_text As String = ""
        Dim article_link As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim k As Integer = 0
        Dim skip_this As Boolean = False
        Dim extra_note As String = ""

        Dim temp_ac_name As String = ""
        Dim temp_engine As String = ""
        Dim temp_eng As String = ""
        Dim temp_av As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_make As String = ""
        Dim temp_temp As String
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split_make() As String
        Dim tcount As Integer = 0




        Try


            System.Threading.Thread.Sleep(10)
            Response.Flush()
            System.Threading.Thread.Sleep(10)

            Dim req As System.Net.WebRequest

            'req = System.Net.WebRequest.Create("https://www.trade-a-plane.com/search?s-type=aircraft&s-advanced=yes&s-custom=&sale_status=For+Sale&category_level1=Jets&days_old-max=10/")
            'req = System.Net.WebRequest.Create("https://www.trade-a-plane.com/search?category_level1=Jets&s-type=aircraft&s-sort_order=asc&s-sort_key=days_since_update")
            'req = System.Net.WebRequest.Create("https://www.trade-a-plane.com/filtered/search?category_level1=Jets&s-type=aircraft&s-custom_style=oneline")

            If page_no > 0 Then
                req = System.Net.WebRequest.Create("https://www.trade-a-plane.com/search?category_level1=Jets&category_level1=Turboprop&category_level1=Turbine+Helicopters&s-type=aircraft&s-page_size=96&s-sort_key=days_since_update&s-sort_order=asc&s-page=2")
            Else
                req = System.Net.WebRequest.Create("https://www.trade-a-plane.com/search?category_level1=Jets&category_level1=Turboprop&category_level1=Turbine+Helicopters&s-type=aircraft&s-page_size=96&s-sort_key=days_since_update&s-sort_order=asc")
            End If
            ''

            '   req = System.Net.WebRequest.Create("https://www.trade-a-plane.com/")
            ' System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls

            Dim resp As System.Net.WebResponse = req.GetResponse


            Str = resp.GetResponseStream
            srRead = New System.IO.StreamReader(Str)
            ' read all the text 
            string_text = srRead.ReadToEnd().ToString

            resp.Close()
            resp = Nothing
            req = Nothing

            string_text = string_text
            original_string_text = string_text


            spot_to_find = InStr(string_text, "result_listing ")
            If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 15)

                array_split = Split(string_text, "result_listing ")

                For i = 0 To array_split.Length - 1
                    tcount = tcount + 1  ' do every third one 
                    string_text = array_split(i)
                    original_string_text = string_text

                    acpub_count = acpub_count + 1

                    temp_ac_name = ""
                    temp_engine = ""
                    temp_eng = ""
                    temp_av = ""

                    pub_reg_no = ""
                    pub_ser_no = ""
                    pub_desc = ""
                    pub_price = ""
                    pub_aftt = ""
                    pub_seller_info = ""
                    pub_picture = ""
                    pub_status = ""
                    pub_url = ""
                    has_pics = False
                    aftt_different = ""
                    temp_year = ""
                    pub_comp_id = 0



                    spot_to_find = InStr(string_text, "a href=")
                    If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 7)

                        spot_to_find = InStr(string_text, "class='")
                        If spot_to_find > 0 Then
                            pub_url = Left(Trim(string_text), spot_to_find - 3)

                            If InStr(pub_url, "http") > 0 Then

                            Else
                                pub_url = "https://www.trade-a-plane.com/" & pub_url
                            End If



                            spot_to_find = InStr(pub_url, " id=""title")
                            If spot_to_find > 0 Then
                                pub_url = Left(Trim(pub_url), spot_to_find - 2)
                            End If '

                        End If
                    End If


                    spot_to_find = InStr(string_text, "</a>")
                    If spot_to_find > 0 Then
                        temp_ac_name = Left(Trim(string_text), spot_to_find - 1)

                        spot_to_find = InStr(temp_ac_name, ">")
                        If spot_to_find > 0 Then
                            temp_ac_name = Right(temp_ac_name, Len(temp_ac_name) - spot_to_find - 1)
                        End If

                        temp_ac_name = Replace(temp_ac_name, "<!-- (For Sale) -->", "")
                        temp_ac_name = Replace(temp_ac_name, "<!-- (Wanted) -->", "")

                        If InStr(temp_ac_name, "<!--  -->") > 0 Then
                            temp_ac_name = temp_ac_name
                        End If

                        temp_ac_name = Replace(temp_ac_name, "<!--  -->", "")


                        spot_to_find = InStr(temp_ac_name, ">")
                        If spot_to_find > 0 Then
                            temp_ac_name = Right(temp_ac_name, Len(temp_ac_name) - spot_to_find - 1)
                        End If
                        cutme(temp_ac_name)
                        'moved to the right 


                    End If

                    If Trim(temp_ac_name) = "" Then
                        temp_ac_name = temp_ac_name
                    End If



                    spot_to_find = InStr(string_text, "txt-price")
                    If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 11)
                        spot_to_find = InStr(string_text, "</div>")
                        If spot_to_find > 0 Then
                            pub_price = Left(Trim(string_text), spot_to_find - 1)
                        End If
                        pub_price = Replace(pub_price, "<span class='callforprice'>", "")
                        pub_price = Replace(pub_price, "</span>", "")
                        cutme(pub_price)
                    End If


                    spot_to_find = InStr(string_text, "<span>Reg#</span>")
                    If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 17)

                        spot_to_find = InStr(string_text, "</div>")
                        If spot_to_find > 0 Then
                            pub_reg_no = Left(Trim(string_text), spot_to_find - 1)
                            pub_reg_no = Replace(pub_reg_no, "<b>", "")
                            pub_reg_no = Replace(pub_reg_no, "Not Listed", "")
                            cutme(pub_reg_no)
                        End If
                    End If

                    spot_to_find = InStr(string_text, "<span>TT: </span>")
                    If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 17)

                        spot_to_find = InStr(string_text, "</div>")
                        If spot_to_find > 0 Then
                            pub_aftt = Left(Trim(string_text), spot_to_find - 2)
                            pub_aftt = Replace(pub_aftt, "</li>", "")
                            cutme(pub_aftt)
                        End If
                    End If


                    '<a href="/search?seller_id=11598&s-type=aircraft" title="David Yonak" itemprop="name">David Yonak</a>
                    ' 	<a href = "/search?seller_id=89027&s-type=aircraft" title="International Aircraft Marketing & Sales - Celia Perkins" itemprop="name">International Aircraft Marketing & Sales - Celia Perkins</a>
                    '	<a href="/search?seller_id=45584&s-type=aircraft" title="Courtesy Aircraft Sales" itemprop="name">Courtesy Aircraft Sales</a>

                    '
                    '/search?seller_id=34896&s-type=aircraft"

                    ' FOR SALE BY 
                    'spot_to_find = InStr(string_text, "/company-search/details?seller_id")
                    spot_to_find = InStr(string_text, "/search?seller_id=")
                    If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 6)

                        spot_to_find = InStr(string_text, ">")
                        If spot_to_find > 0 Then
                            string_text = Right(string_text, Len(string_text) - spot_to_find)
                        End If

                        spot_to_find = InStr(string_text, "</a>")
                        If spot_to_find > 0 Then
                            pub_seller_info = Left(Trim(string_text), spot_to_find - 1)
                            pub_seller_info = Replace(pub_seller_info, "<b>", "")
                            cutme(pub_seller_info)
                        End If
                    End If

                    string_text = string_text


                    acpub_original_name = temp_ac_name & " " & pub_ser_no

                    'spot_to_find = InStr(string_text, "prod_desc")
                    'If spot_to_find > 0 Then
                    '  string_text = Right(string_text, Len(string_text) - spot_to_find - 10)
                    '  spot_to_find = InStr(Trim(string_text), "<a href")
                    '  If spot_to_find > 0 Then
                    '    pub_desc = Left(Trim(string_text), spot_to_find - 2)
                    '    pub_desc = Replace(pub_desc, "'", "''")
                    '  End If
                    'End If


                    'spot_to_find = InStr(string_text, "Year")
                    'If spot_to_find > 0 Then
                    '  string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

                    '  spot_to_find = InStr(string_text, "</b>")
                    '  If spot_to_find > 0 Then
                    '    temp_year = Left(Trim(string_text), spot_to_find - 2)
                    '    temp_year = Replace(temp_year, "<b>", "")
                    '  End If
                    'End If


                    'spot_to_find = InStr(string_text, "S/N")
                    'If spot_to_find > 0 Then
                    '  string_text = Right(string_text, Len(string_text) - spot_to_find - 2)

                    '  spot_to_find = InStr(string_text, "</b>")
                    '  If spot_to_find > 0 Then
                    '    pub_ser_no = Left(Trim(string_text), spot_to_find - 2)
                    '    pub_ser_no = Replace(pub_ser_no, "<b>", "")
                    '  End If
                    'End If


                    'spot_to_find = InStr(string_text, "TTAF")
                    'If spot_to_find > 0 Then
                    '  string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

                    '  spot_to_find = InStr(string_text, "</b>")
                    '  If spot_to_find > 0 Then
                    '    pub_aftt = Left(Trim(string_text), spot_to_find - 2)
                    '    pub_aftt = Replace(pub_aftt, "<b>", "")
                    '    pub_aftt = Replace(pub_aftt, "'", "")
                    '  End If
                    'End If



                    'Response.Write("<Br>")
                    'Response.Write("<Br>" & temp_ac_name)
                    'Response.Write("<Br>" & pub_price)
                    'Response.Write("<Br>" & pub_url)
                    'Response.Write("<Br>" & pub_seller_info)
                    'Response.Write("<Br>" & pub_desc)
                    'Response.Write("<Br>" & temp_year)
                    'Response.Write("<Br>" & pub_ser_no)
                    'Response.Write("<Br>" & pub_aftt)



                    array_split_make = Split(Trim(temp_ac_name), " ")

                    If array_split_make.Length > 4 Then
                        temp_make = array_split_make(2)
                        temp_model = array_split_make(3)
                    ElseIf array_split_make.Length = 4 Then
                        temp_make = array_split_make(2)
                        temp_model = array_split_make(3)
                    ElseIf array_split_make.Length = 3 Then
                        temp_make = array_split_make(1)
                        temp_model = array_split_make(2)
                    Else
                        temp_make = temp_make
                    End If



                    temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
                    If temp_ac_id = 0 Then
                        temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                        If temp_ac_id = 0 Then
                            temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                            If temp_ac_id = 0 Then

                                temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                                If temp_ac_id = 0 Then
                                    temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                                    If temp_ac_id = 0 Then
                                        temp_ac_id = temp_ac_id


                                        ' added in MSW - from controller 
                                        If temp_ac_id = 0 Then
                                            If Trim(pub_ser_no) <> "" Then
                                                temp_ac_id = find_ac_ac_search(pub_ser_no, temp_make, temp_model, "")
                                            End If

                                            If temp_ac_id = 0 Then
                                                If Trim(pub_reg_no) <> "" Then
                                                    temp_ac_id = find_ac_ac_search("", "", "", pub_reg_no) ' just search reg- look for 1 
                                                End If

                                                If temp_ac_id = 0 Then
                                                    If Trim(pub_ser_no) <> "" And Left(Trim(pub_ser_no), 1) = "0" Then
                                                        temp_ac_id = find_ac_ac_search(Right(Trim(pub_ser_no), Len(Trim(pub_ser_no)) - 1), temp_make, temp_model, "")
                                                    End If
                                                End If

                                            End If
                                        End If
                                        temp_ac_id = temp_ac_id




                                    End If
                                End If

                            End If
                        End If

                    End If


                    If pub_comp_id = 0 Then
                        If Not IsNothing(pub_seller_info) Then
                            If Trim(pub_seller_info) <> "" Then
                                'if there is a line feed or break, then try to get just the company name 
                                If InStr(Trim(pub_seller_info), Asc(10)) > 0 Then
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), InStr(Trim(pub_seller_info), Asc(10)) - 1))
                                ElseIf InStr(Trim(pub_seller_info), Asc(13)) > 0 Then
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), InStr(Trim(pub_seller_info), Asc(13)) - 1))
                                Else
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), 17))
                                End If
                            End If
                        End If
                    End If

                    If pub_comp_id = 0 And Trim(pub_seller_info <> "") Then
                        pub_comp_id = find_comp_id_previous_pub(pub_seller_info, 7)
                    End If



                    acpub_status = "N"
                    If On_Naughty_List(temp_ac_name) = True Then
                        ' if its on naughtly list then excldue  
                    Else

                        If temp_ac_id > 0 Then
                            Call find_ac_data(temp_ac_id)

                            If Trim(aftt_different) <> "" Then
                                pub_desc = pub_desc & aftt_different
                            End If

                            If Trim(acpub_price_details) <> "" Then
                                pub_desc = pub_desc & " " & acpub_price_details
                            End If
                        Else
                            acpub_process_status = "For Sale Not Found – No AC Match"
                            acpub_status = "O"

                            If Trim(pub_aftt) <> "" Then
                                pub_desc = "Pub AFTT: " & pub_aftt
                            End If

                            If Trim(pub_price) <> "" Then
                                pub_desc = "Pub Price: " & pub_price
                            End If
                        End If



                        If Trim(aftt_different) <> "" Then
                            pub_desc = pub_desc & aftt_different
                        End If


                        ' If temp_ac_id > 0 Then
                        Call check_insert_ac_pub(temp_ac_id, 3)
                        '  Else
                        '  temp_ac_id = temp_ac_id
                        '  pub_reg_no = pub_reg_no
                        '  temp_make = temp_make
                        '  temp_model = temp_model
                        ' End If
                    End If

                    '   Response.Write("<Br>AC ID:" & temp_ac_id)



                Next


            End If




        Catch ex As Exception
        Finally

        End Try

    End Function

    Public Function scrape_for_AvBuyer(ByVal page_num As Long)



        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim i As Integer = 0
        Dim final_string As String = ""
        Dim original_string_text As String = ""
        Dim article_link As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim spot_to_find3 As Integer = 0
        Dim k As Integer = 0
        Dim skip_this As Boolean = False
        Dim extra_note As String = ""

        Dim temp_ac_name As String = ""
        Dim temp_engine As String = ""
        Dim temp_eng As String = ""
        Dim temp_av As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_make As String = ""
        Dim temp_temp As String
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split_make() As String
        Dim tcount As Integer = 0
        Dim results_table As New DataTable
        Dim temp_pub_id As Long = 0
        Dim Update_Query As String = ""



        Try


            System.Threading.Thread.Sleep(10)
            Response.Flush()
            System.Threading.Thread.Sleep(10)



            tcount = tcount + 1  ' do every third one 

            results_table = find_scraped_aircraft()

            If Not IsNothing(results_table) Then
                If results_table.Rows.Count > 0 Then
                    For Each r As DataRow In results_table.Rows


                        acpub_count = acpub_count + 1

                        temp_ac_id = 0
                        temp_ac_name = ""
                        temp_engine = ""
                        temp_eng = ""
                        temp_av = ""

                        pub_reg_no = ""
                        pub_ser_no = ""
                        pub_desc = ""
                        pub_price = ""
                        pub_aftt = ""
                        pub_seller_info = ""
                        pub_picture = ""
                        pub_status = ""
                        pub_url = ""
                        has_pics = False
                        aftt_different = ""
                        temp_year = ""
                        temp_make = ""
                        temp_model = ""
                        pub_comp_id = 0
                        temp_pub_id = 0
                        Update_Query = ""


                        ' scrp_ac_id, scrp_ac_source, scrp_ac_model, scrp_ac_asking_price , scrp_ac_detail_link, "
                        '     select_query &= " scrp_ac_location , scrp_ac_dealer, scrp_ac_year, scrp_ac_ser_no , scrp_ac_airframe_tot_hrs , scrp_ac_note



                        If Not IsDBNull(r.Item("scrp_ac_id")) Then
                            temp_pub_id = Trim(r.Item("scrp_ac_id"))
                        End If

                        If Not IsDBNull(r.Item("scrp_ac_model")) Then
                            temp_ac_name = Trim(r.Item("scrp_ac_model"))
                        End If

                        If Not IsDBNull(r.Item("scrp_ac_asking_price")) Then
                            pub_price = Trim(r.Item("scrp_ac_asking_price"))

                            pub_price = Replace(pub_price, "Price: USD ", "")
                            pub_price = Replace(pub_price, "Price Reduced", "")
                            pub_price = Replace(pub_price, "Fractional Ownership", "")
                            pub_price = Replace(pub_price, "Price: €", "")
                        End If

                        If Not IsDBNull(r.Item("scrp_ac_detail_link")) Then
                            pub_url = Trim(r.Item("scrp_ac_detail_link"))
                        End If

                        'If Not IsDBNull(r.Item("scrp_ac_location")) Then
                        '    temp_ac_name = Trim(r.Item("scrp_ac_location"))
                        'End If

                        If Not IsDBNull(r.Item("scrp_ac_dealer")) Then
                            pub_seller_info = Trim(r.Item("scrp_ac_dealer"))
                        End If

                        If Not IsDBNull(r.Item("scrp_ac_year")) Then
                            temp_year = Trim(r.Item("scrp_ac_year"))
                        End If

                        If Not IsDBNull(r.Item("scrp_ac_ser_no")) Then
                            pub_ser_no = Trim(r.Item("scrp_ac_ser_no"))
                        End If

                        If Not IsDBNull(r.Item("scrp_ac_airframe_tot_hrs")) Then
                            pub_aftt = Trim(r.Item("scrp_ac_airframe_tot_hrs"))
                        End If

                        If Not IsDBNull(r.Item("scrp_ac_note")) Then
                            pub_desc = Trim(r.Item("scrp_ac_note"))
                        End If



                        acpub_original_name = temp_ac_name & " " & pub_ser_no

                        cutme(acpub_original_name)
                        cutme(temp_ac_name)
                        cutme(pub_price)
                        cutme(pub_url)
                        cutme(pub_seller_info)
                        cutme(pub_desc)
                        cutme(temp_year)
                        cutme(pub_ser_no)
                        cutme(pub_aftt)

                        pub_aftt = Replace(pub_aftt, "Hours", "")


                        Response.Write("<Br>")
                        Response.Write("<Br>" & temp_ac_name)
                        Response.Write("<Br>" & pub_price)
                        Response.Write("<Br>" & pub_url)
                        Response.Write("<Br>" & pub_seller_info)
                        Response.Write("<Br>" & pub_desc)
                        Response.Write("<Br>" & temp_year)
                        Response.Write("<Br>" & pub_ser_no)
                        Response.Write("<Br>" & pub_aftt)



                        array_split_make = Split(Trim(temp_ac_name), " ")

                        If array_split_make.Length = 2 Then
                            temp_make = array_split_make(0)
                            temp_model = array_split_make(1)
                        ElseIf array_split_make.Length = 3 Then
                            temp_make = array_split_make(1)
                            temp_model = array_split_make(2)
                        ElseIf array_split_make.Length = 4 Then
                            temp_make = array_split_make(2)
                            temp_model = array_split_make(3)
                        ElseIf array_split_make.Length = 5 Then
                            temp_make = array_split_make(3)
                            temp_model = array_split_make(4)
                        Else
                            temp_temp = ""
                        End If

                        If InStr(Trim(pub_ser_no), "&dash;") > 0 Then
                            pub_ser_no = ""
                        End If

                        temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
                        If temp_ac_id = 0 Then
                            temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                            If temp_ac_id = 0 Then
                                temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                                If temp_ac_id = 0 Then

                                    temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                                    If temp_ac_id = 0 Then
                                        temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                                        If temp_ac_id = 0 Then
                                            temp_ac_id = temp_ac_id
                                        End If
                                    End If

                                End If
                            End If

                        End If



                        If pub_comp_id = 0 Then
                            If Not IsNothing(pub_seller_info) Then
                                If Trim(pub_seller_info) <> "" Then
                                    'if there is a line feed or break, then try to get just the company name 
                                    If InStr(Trim(pub_seller_info), Asc(10)) > 0 Then
                                        pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), InStr(Trim(pub_seller_info), Asc(10)) - 1))
                                    ElseIf InStr(Trim(pub_seller_info), Asc(13)) > 0 Then
                                        pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), InStr(Trim(pub_seller_info), Asc(13)) - 1))
                                    Else
                                        pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), 17))
                                    End If
                                End If
                            End If
                        End If

                        If pub_comp_id = 0 And Trim(pub_seller_info <> "") Then
                            pub_comp_id = find_comp_id_previous_pub(pub_seller_info, 2)
                        End If


                        acpub_price_details = ""
                        If On_Naughty_List(temp_ac_name) = True Then
                            ' if its on naughtly list then excldue  
                        Else
                            If temp_ac_id > 0 Then
                                Call find_ac_data(temp_ac_id)
                            Else
                                acpub_process_status = "For Sale Not Found – No AC Match"
                                acpub_status = "O"
                            End If

                            If Trim(aftt_different) <> "" Then
                                pub_desc = pub_desc & aftt_different
                            End If

                            If Trim(acpub_price_details) <> "" Then
                                pub_desc = pub_desc & " " & acpub_price_details
                            End If

                            temp_ac_id = temp_ac_id
                            Call check_insert_ac_pub(temp_ac_id, 7)
                        End If

                        Response.Write("<Br>AC ID:" & temp_ac_id)



                        'temp_pub_id

                        Update_Query = " Update scraped_aircraft set scrp_ac_processed = 'Y' where scrp_ac_id = " & temp_pub_id & "  "

                        MySqlCommand_JETNET.CommandText = Update_Query
                        MySqlCommand_JETNET.ExecuteNonQuery()

                        System.Threading.Thread.Sleep(200)


                    Next
                Else
                End If
            End If



        Catch ex As Exception
        Finally

        End Try

    End Function

    Public Function scrape_for_AvBuyer_12_22_21(ByVal page_num As Long)
        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader



        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim i As Integer = 0
        Dim final_string As String = ""
        Dim original_string_text As String = ""
        Dim article_link As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim spot_to_find3 As Integer = 0
        Dim array_split() As String
        Dim k As Integer = 0
        Dim skip_this As Boolean = False
        Dim extra_note As String = ""

        Dim temp_ac_name As String = ""
        Dim temp_engine As String = ""
        Dim temp_eng As String = ""
        Dim temp_av As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_make As String = ""
        Dim temp_temp As String
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split_make() As String
        Dim tcount As Integer = 0




        Try


            System.Threading.Thread.Sleep(10)
            Response.Flush()
            System.Threading.Thread.Sleep(10)

            Dim req As System.Net.WebRequest


            If page_num > 1 Then
                req = System.Net.WebRequest.Create("https://www.avbuyer.com/aircraft/page-" & page_num & "")
            Else
                req = System.Net.WebRequest.Create("https://www.avbuyer.com/aircraft?page=" & page_num & "&rows=25")
            End If



            ' req = System.Net.WebRequest.Create("https://www.avbuyer.com/aircraft/private-jets?page=" & page_num & "")
            'changed from private jets to all 
            req.UseDefaultCredentials = True

            Dim resp As System.Net.WebResponse = req.GetResponse


            Str = resp.GetResponseStream
            srRead = New System.IO.StreamReader(Str)
            ' read all the text 
            string_text = srRead.ReadToEnd().ToString

            resp.Close()
            resp = Nothing
            req = Nothing

            string_text = string_text
            original_string_text = string_text




            spot_to_find = InStr(string_text, "mob-title-price")
            If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 2)

                array_split = Split(string_text, "mob-title-price")

                For i = 1 To array_split.Length - 1
                    tcount = tcount + 1  ' do every third one 

                    string_text = array_split(i)
                    original_string_text = string_text



                    acpub_count = acpub_count + 1

                    temp_ac_id = 0
                    temp_ac_name = ""
                    temp_engine = ""
                    temp_eng = ""
                    temp_av = ""

                    pub_reg_no = ""
                    pub_ser_no = ""
                    pub_desc = ""
                    pub_price = ""
                    pub_aftt = ""
                    pub_seller_info = ""
                    pub_picture = ""
                    pub_status = ""
                    pub_url = ""
                    has_pics = False
                    aftt_different = ""
                    temp_year = ""
                    temp_make = ""
                    temp_model = ""
                    pub_comp_id = 0



                    spot_to_find = InStr(string_text, "a href=")
                    If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 8)

                        spot_to_find = InStr(string_text, ">")
                        If spot_to_find > 0 Then
                            pub_url = Left(Trim(string_text), spot_to_find - 2)

                            pub_url = "https://www.avbuyer.com/" & pub_url

                            string_text = Right(string_text, Len(string_text) - spot_to_find)
                            spot_to_find = InStr(string_text, "<")
                            If spot_to_find > 0 Then
                                temp_ac_name = Left(Trim(string_text), spot_to_find - 1)
                            End If

                            spot_to_find = InStr(Trim(pub_url), "title")
                            If spot_to_find > 0 Then
                                pub_url = Left(Trim(pub_url), spot_to_find - 4)
                            End If
                            If InStr(pub_url, "void(") > 0 Then
                                pub_url = ""
                            End If

                            spot_to_find = InStr(Trim(pub_url), " onclick=")
                            If spot_to_find > 0 Then
                                pub_url = Left(Trim(pub_url), spot_to_find - 2)
                            End If

                        End If
                    End If

                    spot_to_find = InStr(string_text, "<div class=""price"">")
                    If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 18)
                        spot_to_find = InStr(string_text, "<")
                        If spot_to_find > 0 Then
                            pub_price = Left(Trim(string_text), spot_to_find - 1)
                        End If

                        pub_price = Replace(pub_price, "Price: ", "")
                        pub_price = Replace(pub_price, "USD ", "")

                        If Len(Trim(pub_price)) > 15 Then
                            pub_price = pub_price
                        End If
                    End If

                    ' FOR SALE BY 
                    spot_to_find = InStr(string_text, "For Sale by")
                    If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 11)

                        spot_to_find = InStr(string_text, "</b>")
                        If spot_to_find > 0 Then
                            pub_seller_info = Left(Trim(string_text), spot_to_find - 1)
                            pub_seller_info = Replace(pub_seller_info, "<b>", "")
                        End If
                    End If



                    spot_to_find = InStr(string_text, "list-other-dtl"">")
                    If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 16)

                        spot_to_find = InStr(string_text, "Year")
                        If spot_to_find > 0 Then
                            string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

                            spot_to_find = InStr(string_text, "</li>")
                            If spot_to_find > 0 Then
                                temp_year = Left(Trim(string_text), spot_to_find - 2)
                                temp_year = Replace(temp_year, "<li>", "")
                            End If
                        End If


                        spot_to_find = InStr(string_text, "S/N")
                        If spot_to_find > 0 Then
                            string_text = Right(string_text, Len(string_text) - spot_to_find - 2)

                            spot_to_find = InStr(string_text, "</li>")
                            If spot_to_find > 0 Then
                                pub_ser_no = Left(Trim(string_text), spot_to_find - 2)
                                pub_ser_no = Replace(pub_ser_no, "<li>", "")
                            End If
                        Else
                            pub_ser_no = pub_ser_no ' there is no serial number 
                        End If


                        spot_to_find = InStr(string_text, "Total Time")
                        If spot_to_find > 0 Then
                            string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

                            spot_to_find = InStr(string_text, "</li>")
                            If spot_to_find > 0 Then
                                pub_aftt = Left(Trim(string_text), spot_to_find - 2)
                                pub_aftt = Replace(pub_aftt, "<li>", "")
                                pub_aftt = Replace(pub_aftt, "'", "")

                                spot_to_find = InStr(pub_aftt, ":")
                                If spot_to_find > 0 Then
                                    pub_aftt = Left(Trim(pub_aftt), spot_to_find - 1)
                                End If



                                'spot_to_find = InStr(pub_aftt, ",")
                                'spot_to_find2 = InStr(pub_aftt, ".")

                                'If spot_to_find2 > spot_to_find And spot_to_find > 0 Then
                                '    ' then there is a comma and a period
                                '    If spot_to_find2 > 0 Then
                                '        pub_aftt = Left(Trim(pub_aftt), spot_to_find2 - 1)
                                '    End If
                                'ElseIf spot_to_find = 0 And spot_to_find2 > 0 Then
                                '    'if there is no comma, but a period, then look.

                                '    'if the length of whats to the left is more than whats to the right - otherwise just do replace
                                '    If Len(Left(Trim(pub_aftt), spot_to_find2 - 1)) < Len(Right(Trim(pub_aftt), Len(Trim(pub_aftt)) - spot_to_find2)) Then
                                '        pub_aftt = pub_aftt
                                '    Else
                                '        If spot_to_find2 > 0 Then
                                '            pub_aftt = Left(Trim(pub_aftt), spot_to_find2 - 1)
                                '        End If
                                '    End If 

                                'End If

                                pub_aftt = Replace(pub_aftt, " ", "")
                                pub_aftt = Replace(pub_aftt, ".", "")
                                pub_aftt = Replace(pub_aftt, ":", "")
                                pub_aftt = Replace(pub_aftt, ",", "")

                            End If
                        End If
                    Else
                        pub_aftt = pub_aftt   ' there is no other details 
                    End If



                    spot_to_find = InStr(string_text, "list-item-para"">")
                    If spot_to_find > 0 Then
                        string_text = Right(string_text, Len(string_text) - spot_to_find - 16)
                        spot_to_find = InStr(Trim(string_text), "</div>")
                        If spot_to_find > 0 Then
                            pub_desc = Left(Trim(string_text), spot_to_find - 2)
                            pub_desc = Replace(pub_desc, "'", "''")
                        End If
                    End If





                    acpub_original_name = temp_ac_name & " " & pub_ser_no

                    cutme(acpub_original_name)
                    cutme(temp_ac_name)
                    cutme(pub_price)
                    cutme(pub_url)
                    cutme(pub_seller_info)
                    cutme(pub_desc)
                    cutme(temp_year)
                    cutme(pub_ser_no)
                    cutme(pub_aftt)

                    pub_aftt = Replace(pub_aftt, "Hours", "")


                    Response.Write("<Br>")
                    Response.Write("<Br>" & temp_ac_name)
                    Response.Write("<Br>" & pub_price)
                    Response.Write("<Br>" & pub_url)
                    Response.Write("<Br>" & pub_seller_info)
                    Response.Write("<Br>" & pub_desc)
                    Response.Write("<Br>" & temp_year)
                    Response.Write("<Br>" & pub_ser_no)
                    Response.Write("<Br>" & pub_aftt)



                    array_split_make = Split(Trim(temp_ac_name), " ")

                    If array_split_make.Length = 2 Then
                        temp_make = array_split_make(0)
                        temp_model = array_split_make(1)
                    ElseIf array_split_make.Length = 3 Then
                        temp_make = array_split_make(1)
                        temp_model = array_split_make(2)
                    ElseIf array_split_make.Length = 4 Then
                        temp_make = array_split_make(2)
                        temp_model = array_split_make(3)
                    ElseIf array_split_make.Length = 5 Then
                        temp_make = array_split_make(3)
                        temp_model = array_split_make(4)
                    Else
                        temp_temp = ""
                    End If

                    If InStr(Trim(pub_ser_no), "&dash;") > 0 Then
                        pub_ser_no = ""
                    End If

                    temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
                    If temp_ac_id = 0 Then
                        temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                        If temp_ac_id = 0 Then
                            temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                            If temp_ac_id = 0 Then

                                temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                                If temp_ac_id = 0 Then
                                    temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                                    If temp_ac_id = 0 Then
                                        temp_ac_id = temp_ac_id
                                    End If
                                End If

                            End If
                        End If

                    End If



                    If pub_comp_id = 0 Then
                        If Not IsNothing(pub_seller_info) Then
                            If Trim(pub_seller_info) <> "" Then
                                'if there is a line feed or break, then try to get just the company name 
                                If InStr(Trim(pub_seller_info), Asc(10)) > 0 Then
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), InStr(Trim(pub_seller_info), Asc(10)) - 1))
                                ElseIf InStr(Trim(pub_seller_info), Asc(13)) > 0 Then
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), InStr(Trim(pub_seller_info), Asc(13)) - 1))
                                Else
                                    pub_comp_id = find_comp_id_global_search(Left(Trim(pub_seller_info), 17))
                                End If
                            End If
                        End If
                    End If

                    If pub_comp_id = 0 And Trim(pub_seller_info <> "") Then
                        pub_comp_id = find_comp_id_previous_pub(pub_seller_info, 2)
                    End If


                    acpub_price_details = ""
                    If On_Naughty_List(temp_ac_name) = True Then
                        ' if its on naughtly list then excldue  
                    Else
                        If temp_ac_id > 0 Then
                            Call find_ac_data(temp_ac_id)
                        Else
                            acpub_process_status = "For Sale Not Found – No AC Match"
                            acpub_status = "O"
                        End If

                        If Trim(aftt_different) <> "" Then
                            pub_desc = pub_desc & aftt_different
                        End If

                        If Trim(acpub_price_details) <> "" Then
                            pub_desc = pub_desc & " " & acpub_price_details
                        End If

                        temp_ac_id = temp_ac_id
                        Call check_insert_ac_pub(temp_ac_id, 7)
                    End If

                    Response.Write("<Br>AC ID:" & temp_ac_id)





                    If tcount = 3 Then
                        tcount = 0
                    End If

                Next


            End If



        Catch ex As Exception
        Finally

        End Try

    End Function



    Public Function scrape_for_AvBuyer_Original(ByVal page_num As Long)
        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader



        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim i As Integer = 0
        Dim final_string As String = ""
        Dim original_string_text As String = ""
        Dim article_link As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim spot_to_find3 As Integer = 0
        Dim array_split() As String
        Dim k As Integer = 0
        Dim skip_this As Boolean = False
        Dim extra_note As String = ""

        Dim temp_ac_name As String = ""
        Dim temp_engine As String = ""
        Dim temp_eng As String = ""
        Dim temp_av As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_make As String = ""
        Dim temp_temp As String
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split_make() As String
        Dim tcount As Integer = 0




        Try


            System.Threading.Thread.Sleep(10)
            Response.Flush()
            System.Threading.Thread.Sleep(10)

            Dim req As System.Net.WebRequest


            req = System.Net.WebRequest.Create("https://www.avbuyer.com/aircraft?page=" & page_num & "&rows=25")
            ' req = System.Net.WebRequest.Create("https://www.avbuyer.com/aircraft/private-jets?page=" & page_num & "")
            'changed from private jets to all 


            Dim resp As System.Net.WebResponse = req.GetResponse


            Str = resp.GetResponseStream
            srRead = New System.IO.StreamReader(Str)
            ' read all the text 
            string_text = srRead.ReadToEnd().ToString

            resp.Close()
            resp = Nothing
            req = Nothing

            string_text = string_text
            original_string_text = string_text


            spot_to_find = InStr(string_text, "<div class=""clearfix"">")
            If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find)

                spot_to_find = InStr(string_text, "<div class=""clearfix"">")
                If spot_to_find > 0 Then
                    string_text = Right(string_text, Len(string_text) - spot_to_find - 2)

                    array_split = Split(string_text, "<div class=""clearfix"">")

                    For i = 1 To array_split.Length - 1
                        tcount = tcount + 1  ' do every third one 
                        string_text = array_split(i)
                        original_string_text = string_text

                        If tcount = 1 Then

                            acpub_count = acpub_count + 1

                            temp_ac_id = 0
                            temp_ac_name = ""
                            temp_engine = ""
                            temp_eng = ""
                            temp_av = ""

                            pub_reg_no = ""
                            pub_ser_no = ""
                            pub_desc = ""
                            pub_price = ""
                            pub_aftt = ""
                            pub_seller_info = ""
                            pub_picture = ""
                            pub_status = ""
                            pub_url = ""
                            has_pics = False
                            aftt_different = ""
                            temp_year = ""
                            temp_make = ""
                            temp_model = ""
                            pub_comp_id = 0



                            spot_to_find = InStr(string_text, "a href=")
                            If spot_to_find > 0 Then
                                string_text = Right(string_text, Len(string_text) - spot_to_find - 8)

                                spot_to_find = InStr(string_text, ">")
                                If spot_to_find > 0 Then
                                    pub_url = Left(Trim(string_text), spot_to_find - 3)

                                    pub_url = "https://www.avbuyer.com/" & pub_url

                                    string_text = Right(string_text, Len(string_text) - spot_to_find)
                                    spot_to_find = InStr(string_text, "<")
                                    If spot_to_find > 0 Then
                                        temp_ac_name = Left(Trim(string_text), spot_to_find - 1)
                                    End If

                                    spot_to_find = InStr(Trim(pub_url), "title")
                                    If spot_to_find > 0 Then
                                        pub_url = Left(Trim(pub_url), spot_to_find - 4)
                                    End If
                                    If InStr(pub_url, "void(") > 0 Then
                                        pub_url = ""
                                    End If

                                    spot_to_find = InStr(Trim(pub_url), " onclick=")
                                    If spot_to_find > 0 Then
                                        pub_url = Left(Trim(pub_url), spot_to_find - 2)
                                    End If

                                End If
                            End If

                            spot_to_find = InStr(string_text, "itemprop=""price"">")
                            If spot_to_find > 0 Then
                                string_text = Right(string_text, Len(string_text) - spot_to_find - 16)
                                spot_to_find = InStr(string_text, "<")
                                If spot_to_find > 0 Then
                                    pub_price = Left(Trim(string_text), spot_to_find - 1)
                                End If
                            End If

                            ' FOR SALE BY 
                            spot_to_find = InStr(string_text, "fl")
                            If spot_to_find > 0 Then
                                string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

                                spot_to_find = InStr(string_text, "</b>")
                                If spot_to_find > 0 Then
                                    pub_seller_info = Left(Trim(string_text), spot_to_find - 1)
                                    pub_seller_info = Replace(pub_seller_info, "<b>", "")
                                End If
                            End If

                        End If


                        If InStr(temp_ac_name, "Embraer Phenom 100E") > 0 Then
                            temp_ac_name = temp_ac_name
                        End If



                        If tcount = 2 Then

                            spot_to_find = InStr(string_text, "prod_desc")
                            If spot_to_find > 0 Then
                                string_text = Right(string_text, Len(string_text) - spot_to_find - 10)
                                spot_to_find = InStr(Trim(string_text), "<a href")
                                If spot_to_find > 0 Then
                                    pub_desc = Left(Trim(string_text), spot_to_find - 2)
                                    pub_desc = Replace(pub_desc, "'", "''")
                                End If
                            End If

                            spot_to_find = InStr(string_text, "other_info")
                            If spot_to_find > 0 Then
                                string_text = Right(string_text, Len(string_text) - spot_to_find)

                                spot_to_find = InStr(string_text, "Year")
                                If spot_to_find > 0 Then
                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

                                    spot_to_find = InStr(string_text, "</b>")
                                    If spot_to_find > 0 Then
                                        temp_year = Left(Trim(string_text), spot_to_find - 2)
                                        temp_year = Replace(temp_year, "<b>", "")
                                    End If
                                End If


                                spot_to_find = InStr(string_text, "S/N")
                                If spot_to_find > 0 Then
                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 2)

                                    spot_to_find = InStr(string_text, "</b>")
                                    If spot_to_find > 0 Then
                                        pub_ser_no = Left(Trim(string_text), spot_to_find - 2)
                                        pub_ser_no = Replace(pub_ser_no, "<b>", "")
                                    End If
                                End If


                                spot_to_find = InStr(string_text, "TTAF")
                                If spot_to_find > 0 Then
                                    string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

                                    spot_to_find = InStr(string_text, "</b>")
                                    If spot_to_find > 0 Then
                                        pub_aftt = Left(Trim(string_text), spot_to_find - 2)
                                        pub_aftt = Replace(pub_aftt, "<b>", "")
                                        pub_aftt = Replace(pub_aftt, "'", "")

                                        spot_to_find = InStr(pub_aftt, ":")
                                        If spot_to_find > 0 Then
                                            pub_aftt = Left(Trim(pub_aftt), spot_to_find - 1)
                                        End If



                                        spot_to_find = InStr(pub_aftt, ",")
                                        spot_to_find2 = InStr(pub_aftt, ".")

                                        If spot_to_find2 > spot_to_find And spot_to_find > 0 Then
                                            ' then there is a comma and a period
                                            If spot_to_find2 > 0 Then
                                                pub_aftt = Left(Trim(pub_aftt), spot_to_find2 - 1)
                                            End If
                                        ElseIf spot_to_find = 0 And spot_to_find2 > 0 Then
                                            'if there is no comma, but a period, then look.

                                            'if the length of whats to the left is more than whats to the right - otherwise just do replace
                                            If Len(Left(Trim(pub_aftt), spot_to_find2 - 1)) < Len(Right(Trim(pub_aftt), Len(Trim(pub_aftt)) - spot_to_find2)) Then
                                                pub_aftt = pub_aftt
                                            Else
                                                If spot_to_find2 > 0 Then
                                                    pub_aftt = Left(Trim(pub_aftt), spot_to_find2 - 1)
                                                End If
                                            End If




                                        End If

                                        pub_aftt = Replace(pub_aftt, " ", "")
                                        pub_aftt = Replace(pub_aftt, ".", "")
                                        pub_aftt = Replace(pub_aftt, ":", "")
                                        pub_aftt = Replace(pub_aftt, ",", "")

                                    End If
                                End If


                            End If
                        End If



                        If tcount = 3 Then

                            If InStr(Trim(temp_ac_name), "Learjet 45XR") > 0 Then
                                temp_ac_name = temp_ac_name
                            End If



                            acpub_original_name = temp_ac_name & " " & pub_ser_no

                            cutme(acpub_original_name)
                            cutme(temp_ac_name)
                            cutme(pub_price)
                            cutme(pub_url)
                            cutme(pub_seller_info)
                            cutme(pub_desc)
                            cutme(temp_year)
                            cutme(pub_ser_no)
                            cutme(pub_aftt)

                            pub_aftt = Replace(pub_aftt, "Hours", "")


                            Response.Write("<Br>")
                            Response.Write("<Br>" & temp_ac_name)
                            Response.Write("<Br>" & pub_price)
                            Response.Write("<Br>" & pub_url)
                            Response.Write("<Br>" & pub_seller_info)
                            Response.Write("<Br>" & pub_desc)
                            Response.Write("<Br>" & temp_year)
                            Response.Write("<Br>" & pub_ser_no)
                            Response.Write("<Br>" & pub_aftt)



                            array_split_make = Split(Trim(temp_ac_name), " ")

                            If array_split_make.Length = 2 Then
                                temp_make = array_split_make(0)
                                temp_model = array_split_make(1)
                            ElseIf array_split_make.Length = 3 Then
                                temp_make = array_split_make(1)
                                temp_model = array_split_make(2)
                            ElseIf array_split_make.Length = 4 Then
                                temp_make = array_split_make(2)
                                temp_model = array_split_make(3)
                            ElseIf array_split_make.Length = 5 Then
                                temp_make = array_split_make(3)
                                temp_model = array_split_make(4)
                            Else
                                temp_temp = ""
                            End If

                            If InStr(Trim(pub_ser_no), "&dash;") > 0 Then
                                pub_ser_no = ""
                            End If

                            temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
                            If temp_ac_id = 0 Then
                                temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                                If temp_ac_id = 0 Then
                                    temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                                    If temp_ac_id = 0 Then

                                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                                        If temp_ac_id = 0 Then
                                            temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                                            If temp_ac_id = 0 Then
                                                temp_ac_id = temp_ac_id
                                            End If
                                        End If

                                    End If
                                End If

                            End If



                            acpub_price_details = ""
                            If On_Naughty_List(temp_ac_name) = True Then
                                ' if its on naughtly list then excldue  
                            Else
                                If temp_ac_id > 0 Then
                                    Call find_ac_data(temp_ac_id)
                                Else
                                    acpub_process_status = "For Sale Not Found – No AC Match"
                                    acpub_status = "O"
                                End If

                                If Trim(aftt_different) <> "" Then
                                    pub_desc = pub_desc & aftt_different
                                End If

                                If Trim(acpub_price_details) <> "" Then
                                    pub_desc = pub_desc & " " & acpub_price_details
                                End If

                                Call check_insert_ac_pub(temp_ac_id, 7)
                            End If

                            Response.Write("<Br>AC ID:" & temp_ac_id)
                        End If





                        If tcount = 3 Then
                            tcount = 0
                        End If

                    Next


                End If
            Else
                temp_ac_id = temp_ac_id ' NO PAGE ITEMS FOUND - 
            End If




        Catch ex As Exception
        Finally

        End Try

    End Function
    Public Function find_scraped_aircraft() As DataTable

        find_scraped_aircraft = Nothing
        Dim Insert_Query As String = ""
        Dim select_query As String = ""
        Dim atemptable As New DataTable
        Dim passed_test As Boolean = True
        Dim orig_serno As String = ""

        Try

            select_query = "  Select distinct scrp_ac_id, scrp_ac_source, scrp_ac_model, scrp_ac_asking_price , scrp_ac_detail_link, "
            select_query &= " scrp_ac_location , scrp_ac_dealer, scrp_ac_year, scrp_ac_ser_no , scrp_ac_airframe_tot_hrs , scrp_ac_note  "
            select_query &= " from scraped_aircraft "
            select_query &= " where scrp_ac_processed = 'N' and scrp_ac_cleansed = 'Y' "
            ' select_query &= " where scrp_ac_asking_price = 'Off market' "

            MySqlCommand_JETNET.CommandText = select_query
            MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

            Try
                atemptable.Load(MyAircraftReader_JETNET)
            Catch constrExc As System.Data.ConstraintException
                Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
            End Try

            find_scraped_aircraft = atemptable

        Catch ex As Exception
        Finally
            If Not MyAircraftReader_JETNET.IsClosed Then
                MyAircraftReader_JETNET.Close()
            End If
            MySqlCommand_JETNET.Dispose()
            atemptable = Nothing
        End Try
    End Function


    Public Function find_Pending_Reg(ByVal reg_no As String, ByVal ac_id As Long) As Boolean

        find_Pending_Reg = False

        Dim Insert_Query As String = ""
        Dim select_query As String = ""
        Dim atemptable As New DataTable
        Dim passed_test As Boolean = True
        Dim orig_serno As String = ""



        Try

            reg_no = Trim(reg_no)

            If ac_id > 0 Then


                select_query = " select * From Aircraft_FAA_Document "
                select_query &= " where  acfaa_ac_id =" & ac_id
                select_query &= " and acfaa_journ_id = 0 and (acfaa_reg_no1 = '" & reg_no & "' or acfaa_reg_no2 ='" & reg_no & "' or acfaa_reg_no3 ='" & reg_no & "'  or acfaa_reg_no4 ='" & reg_no & "') "


                MySqlCommand_JETNET.CommandText = select_query
                MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

                Try
                    atemptable.Load(MyAircraftReader_JETNET)
                Catch constrExc As System.Data.ConstraintException
                    Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                    ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
                End Try

                If atemptable.Rows.Count >= 1 Then
                    ' tepm_Ac_id = atemptable(0).Item("acfaa_ac_id")
                    find_Pending_Reg = True
                End If
                atemptable.Clear()

            ElseIf Trim(reg_no) <> "" Then

                select_query = " select * From Aircraft_FAA_Document "
                select_query &= " where (acfaa_reg_no1 = '" & reg_no & "' or acfaa_reg_no2 ='" & reg_no & "' or acfaa_reg_no3 ='" & reg_no & "'  or acfaa_reg_no4 ='" & reg_no & "') and acfaa_journ_id = 0 "


                MySqlCommand_JETNET.CommandText = select_query
                MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

                Try
                    atemptable.Load(MyAircraftReader_JETNET)
                Catch constrExc As System.Data.ConstraintException
                    Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                    ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
                End Try

                If atemptable.Rows.Count >= 1 Then
                    '  find_Pending_Reg = atemptable(0).Item("acfaa_ac_id")
                    find_Pending_Reg = True
                End If
                atemptable.Clear()
            End If


        Catch ex As Exception
        Finally
            If Not MyAircraftReader_JETNET.IsClosed Then
                MyAircraftReader_JETNET.Close()
            End If
            MySqlCommand_JETNET.Dispose()
            atemptable = Nothing
        End Try
    End Function


    Public Function find_ac_global_search(ByVal ser_no As String, ByVal make As String, ByVal model1 As String, ByVal reg_no As String) As Long

        find_ac_global_search = 0

        Dim Insert_Query As String = ""
        Dim select_query As String = ""
        Dim atemptable As New DataTable
        Dim passed_test As Boolean = True
        Dim orig_serno As String = ""

        Try
            ' only while we are in here - replace the - so it searches more correctly 
            orig_serno = ser_no
            ser_no = Replace(ser_no, "-", "")
            make = Replace(make, "-", "")
            model1 = Replace(model1, "-", "")

            select_query = " select fts_ac_id FROM Full_Text_Search WITH(NOLOCK)  "
            select_query &= " inner join Aircraft with (NOLOCK) on ac_id = fts_ac_id and ac_journ_id = 0 "

            If Trim(ser_no) <> "" Then
                select_query &= " WHERE (contains (Full_Text_Search.*, '""" & Replace(LCase(ser_no), " ", "") & "*""')"

                'if we have a legti serial number, make sure its for the serial number and not a reg or previous reg 
                If Len(Trim(ser_no)) > 2 Then
                    select_query &= " and (ac_ser_no_full like '%" & Replace(Replace(LCase(ser_no), "#", ""), " ", "") & "%' or ac_ser_no_full like '%" & Replace(Replace(LCase(orig_serno), "#", ""), " ", "") & "%' or replace(ac_ser_no_full, '-', '') like '%" & Replace(Replace(LCase(ser_no), "#", ""), " ", "") & "%')   "
                    'select_query &= " and ac_prev_reg_no not like '%" & Replace(Replace(LCase(ser_no), "#", ""), " ", "") & "%' " ' and its  not the prev reg 
                    ' select_query &= " and ac_reg_no not like '%" & Replace(Replace(LCase(ser_no), "#", ""), " ", "") & "%' "
                End If


                If Trim(make) <> "" Then
                    If InStr(make, "BOMBARDIER/CHALLENGER") > 0 Then
                        select_query &= " AND contains (Full_Text_Search.*, '""" & Replace(UCase(Replace(LCase(make), " ", "")), "BOMBARDIER/CHALLENGER", "CHALLENGER") & "*""') "
                    Else
                        select_query &= " AND contains (Full_Text_Search.*, '""" & Replace(LCase(make), " ", "") & "*""') "
                    End If
                End If

                If Trim(model1) <> "" Then
                    select_query &= " AND contains (Full_Text_Search.*, '""" & Replace(LCase(model1), " ", "") & "*""')"
                End If

                If Trim(reg_no) <> "" Then
                    select_query &= " AND contains (Full_Text_Search.*, '""" & Replace(Replace(LCase(reg_no), "#", ""), " ", "") & "*""')"
                    select_query &= " and (ac_prev_reg_no <> '" & Replace(Replace(LCase(reg_no), "#", ""), " ", "") & "'  or ac_prev_reg_no is null) " ' and its  not the prev reg  
                End If

                select_query &= ") "
            ElseIf Trim(reg_no) <> "" Then


                select_query &= " WHERE contains (Full_Text_Search.*, '""" & Replace(Replace(LCase(reg_no), "#", ""), " ", "") & "*""')"
                select_query &= " and (ac_prev_reg_no <> '" & Replace(Replace(LCase(reg_no), "#", ""), " ", "") & "'  or ac_prev_reg_no is null) " ' and its  not the prev reg 


                If Trim(make) <> "" Then
                    If InStr(make, "BOMBARDIER/CHALLENGER") > 0 Then
                        select_query &= " AND contains (Full_Text_Search.*, '""" & Replace(UCase(Replace(LCase(make), " ", "")), "BOMBARDIER/CHALLENGER", "CHALLENGER") & "*""') "
                    Else
                        select_query &= " AND contains (Full_Text_Search.*, '""" & Replace(LCase(make), " ", "") & "*""') "
                    End If
                End If

                If Trim(model1) <> "" Then
                    select_query &= " AND contains (Full_Text_Search.*, '""" & Replace(LCase(model1), " ", "") & "*""')"
                End If
            End If

            select_query &= " and fts_ac_id > 0 "

            If Trim(ser_no) <> "" Or Trim(reg_no) <> "" Then
                MySqlCommand_JETNET.CommandText = select_query
                MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

                Try
                    atemptable.Load(MyAircraftReader_JETNET)
                Catch constrExc As System.Data.ConstraintException
                    Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                    ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
                End Try

                If atemptable.Rows.Count = 1 Then
                    find_ac_global_search = atemptable(0).Item("fts_ac_id")
                End If
                atemptable.Clear()
            End If


        Catch ex As Exception
        Finally
            If Not MyAircraftReader_JETNET.IsClosed Then
                MyAircraftReader_JETNET.Close()
            End If
            MySqlCommand_JETNET.Dispose()
            atemptable = Nothing
        End Try
    End Function

    Public Function find_ac_ac_search(ByVal ser_no As String, ByVal make As String, ByVal model1 As String, ByVal reg_no As String) As Long

        find_ac_ac_search = 0

        Dim Insert_Query As String = ""
        Dim select_query As String = ""
        Dim atemptable As New DataTable
        Dim passed_test As Boolean = True
        Dim orig_serno As String = ""

        Try
            ' only while we are in here - replace the - so it searches more correctly 
            orig_serno = ser_no

            select_query = " select ac_id FROM aircraft WITH(NOLOCK)  "
            select_query &= " inner Join aircraft_model with (NOLOCK) on ac_amod_id = amod_id "
            select_query &= " where ac_journ_id = 0 "

            If Trim(make) <> "" Then
                select_query &= " and amod_make_name = '" & make & "' "
            End If

            If Trim(model1) <> "" Then
                select_query &= " and amod_model_name = '" & model1 & "' "
            End If

            If Trim(ser_no) <> "" Then
                If IsNumeric(Trim(ser_no)) Then
                    select_query &= " and (ac_ser_no = '" & Trim(ser_no) & "' or ac_ser_no_value = '" & Trim(ser_no) & "' or ac_ser_no_full = '" & Trim(ser_no) & "') "
                Else
                    select_query &= " and (ac_ser_no = '" & Trim(ser_no) & "' or ac_ser_no_full = '" & Trim(ser_no) & "') "
                End If
            End If

            If Trim(reg_no) <> "" Then
                select_query &= " and (ac_reg_no = '" & Trim(reg_no) & "') "
            End If


            MySqlCommand_JETNET.CommandText = select_query
            MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

            Try
                atemptable.Load(MyAircraftReader_JETNET)
            Catch constrExc As System.Data.ConstraintException
                Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
            End Try

            If atemptable.Rows.Count = 1 Then
                find_ac_ac_search = atemptable(0).Item("ac_id")
            End If
            atemptable.Clear()


        Catch ex As Exception
        Finally
            If Not MyAircraftReader_JETNET.IsClosed Then
                MyAircraftReader_JETNET.Close()
            End If
            MySqlCommand_JETNET.Dispose()
            atemptable = Nothing
        End Try
    End Function


    Public Function find_comp_id_blind_pub(ByVal company_id As String, ByVal pub_source As Long) As Boolean

        find_comp_id_blind_pub = False

        Dim Insert_Query As String = ""
        Dim select_query As String = ""
        Dim atemptable As New DataTable
        Dim passed_test As Boolean = True

        Try

            select_query = " select distinct publist_comp_id from Publication_Listing with (NOLOCK) "

            select_query &= " where publist_comp_id= '" & Trim(company_id) & "'  "
            select_query &= " and publist_source = '" & pub_source & "' "
            select_query &= " and publist_comp_id > 0  "
            select_query &= " and publist_research_note like '%no blind pubs%'  "
            select_query &= " And publist_entry_date >= (getdate() - 90) "

            MySqlCommand_JETNET.CommandText = select_query
            MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

            Try
                atemptable.Load(MyAircraftReader_JETNET)
            Catch constrExc As System.Data.ConstraintException
                Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
            End Try

            If atemptable.Rows.Count = 1 Then
                find_comp_id_blind_pub = True
            End If
            atemptable.Clear()


        Catch ex As Exception
        Finally
            MySqlCommand_JETNET.Dispose()
            atemptable = Nothing
        End Try
    End Function
    Public Function find_comp_id_previous_pub(ByVal company_name As String, ByVal pub_source As Long) As Long

        find_comp_id_previous_pub = 0

        Dim Insert_Query As String = ""
        Dim select_query As String = ""
        Dim atemptable As New DataTable
        Dim passed_test As Boolean = True

        Try

            select_query = " select distinct publist_comp_id from Publication_Listing with (NOLOCK) "

            select_query &= " where (publist_seller_info = '" & Trim(company_name) & "' or publist_seller_info like '%" & Trim(company_name) & "%') "
            select_query &= " and publist_source = '" & pub_source & "' "
            select_query &= " and publist_comp_id > 0  "

            MySqlCommand_JETNET.CommandText = select_query
            MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

            Try
                atemptable.Load(MyAircraftReader_JETNET)
            Catch constrExc As System.Data.ConstraintException
                Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
            End Try

            If atemptable.Rows.Count = 1 Then
                find_comp_id_previous_pub = atemptable(0).Item("publist_comp_id")
            End If
            atemptable.Clear()


        Catch ex As Exception
        Finally
            MySqlCommand_JETNET.Dispose()
            atemptable = Nothing
        End Try
    End Function

    Public Function find_comp_id_global_search(ByVal company_name As String, Optional ByVal pub_city As String = "") As Long

        find_comp_id_global_search = 0

        Dim Insert_Query As String = ""
        Dim select_query As String = ""
        Dim atemptable As New DataTable
        Dim passed_test As Boolean = True

        Try

            select_query = " select fts_comp_id FROM Full_Text_Search WITH(NOLOCK)  "
            select_query &= " WHERE  contains (Full_Text_Search.*, '""" & LCase(company_name) & "*""') "
            select_query &= " and fts_comp_id > 0 "


            If Trim(pub_city) <> "" Then
                select_query &= " and  contains (Full_Text_Search.*, '""" & LCase(pub_city) & "*""') "
            End If



            MySqlCommand_JETNET.CommandText = select_query
            MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

            Try
                atemptable.Load(MyAircraftReader_JETNET)
            Catch constrExc As System.Data.ConstraintException
                Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
            End Try

            If atemptable.Rows.Count = 1 Then
                find_comp_id_global_search = atemptable(0).Item("fts_comp_id")
            End If
            atemptable.Clear()


        Catch ex As Exception
        Finally
            MySqlCommand_JETNET.Dispose()
            atemptable = Nothing
        End Try
    End Function
    Public Function find_ac_data(ByVal ac_id As Long) As String
    find_ac_data = ""

    Dim select_query As String = ""
    Dim atemptable As New DataTable
    Dim passed_test As Boolean = True
    Dim temp_aftt As String = ""
    Dim temp_for_sale As String = "N"
    Dim asking_price As Long = 0
    Dim temp_asking As String = ""
    Dim temp_seq As Integer = 0
    Dim temp_landings As String = ""

    Try

      If has_pics = True Then
        select_query = " select top 1 ac_forsale_flag, ac_asking, ac_asking_price, ac_airframe_tot_hrs, acpic_seq_no, ac_airframe_tot_landings "
        select_query &= " from Aircraft with (NOLOCK) "
        select_query &= " left outer join aircraft_pictures with (NOLOCK) on acpic_ac_id = ac_id   and acpic_hide_flag = 'N'  "
        select_query &= " where ac_id = " & ac_id & " and ac_journ_id = 0 "
        select_query &= " order by acpic_seq_no asc "
      Else
        select_query = " select  ac_forsale_flag, ac_asking, ac_asking_price, ac_airframe_tot_hrs, ac_airframe_tot_landings  from Aircraft with (NOLOCK) where ac_id = " & ac_id & " and ac_journ_id = 0 "
      End If

      If ac_id = 117692 Or ac_id = 32249 Then
        ac_id = ac_id
      End If


      MySqlCommand_JETNET.CommandText = select_query
      MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

      Try
        atemptable.Load(MyAircraftReader_JETNET)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

      If atemptable.Rows.Count = 1 Then

        If Not IsDBNull(atemptable(0).Item("ac_forsale_flag")) Then
          temp_for_sale = atemptable(0).Item("ac_forsale_flag")
        End If

        If Not IsDBNull(atemptable(0).Item("ac_asking")) Then
          temp_asking = atemptable(0).Item("ac_asking")
        End If

        If Not IsDBNull(atemptable(0).Item("ac_asking_price")) Then
          asking_price = atemptable(0).Item("ac_asking_price")
        End If

        If Not IsDBNull(atemptable(0).Item("ac_airframe_tot_hrs")) Then
          temp_aftt = atemptable(0).Item("ac_airframe_tot_hrs")
        End If

        If Not IsDBNull(atemptable(0).Item("ac_airframe_tot_landings")) Then
          temp_landings = atemptable(0).Item("ac_airframe_tot_landings")
        End If

        If has_pics = True Then
          If Not IsDBNull(atemptable(0).Item("acpic_seq_no")) Then
            temp_seq = atemptable(0).Item("acpic_seq_no")
          Else
            temp_seq = 0
          End If
        Else
          temp_seq = 0
        End If

      End If
      atemptable.Clear()


      acpub_price_details = ""
      acpub_process_status = ""
      If Trim(temp_for_sale) <> "Y" Then
        acpub_process_status = "For Sale Not Found – Record Not for Sale"
        If IsNumeric(pub_price) = True Then
          If CInt(pub_price) = 0 Then
            acpub_price_details = "No Listed Price vs. Not For Sale"
          Else
            acpub_price_details = FormatNumber(pub_price, 0) & " vs. Not For Sale"
          End If
        Else
          acpub_price_details = pub_price & " vs. Not For Sale"
        End If
        acpub_status = "O"
      Else
                If Trim(pub_price) <> "" Then
                    If Trim(pub_price) = "Call" Or Trim(pub_price) = "Call for Price" Or Trim(pub_price) = "Contact Broker" Or Trim(pub_price) = "Contact Broker" Or Trim(pub_price) = "Make an Offer" Or Trim(pub_price) = "Make Offer" Or Trim(pub_price) = "Please Call" Or InStr(Trim(pub_price), "Inquire") > 0 Then  ' if its for sale, call then ours is better 
                        acpub_process_status = "For Sale Found  – Exact Match"
                        acpub_status = "N"
                    ElseIf Trim(pub_price) = "Now Sold" Or Trim(pub_price) = "Deal Pending" Then
                        ' if it says now sold or deal pending and we still have for sale 
                        acpub_process_status = "For Sale Other"
                    ElseIf IsNumeric(pub_price) = True Then
                        If FormatNumber(pub_price, 0) = FormatNumber(asking_price, 0) Then
                            acpub_process_status = "For Sale Found  – Exact Match"
                            acpub_status = "N"
                        ElseIf ((FormatNumber(pub_price, 0) * 0.05) + FormatNumber(pub_price, 0) > FormatNumber(asking_price, 0)) And (FormatNumber(pub_price, 0) - (FormatNumber(pub_price, 0) * 0.05) < FormatNumber(asking_price, 0)) Then
                            acpub_process_status = "For Sale Found  – Close To Match"
                            acpub_status = "N"
                        Else
                            acpub_process_status = "For Sale Found – Price Difference"
                            acpub_price_details = FormatNumber(pub_price, 0) & " vs. " & FormatNumber(asking_price, 0)
                            acpub_status = "O"
                        End If
                    ElseIf IsNumeric(Trim(pub_price)) = False Then ' added in case we miss the wording
                        acpub_process_status = "For Sale Found  – Exact Match"
                        acpub_status = "N"
                    Else
                        acpub_status = acpub_status
                    End If
                ElseIf Trim(pub_price) = "" And asking_price = 0 Then  '  Added MSW - if we are both no prices listed then here 
                    acpub_process_status = "For Sale Found  – Exact Match"
                    acpub_status = "N"
                Else
          acpub_status = acpub_status
        End If
      End If


      aftt_different = ""
            If Trim(temp_aftt) <> "" And Trim(pub_aftt) <> "" Then
                If CDbl(Trim(temp_aftt)) = CDbl(Trim(pub_aftt)) Then
                    aftt_different = ""
                    'what we have is greater than thers, yet theirs + 10 >= ours 
                ElseIf CDbl(Trim(temp_aftt)) > CDbl(Trim(pub_aftt)) And CDbl(CDbl(Trim(pub_aftt)) + 10) >= CDbl(Trim(temp_aftt)) Then
                    aftt_different = ""
                    ' else if ours is less than theirs and theirs minus 10 is less than or = ours
                ElseIf CDbl(Trim(temp_aftt)) < CDbl(Trim(pub_aftt)) And CDbl(CDbl(Trim(pub_aftt)) - 10) <= CDbl(Trim(temp_aftt)) Then
                    aftt_different = ""
                ElseIf CDbl(Trim(temp_aftt)) > CDbl(Trim(pub_aftt)) Then
                    aftt_different = "" ' if our aftt > theirs
                Else
                    aftt_different = "AFTT Difference: " & pub_aftt & " vs. " & temp_aftt
                    acpub_status = "O"
                End If
            ElseIf Trim(pub_aftt) <> "" Then ' their value 
                aftt_different = "AFTT Difference: " & pub_aftt & " vs. EMPTY"
                acpub_status = "O"
            ElseIf Trim(temp_aftt) <> "" Then  ' our value 
                aftt_different = ""
       End If

      landings_different = ""
      If Trim(pub_landings) <> "" Then
        If Trim(temp_landings) <> Trim(pub_landings) Then
          If IsNumeric(Trim(temp_landings)) And IsNumeric(Trim(pub_landings)) Then
            If (CDbl(Trim(temp_landings)) > CDbl(Trim(pub_landings)) And CDbl(CDbl(Trim(pub_landings)) + 10) >= CDbl(Trim(temp_landings)) Or CDbl(Trim(temp_landings)) < CDbl(Trim(pub_landings)) And CDbl(CDbl(Trim(pub_landings)) - 10) <= CDbl(Trim(temp_landings))) Then
              ' then they are within 10 of eachother 
              temp_landings = temp_landings
            Else
              landings_different = "Landings Difference: " & pub_landings & " vs. " & temp_landings
            End If
          ElseIf Trim(temp_landings) = "" Then
            landings_different = "Landings Difference: " & pub_landings & " vs. 0 "
          End If
        End If
      End If

      If has_pics = True And temp_seq = 0 Then
        pub_desc = pub_desc & " Pictures Found"
        acpub_status = "O"
      End If



      'acpub_status = "R"


    Catch ex As Exception
    Finally
      MySqlCommand_JETNET.Dispose()
      atemptable = Nothing
    End Try
  End Function
   Public Function Find_Naughty_Models() As String
    Find_Naughty_Models = ""

    Dim select_query As String = ""
    Dim atemptable As New DataTable
    Dim passed_test As Boolean = True
    Dim temp_aftt As String = ""
    Dim temp_for_sale As String = "N"
    Dim asking_price As Long = 0
    Dim temp_asking As String = ""
    Dim temp_seq As Integer = 0

    Try

            select_query = " select distinct pubnot_phrase from Publication_Models_Not_Processed with (NOLOCK) "
            MySqlCommand_JETNET.CommandText = select_query
      MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

      Try
        atemptable.Load(MyAircraftReader_JETNET)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

      If atemptable.Rows.Count > 0 Then
        For Each r As DataRow In atemptable.Rows

          If Not IsDBNull(r.Item("pubnot_phrase")) Then
            Naughty_List_Of_Models(Naughty_List_Size) = r.Item("pubnot_phrase")
            Naughty_List_Size = Naughty_List_Size + 1
          End If

        Next
      End If
      atemptable.Clear()




    Catch ex As Exception
    Finally
      MySqlCommand_JETNET.Dispose()
      atemptable = Nothing
    End Try
  End Function
  Public Function insert_into_eventlog(ByVal temp_type_run As String, ByVal type_of As String, Optional ByVal date_to_enter As String = "") As Boolean


    Dim Insert_Query As String = ""
    Dim select_query As String = ""
    Dim passed_test As Boolean = True


    Try



      Insert_Query = " INSERT INTO EventLog"
      Insert_Query &= " (evtl_date"
      Insert_Query &= " ,evtl_user_id"
      Insert_Query &= " ,evtl_type"
      Insert_Query &= " ,evtl_message"
      Insert_Query &= " ,evtl_ac_id"
      Insert_Query &= " ,evtl_journ_id"
      Insert_Query &= " ,evtl_comp_id"
      Insert_Query &= " ,evtl_host_name"
      Insert_Query &= " ,evtl_app_name"
      Insert_Query &= " ,evtl_yacht_id"
      Insert_Query &= " ,evtl_contact_id)"
      Insert_Query &= " VALUES( "
      If Trim(date_to_enter) <> "" Then
        Insert_Query &= " '" & Trim(date_to_enter) & "' "
      Else
        Insert_Query &= " '" & Date.Now & "' "
      End If


      Insert_Query &= ",'mvit'"
      Insert_Query &= ", '" & type_of & "'"
      Insert_Query &= ", '" & temp_type_run & "'"
      Insert_Query &= ", '0'"
      Insert_Query &= ", '0'"
      Insert_Query &= ", '0'"
      Insert_Query &= ", 'RASSIST'"
      Insert_Query &= ", 'Research Assistant'"
      Insert_Query &= ", '0'"
      Insert_Query &= ", '0')"


      ' Response.Write("<Br>" & Insert_Query)
      Insert_Query = Insert_Query
      MySqlCommand_JETNET.CommandText = Insert_Query
      MySqlCommand_JETNET.ExecuteNonQuery()


    Catch ex As Exception
    Finally
      MySqlCommand_JETNET.Dispose()
    End Try
  End Function
    Public Function check_insert_ac_pub(ByVal ac_id As Long, ByVal temp_publog_source As String) As Boolean


        Dim Insert_Query As String = ""
        Dim select_query As String = ""
        Dim atemptable As New DataTable
        Dim passed_test As Boolean = True
        Dim used_url As Boolean = False
        Dim newdate As New Date
        Dim skip_insert_update As Boolean = False
        Dim record_has_ac_id As Boolean = False
        Dim record_has_comp_id As Boolean = False
        Dim find_select As String = ""

        Try


            cutme(pub_reg_no)
            cutme(pub_ser_no)
            cutme(pub_aftt)
            cutme(acpub_original_name)
            cutme_LF(acpub_original_name)
            pub_aftt = Replace(pub_aftt, "Hours", "")

            pub_price = Replace(pub_price, "</span>", "")
            pub_price = Replace(pub_price, "</span", "")

            pub_desc = Replace(pub_desc, "</span>", "")
            pub_desc = Replace(pub_desc, "</span", "")

            pub_aftt = Replace(pub_aftt, "Deliveryhoursonly", "")
            pub_aftt = Replace(pub_aftt, "DeliveryTimeOnly", "")

            pub_seller_info = Replace(pub_seller_info, "'", "")



            If InStr(acpub_original_name, "Embraer Phenom 100") > 0 Then
                acpub_original_name = acpub_original_name
            End If

            acpub_original_name = Replace(acpub_original_name, "&dash;", "")


            If pub_comp_id > 0 Then
                If find_comp_id_blind_pub(pub_comp_id, temp_publog_source) = True Then
                    skip_insert_update = True
                End If
            End If




            select_query = " select publist_ac_id from Publication_Listing with (NOLOCK) "

            select_query &= " where REPLACE(REPLACE(publist_original_desc, char(10), ''), char(13), '')  = '" & Trim(acpub_original_name) & "' "
            'select_query &= " where dbo.LeaveAlphaAndNumericAndSpace(publist_original_desc) = '" & Trim(acpub_original_name) & "' "

            '  publist_ac_id = " & ac_id & "" 
            newdate = DateAdd(DateInterval.Day, -14, Now())

            select_query &= " and publist_entry_date >= '" & newdate.Year & "-" & newdate.Month & "-" & newdate.Day & "' "
            select_query &= " and publist_url = '" & Trim(pub_url) & "' "

            If Trim(temp_publog_source) <> "" Then
                select_query &= " and publist_source = '" & temp_publog_source & "' "
            End If


            MySqlCommand_JETNET.CommandText = select_query
            find_select = select_query
            MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

            Try
                atemptable.Load(MyAircraftReader_JETNET)
            Catch constrExc As System.Data.ConstraintException
                Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
            End Try




            If atemptable.Rows.Count > 0 Then ' if we didnt find any, replace the tbd, see if we do 
                ' then dont worry 
            ElseIf Right(Trim(UCase(acpub_original_name)), 3) = "TBD" Then
                ' if we didnt find anything, make sure we dont for the different source.  

                MyAircraftReader_JETNET.Close()

                If Right(Trim(UCase(acpub_original_name)), 3) = "TBD" Then
                    acpub_original_name = Left(Trim(acpub_original_name), Len(Trim(acpub_original_name)) - 3)
                End If
                cutme(acpub_original_name)
                cutme_LF(acpub_original_name)
                acpub_original_name = Trim(acpub_original_name)

                select_query = " select publist_ac_id from Publication_Listing with (NOLOCK) "

                select_query &= " where REPLACE(REPLACE(publist_original_desc, char(10), ''), char(13), '')  = '" & Trim(acpub_original_name) & "' "
                'select_query &= " where dbo.LeaveAlphaAndNumericAndSpace(publist_original_desc) = '" & Trim(acpub_original_name) & "' "

                select_query &= " and publist_entry_date >= '" & newdate.Year & "-" & newdate.Month & "-" & newdate.Day & "' "
                select_query &= " and publist_url = '" & Trim(pub_url) & "' "

                '  publist_ac_id = " & ac_id & ""

                If Trim(temp_publog_source) <> "" Then
                    select_query &= " and publist_source = '" & temp_publog_source & "' "
                End If

                If InStr(acpub_original_name, "1994 Hawker 400XPR") > 0 Then
                    acpub_original_name = acpub_original_name
                End If


                MySqlCommand_JETNET.CommandText = select_query
                MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

                Try
                    atemptable.Load(MyAircraftReader_JETNET)
                Catch constrExc As System.Data.ConstraintException
                    Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                    ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
                End Try

            End If


            If atemptable.Rows.Count > 0 Then ' if we didnt find any, replace the tbd, see if we do 
                ' then dont worry 
            Else
                ' if we didnt find anything, make sure we dont for the different source.  

                MyAircraftReader_JETNET.Close()

                    select_query = " select top 1 publist_ac_id from Publication_Listing with (NOLOCK) "

                    select_query &= " where publist_url = '" & Trim(pub_url) & "' "
                    select_query &= " and publist_entry_date >= '" & newdate.Year & "-" & newdate.Month & "-" & newdate.Day & "' "


                    If Trim(temp_publog_source) <> "" Then
                        select_query &= " and publist_source = '" & temp_publog_source & "' "
                    End If
                    select_query &= " order by publist_id desc "

                    MySqlCommand_JETNET.CommandText = select_query
                    MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

                    Try
                        atemptable.Load(MyAircraftReader_JETNET)
                    Catch constrExc As System.Data.ConstraintException
                        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
                    End Try

                    If atemptable.Rows.Count > 0 Then
                        select_query = select_query
                        used_url = True
                    End If
                End If



            ' if its controller
            If temp_publog_source = 2 And atemptable.Rows.Count = 0 Then
                ' if there is no current record , but there was a record over 2 weeks ago that looks the same, then skip cause it hasnt changed

                select_query = " select publist_ac_id from Publication_Listing with (NOLOCK) "

                select_query &= " where REPLACE(REPLACE(publist_original_desc, char(10), ''), char(13), '')  = '" & Trim(acpub_original_name) & "' "


                '  publist_ac_id = " & ac_id & ""
                newdate = DateAdd(DateInterval.Day, -14, Now())

                select_query &= " and publist_entry_date <= '" & newdate.Year & "-" & newdate.Month & "-" & newdate.Day & "' "
                select_query &= " and publist_url = '" & Trim(pub_url) & "' "
                select_query &= " and publist_source = '" & temp_publog_source & "' "
                select_query &= " and publist_status not in ('O', 'I')  "
                select_query &= " and publist_description = '" & Left(pub_desc, 799) & "' "
                select_query &= " and publist_seller_info = '" & Left(pub_seller_info, 799) & "' " ' added MSw - if its not seller info matching, we should re-open 


                select_query = select_query

                MySqlCommand_JETNET.CommandText = select_query
                MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

                Try
                    atemptable.Load(MyAircraftReader_JETNET)
                Catch constrExc As System.Data.ConstraintException
                    Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                    ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
                End Try


                If atemptable.Rows.Count > 0 Then
                    skip_insert_update = True
                Else
                    skip_insert_update = skip_insert_update
                End If
            End If


            If skip_insert_update = False Then


                ac_found_count = ac_found_count + 1

                ' will now fail test for bad dates
                If atemptable.Rows.Count = 0 And passed_test = True Then

                    If Trim(pub_ser_no) <> "" Then
                        pub_ser_no = pub_ser_no
                    End If


                    Insert_Query = " INSERT INTO Publication_Listing"
                    Insert_Query &= " (publist_ac_id"
                    Insert_Query &= " ,publist_journ_id"
                    Insert_Query &= " ,publist_source"
                    Insert_Query &= " ,publist_reg_no"
                    Insert_Query &= " ,publist_ser_no"
                    Insert_Query &= " ,publist_description"
                    Insert_Query &= " ,publist_price"
                    Insert_Query &= " ,publist_aftt"
                    Insert_Query &= " ,publist_seller_info"
                    Insert_Query &= " ,publist_picture"
                    Insert_Query &= " ,publist_status"
                    Insert_Query &= " ,publist_url"
                    Insert_Query &= "  ,publist_clear_date"
                    Insert_Query &= "  ,publist_acct_rep"
                    Insert_Query &= "  ,publist_entry_date"
                    Insert_Query &= "  ,publist_update_date"
                    Insert_Query &= "  ,publist_original_desc"
                    Insert_Query &= "  ,publist_latest_change"
                    Insert_Query &= "  ,publist_user_id"
                    Insert_Query &= "  , publist_type "
                    Insert_Query &= "  , publist_comp_id "
                    Insert_Query &= "  ,publist_process_status)"
                    Insert_Query &= " VALUES( "
                    Insert_Query &= " " & ac_id & ""
                    Insert_Query &= ",0"
                    Insert_Query &= ", '" & temp_publog_source & "'"
                    Insert_Query &= ", '" & pub_reg_no & "'"
                    Insert_Query &= ", '" & pub_ser_no & "'"
                    Insert_Query &= ", '" & Left(Replace(pub_desc, "'", ""), 799) & "'"
                    '  Insert_Query &= ", '" & Left(pub_desc, 119) & "'"

                    Insert_Query &= ", '" & pub_price & "'"
                    Insert_Query &= ", '" & pub_aftt & "'"
                    Insert_Query &= ", '" & Replace(pub_seller_info, "'", "") & "'"
                    Insert_Query &= ", '" & pub_picture & "'"
                    Insert_Query &= ", '" & acpub_status & "'"
                    Insert_Query &= ", '" & pub_url & "'"
                    Insert_Query &= ", ''"  'clear date
                    Insert_Query &= ", 'PUB1'"  'acct rep
                    Insert_Query &= ", '" & Date.Now & "'"  'entry date
                    Insert_Query &= ", ''"  'update date
                    Insert_Query &= ", '" & Trim(acpub_original_name) & "'"  'original desc
                    Insert_Query &= ", ''"  'latest change
                    Insert_Query &= ", 'mvit'"  'user id  
                    Insert_Query &= ", 'Aircraft'"  'type 
                    Insert_Query &= ", '" & pub_comp_id & "'"  'comp id

                    Insert_Query &= ", '" & acpub_process_status & "')"


                    ' Response.Write("<Br>" & Insert_Query)
                    Insert_Query = Insert_Query
                    MySqlCommand_JETNET.CommandText = Insert_Query
                    MySqlCommand_JETNET.ExecuteNonQuery()

                    System.Threading.Thread.Sleep(200)
                    acpub_insert_count = acpub_insert_count + 1
                    check_insert_ac_pub = True
                    inserted_any_info = True
                Else



                    '----------- SEE IF AC ID EXISTS IN RECORD ----------------------------- 
                    atemptable.Clear()
                    select_query = select_query
                    select_query = Replace(select_query, "order by publist_id desc", "")

                    MySqlCommand_JETNET.CommandText = select_query & " and publist_ac_id > 0 "
                    MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

                    Try
                        atemptable.Load(MyAircraftReader_JETNET)
                    Catch constrExc As System.Data.ConstraintException
                        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
                    End Try

                    If Not IsNothing(atemptable) Then
                        If atemptable.Rows.Count > 0 Then
                            record_has_ac_id = True
                        End If
                    End If
                    '----------- SEE IF AC ID EXISTS IN RECORD -----------------------------


                    '----------- SEE IF AC ID EXISTS IN RECORD ----------------------------- 
                    atemptable.Clear()
                    select_query = select_query
                    select_query = Replace(select_query, "order by publist_id desc", "")

                    MySqlCommand_JETNET.CommandText = select_query & " and publist_comp_id > 0 "
                    MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

                    Try
                        atemptable.Load(MyAircraftReader_JETNET)
                    Catch constrExc As System.Data.ConstraintException
                        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
                        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
                    End Try

                    If Not IsNothing(atemptable) Then
                        If atemptable.Rows.Count > 0 Then
                            record_has_comp_id = True
                        End If
                    End If
                    '----------- SEE IF AC ID EXISTS IN RECORD -----------------------------





                    '---- SEE IF THE RECORD STILL MATCHES ORIGINAL
                    If Trim(acpub_original_name) <> "" Then

                        Try     '' THE FIRST UPDATE UPDATES ALL RECRODS TO PUB 1 that were CLEARED OR NOTHING TO DO 
                            Insert_Query = ""
                            Insert_Query = " UPDATE Publication_Listing set publist_acct_rep = 'PUB1', "

                            Insert_Query &= " publist_description = '" & Left(Replace(pub_desc, "'", ""), 799) & "',  "
                            Insert_Query &= " publist_price = '" & pub_price & "',  "
                            Insert_Query &= " publist_aftt = '" & pub_aftt & "',  "

                            If record_has_comp_id = False Then
                                Insert_Query &= " publist_comp_id = '" & pub_comp_id & "',  "
                            End If

                            If ac_id > 0 And record_has_ac_id = False Then
                                Insert_Query &= " publist_ser_no = '" & pub_ser_no & "',  "
                                Insert_Query &= " publist_reg_no = '" & pub_reg_no & "',  "
                                Insert_Query &= " publist_ac_id = '" & ac_id & "',  "
                            End If

                            '  Insert_Query &= " publist_acct_rep = 'PUB1',  "

                            Insert_Query &= " publist_seller_info = '" & Replace(pub_seller_info, "'", "") & "',  "
                            Insert_Query &= " publist_picture = '" & pub_picture & "',  "
                            'if we should close or open it based on the data compared to jetnet data, then do it 
                            If acpub_status = "O" Or acpub_status = "N" Then
                                If acpub_status = "O" And Trim(pub_desc) = "" Then 'if its open but we have no description then dont re-open, leave where it was 
                                    acpub_status = acpub_status
                                Else
                                    Insert_Query &= " publist_status = '" & acpub_status & "',  "
                                End If
                            End If

                            Insert_Query &= " publist_entry_date = '" & Date.Now & "',  "
                            Insert_Query &= " publist_process_status = '" & acpub_process_status & "' "

                            'should always update the url in case the site has changed how the urls r structured 
                            If Trim(pub_url) <> "" Then
                                Insert_Query &= ", publist_url = '" & pub_url & "' " ' update the url 
                            End If

                            Insert_Query &= " WHERE (REPLACE(REPLACE(publist_original_desc, char(10), ''), char(13), '')   = '" & Trim(acpub_original_name) & "' "

                            If used_url = True Then
                                Insert_Query &= " or publist_url = '" & Trim(pub_url) & "' "  ' if we had to use url, then put in, but still only update on change 
                            End If
                            Insert_Query &= ") "


                            If Trim(pub_desc) = "" And temp_publog_source = 2 Then
                                temp_publog_source = temp_publog_source
                            Else
                                Insert_Query &= " and publist_description <> '" & Left(pub_desc, 799) & "' "
                            End If


                            Insert_Query &= "  and publist_source = '" & temp_publog_source & "' "
                            ' publist_status = '" & acpub_status & "',  "

                            ' added this in to update only when any of these things has also changed 
                            ' Insert_Query &= "  and (publist_description <> '" & Left(pub_desc, 799) & "' or (publist_aftt <> '" & pub_aftt & "' and publist_research_note not like '%ignore aftt%') or (publist_price <> '" & pub_price & "' and publist_research_note not like '%ignore price%')) "
                            ' Insert_Query &= "  and (publist_description <> '" & Left(pub_desc, 799) & "' or (publist_aftt <> '" & pub_aftt & "' and not (publist_research_note like '%ignore aftt%' and publist_entry_date >= getdate() -7) ) or (publist_price <> '" & pub_price & "'  and not (publist_research_note like '%ignore price%' and publist_entry_date >= getdate() -7))  ) "


                            'so if the description changes then check to see if the aftt has changed or price has changed with no ignore
                            ' added or (publist_seller_info <> '" & Replace(pub_seller_info, "'", "") & "')  in MSW - 10/22/19
                            Insert_Query &= "  and (publist_description <> '" & Left(pub_desc, 799) & "' or (publist_seller_info <> '" & Replace(pub_seller_info, "'", "") & "')  or (publist_aftt <> '" & pub_aftt & "' and not (publist_research_note like '%ignore aftt%' and publist_entry_date >= getdate() -7) ) or (publist_price <> '" & pub_price & "'  and not (publist_research_note like '%ignore price%' and publist_entry_date >= getdate() -7))  ) "
                            ' currently redundany check .. not just make sure no items have been ignored in the last week 
                            Insert_Query &= "  and  not (publist_research_note like '%ignore aftt%' and publist_entry_date >= getdate() -7 ) "
                            Insert_Query &= "  and  not (publist_research_note like '%ignore price%' and publist_entry_date >= getdate() -7 ) "


                            Insert_Query &= "  and  (   (not (publist_research_note like '%no blind pub%' and publist_entry_date >= getdate() - 90))   or  (publist_ser_no <> '" & pub_ser_no & "')  or (publist_reg_no <> '" & pub_reg_no & "')    or (publist_seller_info <> '" & Replace(pub_seller_info, "'", "") & "')    ) "



                            Insert_Query &= " and publist_entry_date >= '" & newdate.Year & "-" & newdate.Month & "-" & newdate.Day & "' "
                            Insert_Query &= " and publist_url = '" & Trim(pub_url) & "' "
                            Insert_Query &= " and publist_status in ('C', 'N','D') "

                            ' dont re-open it if the description has changed and the desription was in the last desription 
                            If Trim(pub_desc) = "" And temp_publog_source = 2 Then
                                temp_publog_source = temp_publog_source ' then dont add in 
                            Else
                                Insert_Query &= " and publist_description not like '%" & Left(Replace(pub_desc, "'", ""), 799) & "%' "
                            End If


                            MySqlCommand_JETNET.CommandText = Insert_Query
                            MySqlCommand_JETNET.ExecuteNonQuery()

                            System.Threading.Thread.Sleep(200)

                        Catch ex As Exception

                        End Try



                        Try '' THE SECOND UPDATE UPDATES ALL RECRODS, LEAVING THE PUB WHERE IT WAS THAT WERE OPEN OR IN PROGRESS
                            Insert_Query = ""
                            Insert_Query = " UPDATE Publication_Listing set "

                            Insert_Query &= " publist_description = '" & Left(Replace(pub_desc, "'", ""), 799) & "',  "
                            Insert_Query &= " publist_price = '" & pub_price & "',  "
                            Insert_Query &= " publist_aftt = '" & pub_aftt & "',  "

                            If record_has_comp_id = False Then
                                Insert_Query &= " publist_comp_id = '" & pub_comp_id & "',  "
                            End If

                            If ac_id > 0 And record_has_ac_id = False Then
                                Insert_Query &= " publist_ser_no = '" & pub_ser_no & "',  "
                                Insert_Query &= " publist_reg_no = '" & pub_reg_no & "',  "
                                Insert_Query &= " publist_ac_id = '" & ac_id & "',  "
                            End If

                            '  Insert_Query &= " publist_acct_rep = 'PUB1',  "

                            Insert_Query &= " publist_seller_info = '" & Replace(pub_seller_info, "'", "") & "',  "
                            Insert_Query &= " publist_picture = '" & pub_picture & "',  "
                            'if we should close or open it based on the data compared to jetnet data, then do it 
                            If acpub_status = "O" Or acpub_status = "N" Then
                                If acpub_status = "O" And Trim(pub_desc) = "" Then 'if its open but we have no description then dont re-open, leave where it was 
                                    acpub_status = acpub_status
                                Else
                                    Insert_Query &= " publist_status = '" & acpub_status & "',  "
                                End If
                            End If

                            Insert_Query &= " publist_process_status = '" & acpub_process_status & "', "
                            Insert_Query &= " publist_entry_date = '" & Date.Now & "'  "


                            'should always update the url in case the site has changed how the urls r structured 
                            If Trim(pub_url) <> "" Then
                                Insert_Query &= ", publist_url = '" & pub_url & "' " ' update the url 
                            End If

                            Insert_Query &= " WHERE (REPLACE(REPLACE(publist_original_desc, char(10), ''), char(13), '')   = '" & Trim(acpub_original_name) & "' "

                            If used_url = True Then
                                Insert_Query &= " or publist_url = '" & Trim(pub_url) & "' "  ' if we had to use url, then put in, but still only update on change 
                            End If
                            Insert_Query &= ") "


                            If Trim(pub_desc) = "" And temp_publog_source = 2 Then
                                temp_publog_source = temp_publog_source ' then dont add in 
                            Else
                                Insert_Query &= "   And publist_description <> '" & Left(pub_desc, 799) & "' "
                            End If

                            Insert_Query &= " And publist_source = '" & temp_publog_source & "' "
                            ' publist_status = '" & acpub_status & "',  "

                            ' added this in to update only when any of these things has also changed 
                            Insert_Query &= "  and (publist_description <> '" & Left(pub_desc, 799) & "' or (publist_aftt <> '" & pub_aftt & "' and not (publist_research_note like '%ignore aftt%' and publist_entry_date >= getdate() -7) ) or (publist_price <> '" & pub_price & "'  and not (publist_research_note like '%ignore price%' and publist_entry_date >= getdate() -7))  ) "
                            Insert_Query &= "  and  not (publist_research_note like '%ignore aftt%' and publist_entry_date >= getdate() -7 ) "
                            Insert_Query &= "  and  not (publist_research_note like '%ignore price%' and publist_entry_date >= getdate() -7 ) "

                            Insert_Query &= "  and  (   (not (publist_research_note like '%no blind pub%' and publist_entry_date >= getdate() - 90))   or  (publist_ser_no <> '" & pub_ser_no & "')  or (publist_reg_no <> '" & pub_reg_no & "')    or (publist_seller_info <> '" & Replace(pub_seller_info, "'", "") & "')    ) "



                            Insert_Query &= " and publist_entry_date >= '" & newdate.Year & "-" & newdate.Month & "-" & newdate.Day & "' "
                            Insert_Query &= " and publist_url = '" & Trim(pub_url) & "' "
                            'Insert_Query &= " and publist_status in ('O', 'I') and publist_acct_rep <> 'PUB1' "   ' changed- MSW - 10/6/2020 - from patty - dont update status on In progress
                            Insert_Query &= " and publist_status in ('O') "    ' and publist_acct_rep <> 'PUB1' moved to below 
                            Insert_Query &= " and publist_acct_rep not in ('SPEC','NEW','CUST','PUB1') "   ' added msw - 10/7/2020 
                            ' so that open items, that arent on spec, new or cust, or pub 1 will be updated 

                            ' dont re-open it if the description has changed and the desription was in the last desription 
                            If Trim(pub_desc) = "" And temp_publog_source = 2 Then
                                temp_publog_source = temp_publog_source ' then dont add in 
                            Else
                                Insert_Query &= " and publist_description not like '%" & Left(Replace(pub_desc, "'", ""), 799) & "%' "
                            End If
                            MySqlCommand_JETNET.CommandText = Insert_Query
                            MySqlCommand_JETNET.ExecuteNonQuery()

                            System.Threading.Thread.Sleep(200)

                        Catch ex As Exception

                        End Try


                    End If

                    Try
                        ''do all in progress, but dont update the status 
                        Insert_Query = Replace(Insert_Query, "publist_status = '" & acpub_status & "',", " ")
                        Insert_Query = Replace(Insert_Query, "and publist_status in ('O')", " and publist_status in ('I')")

                        MySqlCommand_JETNET.CommandText = Insert_Query
                        MySqlCommand_JETNET.ExecuteNonQuery()


                        ' Re-run the statement that was run for open items, on in progress items, and dont update the status - MSW - 10/6/2020  
                        ' get rid of update of status 
                        Insert_Query = Replace(Insert_Query, "and publist_status in ('I')", " and publist_status in ('O','I')")
                        Insert_Query = Replace(Insert_Query, "and publist_acct_rep not in ('SPEC','NEW','CUST','PUB1')", "and publist_acct_rep in ('SPEC','NEW','CUST','PUB1')")

                        ' DONT DO THE STATUS ONES 
                        MySqlCommand_JETNET.CommandText = Insert_Query
                        MySqlCommand_JETNET.ExecuteNonQuery()
                    Catch ex As Exception

                    End Try


                    check_insert_ac_pub = check_insert_ac_pub
                        acpub_match_count = acpub_match_count + 1
                    End If
                End If


            ' make sure they clear
            pub_comp_id = 0
            pub_seller_info = ""

        Catch ex As Exception
        Finally
            MyAircraftReader_JETNET.Close()
            MySqlCommand_JETNET.Dispose()
            atemptable = Nothing
        End Try
    End Function
    Public Function get_MMSI_NUMBERS() As String
    get_MMSI_NUMBERS = ""

    Dim we_have_count As Integer = 0
    Dim we_dont_have_count As Integer = 0
    Dim found_news As Integer = 0
    Dim didnt_find_yacht As String = ""
    Dim results As String = ""
    Dim query As String = ""
    Dim temp_link As String = ""
    Dim results_table As New DataTable
    Dim continue_search As Boolean = False
    Dim company_count As Integer = 0
    Dim company_string As String = ""
    Dim this_company_yachts As Integer = 0
    Dim total_found_temp As Integer = 0
    Dim i As Integer = 0
    Dim found_yacht As Integer = 0
    Dim connected_yachts As Integer = 0
    Dim found_companies As Integer = 0
    Dim yt_found_count As Integer = 0
    Dim yt_count As Integer = 0

    Try

      ' MySqlConn_JETNET.ConnectionString = "Data Source=www.jetnetsql2.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
      '' MySqlConn_JETNET2.ConnectionString = "Data Source=www.jetnetsql2.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
      '   MySqlConn_JETNET.ConnectionString = "Data Source=128.1.21.40;Initial Catalog=jetnet_ra_test;Persist Security Info=True;User ID=sa;Password=moejive"
      ' MySqlConn_JETNET.ConnectionString = "Data Source=128.1.21.200;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=sa;Password=moejive"

      ' MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
      MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120


      results_table = GET_YACHT_MMSI_LIST()

      If Not IsNothing(results_table) Then
        If results_table.Rows.Count > 0 Then
          For Each r As DataRow In results_table.Rows
            If Not IsDBNull(r.Item("yt_mmsi_mobile_nbr")) Then

              continue_search = FIND_MARINE_MMSI("http://www.marinetraffic.com/en/ais/details/ships/mmsi:" & r.Item("yt_mmsi_mobile_nbr") & "/")

              If continue_search = True Then
                yt_found_count = yt_found_count + 1

                ' insert statement

              End If
              yt_count = yt_count + 1
            End If
          Next
        End If
      End If


      results = "<br>Yachts With MMSI: " & CStr(yt_count) & "" & Chr(13) & Chr(10)
      results &= "<br>Yachts Connected to Marine Traffic w MMSI: " & CStr(yt_found_count) & "" & Chr(13) & Chr(10)


      Me.text_label.Text = results
      get_MMSI_NUMBERS = results

    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()
    End Try
  End Function


  Public Function get_AC_Details() As String
    get_AC_Details = ""

    Dim we_have_count As Integer = 0
    Dim we_dont_have_count As Integer = 0
    Dim found_news As Integer = 0
    Dim didnt_find_yacht As String = ""
    Dim results As String = ""
    Dim query As String = ""
    Dim temp_link As String = ""
    Dim results_table As New DataTable
    Dim continue_search As Boolean = False
    Dim company_count As Integer = 0
    Dim company_string As String = ""
    Dim this_company_yachts As Integer = 0
    Dim total_found_temp As Integer = 0
    Dim i As Integer = 0
    Dim found_yacht As Integer = 0
    Dim connected_yachts As Integer = 0
    Dim found_companies As Integer = 0
    Dim select_query As String = ""
    Dim update_query As String = ""
    Dim split_array(10) As String
    Dim z As Integer = 0
    Dim hour_inspections(20) As Integer


    Try

      hour_inspections(0) = 100
      hour_inspections(1) = 150
      hour_inspections(2) = 200
      hour_inspections(3) = 300
      hour_inspections(4) = 400
      hour_inspections(5) = 100
      hour_inspections(6) = 600
      hour_inspections(7) = 1000
      hour_inspections(8) = 2000
      hour_inspections(9) = 1200
      hour_inspections(10) = 1600
      hour_inspections(11) = 4800
      hour_inspections(12) = 5000
      hour_inspections(13) = 10000
      hour_inspections(14) = 500
      hour_inspections(15) = 1800
      hour_inspections(16) = 3000
      hour_inspections(17) = 1500


      ' MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN
      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120



      If Trim(Request("clear")) = "Y" Then
        ' skip the rest
      Else
        'select_query = " select * from Maintenance_Item with (NOLOCK) " 

        'MySqlCommand_JETNET.CommandText = select_query
        'MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

        'Try
        '  results_table.Load(MyAircraftReader_JETNET)
        'Catch constrExc As System.Data.ConstraintException
        '  Dim rowsErr As System.Data.DataRow() = results_table.GetErrors()
        '  ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
        'End Try



        'If Not IsNothing(results_table) Then
        '  If results_table.Rows.Count > 0 Then
        '    For Each r As DataRow In results_table.Rows

        '    Next
        '  End If
        'End If
        ' results_table.Rows.Clear()

        ' STILL GET THE DETAILS FROM OUTSIDE
        results_table = Get_AC_Maint_Details()


        If Not IsNothing(results_table) Then
          If results_table.Rows.Count > 0 Then
            For Each r As DataRow In results_table.Rows
              is_due_date = False
              is_found = False
              found_36_96 = ""   ' clear this for every ac
              temp_ac_id = 0
              temp_ac_id = Trim(r.Item("ac_id"))
              temp_details = Trim(r.Item("adet_data_description"))
              temp_amod_id = Trim(r.Item("ac_amod_id"))
              If Not IsDBNull(r.Item("ac_forsale_flag")) Then
                temp_sale_flag = Trim(r.Item("ac_forsale_flag"))
              Else
                temp_sale_flag = "N"
              End If
              found_any_info = False
              inserted_any_info = False

              ReDim split_array(10)

              For z = 0 To 10
                split_array(z) = ""
              Next


              split_array = Split(temp_details, ".")


              For z = 0 To split_array.Length - 1
                temp_details = split_array(z) & "." ' put the period back in 


                If only_run_this_section = True Then




                Else

                  If find_details_in_string(temp_details, "Certificate of Airworthiness", "c of a", "certificate of airworthiness", "certificate of airworthiness", "Certificate of Airwothiness", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "Entered into Service", "entered into service", "placed into service", "service date", "entry into service", "aircraft entered service", "", "", False) Then

                  End If


                  '36-Month Inspection
                  If find_details_in_string(temp_details, "36-Month Inspection", "36-month inspection", "36 month inspection", "36-Month Items", "36-Month Check", "", "", "", False) Then

                  End If

                  '192-Month Inspection
                  If find_details_in_string(temp_details, "192-Month Inspection", "192-Month Inspection", "192 month inspection", "192 month items", "192-Month Items", "192-Month Due items", "192-Month Airframe inspection", "", False) Then

                  End If

                  '96-Month Inspection
                  If find_details_in_string(temp_details, "96-Month Inspection", "96-Month Inspection", "96 month inspection", "96 month items", "96-Month Items", "96-Month Due items", "", "", False) Then

                  End If

                  '192-Month Inspection
                  If find_details_in_string(temp_details, "192-Month Inspection", "192-Month Inspection", "192 month inspection", "192 month items", "192-Month Items", "192-Month Due items", "192-Month Airframe inspection", "", False) Then

                  End If

                  '6-Month Inspection
                  If find_details_in_string(temp_details, "6-Month Inspection", "6-month inspection", "6 month inspection", "6-Month Check", "", "", "", "", False) Then

                  End If


                  '12-Month Inspection
                  If find_details_in_string(temp_details, "12-Month Inspection", "12-month inspection", "12 month inspection", "12-Month Check", "12-Month calendar inspection", "", "", "", False) Then

                  End If


                  '24-Month Inspection
                  If find_details_in_string(temp_details, "24-Month Inspection", "24-month inspection", "24 month inspection", "24 month items", "24-Month Items", "24-Month due items", "24-Month Check", "", False) Then

                  End If



                  '48-Month Inspection
                  If find_details_in_string(temp_details, "48-Month Inspection", "48-month inspection", "48 month inspection", "48-Month Items", "48-Month Check", "48-Month c/w", "", "", False) Then

                  End If


                  '60-Month Inspection
                  If find_details_in_string(temp_details, "60-Month Inspection", "60-month inspection", "60 month inspection", "", "", "", "", "", False) Then

                  End If


                  '72-Month Inspection
                  If find_details_in_string(temp_details, "72-Month Inspection", "72-month inspection", "72 month inspection", "72 -Month inspection", "72-Month", "72-Month items", "6-Year inspection", "", False) Then

                  End If





                  '120-Month Inspection    
                  If find_details_in_string(temp_details, "120-Month Inspection", "120-month inspection", "120 month inspection", "10-Year inspection", "10 Year inspection", "120-Month Check", "", "", False) Then

                  End If

                  '120-Month Inspection
                  If find_details_in_string(temp_details, "180-Month Inspection", "180-Month Inspection", "180 month inspection", "180-Month Check", "180 Month Check", "", "", "", False) Then

                  End If

                  '240-Month Inspection
                  If find_details_in_string(temp_details, "240-Month Inspection", "240-month inspection", "240 month inspection", "240-Month Check", "240 Month Check", "20-Year", "", "", False) Then

                  End If



                  '18-Month Inspection
                  If find_details_in_string(temp_details, "18-Month Inspection", "18-month inspection", "18 month insp", "18-month insp.", "18-month c/w", "", "", "", False) Then

                  End If


                  '12-year or 144 month
                  If find_details_in_string(temp_details, "12-Year Inspection", "12-Year inspection", "", "", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "144-Month Inspection", "144-Month Inspection", "144- Month items", "", "", "", "", "", False) Then

                  End If


                  '--------------------------------------------
                  '12/24-Month items
                  If find_details_in_string(temp_details, "12/24-Month items", "12/24-Month items", "12 & 24-Month inspections", "12-Month & 24-Month inspections", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "12/24/36/48/72/144-Month inspections", "12/24/36/48/72/144-Month inspections", "", "", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "12/24/36/72/144-Month inspections", "12/24/36/72/144-Month inspections", "", "", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "12/48/144-Month inspections", "12/48/144-Month inspections", "", "", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "72/144-Month inspections", "72/144-Month inspections", "", "", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "12/24/48/180-Month inspections", "12/24/48/180-Month inspections", "", "", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "12, 24, 48-Month inspection", "12, 24, 48-Month inspection", "", "", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "12/24/48 & 96-Month inspection", "12/24/48 & 96-Month inspection", "12/24/48 & 96-Month inspections", "12/24/48 & 96-Month inspections", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "180/240-Month Checks", "180/240-Month Checks", "", "", "", "", "", "", False) Then

                  End If



                  For i = 1 To 21
                    'Phase 1 Inspection
                    If find_details_in_string(temp_details, "Phase " & Trim(i) & " Inspection", "Phase " & Trim(i) & " Inspection", "Phase-" & Trim(i) & "  Inspection", "", "", "", "", "", False) Then

                    End If
                  Next

                  For i = 0 To 17
                    If find_details_in_string(temp_details, "" & Trim(hour_inspections(i)) & "-Hour Inspection", "" & Trim(hour_inspections(i)) & " Hour Inspection", "" & Trim(FormatNumber(hour_inspections(i), 0)) & " Hour Inspection", "" & Trim(hour_inspections(i)) & " Hour", "" & Trim(hour_inspections(i)) & "-Hour Inspection", "" & Trim(hour_inspections(i)) & "-Hour", "", "", False) Then

                    End If
                  Next

                  'Phase 1 - 4 Inspection
                  If find_details_in_string(temp_details, "Phase 1 - 4 Inspection", "Phase 1 - 4 inspection", "Phase 1-4 inspection", "Phase 1, 2, 3 & 4 inspections", "", "", "", "", False) Then

                  End If

                  'Phase 1 - 5 Inspection
                  If find_details_in_string(temp_details, "Phase 1 - 5 Inspection", "Phase 1 - 5 inspection", "Phase 1-5 inspection", "Phase 1- 5 inspections", "Phase 1-5 inspections", "Phase 1 - 5 inspections", "Phases I - V inspections", "Phase 1 -5 inspections", False) Then

                  End If


                  'Phase 3 - 5 inspections
                  If find_details_in_string(temp_details, "Phase 3 - 5 inspections", "Phase 3 - 5 inspections", "", "", "", "", "", "", False) Then

                  End If


                  'Phase 3 - 5 inspections
                  If find_details_in_string(temp_details, "Phase 1 & 2 inspections", "Phase 1 & 2 inspections", "Phase 1 & 2 inspection", "Phase 1 - 2 inspections", "", "", "", "", False) Then

                  End If

                  'Phase 3 - 5 inspections
                  If find_details_in_string(temp_details, "Phase 3 & 4 inspections", "Phase 3 & 4 inspections", "Phase 3 & 4 inspection", "Phase 3 - 4 inspections", "", "", "", "", False) Then

                  End If
                  ' c/w 10/00.  





                  If find_details_in_string(temp_details, "Phase A Inspection", "Phase A Inspection", "Phase-A Inspection", "", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "Phase B Inspection", "Phase B Inspection", "Phase-B Inspection", "", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "Phase C Inspection", "Phase C Inspection", "Phase-C Inspection", "", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "Phase D Inspection", "Phase D Inspection", "Phase-D Inspection", "", "", "", "", "", False) Then

                  End If


                  ' GOING TO BE CONVERTED TO LANDING GEAR INSPECTION
                  If find_details_in_string(temp_details, "Landing Gear inspection", "Landing Gear inspection", "Landing Gear Detailed Inspection", "Landing Gear Corrosion Inspection", "Gear inspection", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "Landing Gear Overhaul", "Landing Gear Overhauld", "Landing Gear Overhauled", "Landing Gear Overhaul", "Gear Overhaul", "Landing Gear OH", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "Engine Overhaul", "Engine Overhaul", "Engine Overhaul/Hot Section inspections", "", "", "", "", "", False) Then

                  End If



                  If find_details_in_string(temp_details, "Pre-Purchase inspection", "Pre-Purchase inspection", "Pre-Purchase insp", "Pre Purchase insp", "Pre Purchase inspection", "Pre-buy inspection", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "Airworthiness Review Certificate", " ARC ", "", "", "", "", "", "", False) Then

                  End If



                  If find_details_in_string(temp_details, "Annual Inspection", "Annual Inspection", "", "", "", "", "", "", False) Then

                  End If



                  If find_details_in_string(temp_details, "2A Inspection", "2A Inspection", "2a Inspection", "2A-Inspection", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "2A+ Inspection", "2A+ Inspection", "2a+ Inspection", "2a + Inspection", "2A + Inspection", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "2B Inspection", "2B Inspection", "2b Inspection", "2B-Inspection", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "Engine Boroscope Inspection", "Engine Boroscope Inspection", "Engine Borescope Inspection", "Engine Boroscope", "Engine Borescope", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "3B Inspection", "3B Inspection", "3b Inspection", "3B-Inspection", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "4A Inspection", "4A Inspection", "4a Inspection", "4A-Inspection", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "6000-Landings Inspection", "6000 Landings Inspection", "6000 Landings", "6000 Landing Gear Inpection", "6000 Landings Inpection", "6000-Landings", "6000-Landings Inpection", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "A-Check", "A Check", "A-Check", "a Check", "a-Check", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "B-Check", "B Check", "B-Check", "b Check", "b-Check", "", "", "", False) Then

                  End If



                  If find_details_in_string(temp_details, "C-Check", "C Check", "C-Check", "c Check", "c-Check", "", "", "", False) Then

                  End If



                  If find_details_in_string(temp_details, "D-Check", "D Check", "D-Check", "d Check", "d-Check", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "Beech Status Inspection", "Beech Status Inspection", "Beech Status", "", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "Chapter 5 Inspection", "Chapter 5 Inspection", "Chapter 5 Corrosion Inspection", "Chapter-5 Inspection", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "Event 2 Inspection", "Event 2 Inspection", "Event II Inspection", "Event 2", "Event II", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "Event 1 Inspection", "Event 1 Inspection", "Event I Inspection", "Event 1", "Event I", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "Life Limit Inspection", "Life Limit Inspection", "Life Limits", "life limited items", "life limited parts", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "Major Corrosion Inspection", "MCI Inspection", "MCI", "MCI Inspections", "Major Corrosion Inspection", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "Tank & Plank Inspection", "Tank & Plank Inspection", "Tank & Plank Inspection", "Tank and Plank Inspection", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "Wing Spar Inspection", "Wing Spar Inspection", "Wing Spar", "", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "2BInspection", "2BInspection", "2bInspection", "", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "2BInspection", "2BInspection", "2bInspection", "", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "2BInspection", "2BInspection", "2bInspection", "", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "Engine Hot Section Inspection", "Engine Hot Section Inspection", "Hot Section Inspection", "", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "APU Inspection", "APU Inspection", "APU Overhaul", "", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "Engine Inspection", "Engine Inspection", "", "", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "Right Engine Hot Section Inspection", "Right engine Hot Section inspection", "Right Engine Hot Section inspection", "Right engine Hot Section inspection", "", "", "", "", False) Then

                  End If







                  '---- DOCUMENT INSPECTIONS

                  For i = 1 To 18
                    If find_details_in_string(temp_details, "Document " & i & " Inspection", "Document " & i & " Inspection", "Document " & i & " Inspections", "Document " & i & " Insp", "Doc " & i & " inspection", "Document " & i & "", "", "", False) Then

                    End If

                  Next

                  If find_details_in_string(temp_details, "Document 22 Inspection", "Document 22 Inspection", "Document 22 Inspections", "Document 22 Insp", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "Document 36 Inspection", "Document 36 Inspection", "Document 36 Inspections", "Document 36 Insp", "", "", "", "", False) Then

                  End If


                  If find_details_in_string(temp_details, "Documents 8 & 10", "Documents 8 & 10", "", "", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "Document 1, 2, 4, 10, 11, 12, 13, 17, 36 inspections", "Document 1, 2, 4, 10, 11, 12, 13, 17, 36 inspections", "", "", "", "", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "D1 - D12 inspections", "D1 - D12 inspections", "D1 - D12 inspection", "", "", "", "", "", False) Then

                  End If



                  If find_details_in_string(temp_details, "12/36/72-Month Inspection", "12/36/72-Month Inspection", "12/36/72 Month Inspection", "12/36/72-month onspection", "12/36/72-Month", "12/36/72 month", "", "", False) Then

                  End If

                  If find_details_in_string(temp_details, "12/24/36/48/96-Month Inspection", "12/24/36/48/96-Month Inspection", "12/24/36/48/96 Month Inspection", "12/24/36/48/96 month inspection", "12/24/36/48/96 month", "12/24/36/48/96-month", "12/24/36/48/96 Month inspection", "", False) Then

                  End If

                  For i = 1 To 12
                    If find_details_in_string(temp_details, Trim(i) & "C Inspection", Trim(i) & "-C Inspection", Trim(i) & "-C inspection", Trim(i) & "C Inspection", Trim(i) & "C inspection", "", "", "", False) Then

                    End If
                  Next



                End If  ' for only run this section 





              Next


              ' suggested by dad
              ' 12-Year corrosion inspection c/w 07/2016 by gulfstream, Savannah, GA.
              ' Fresh 12C-inspetion c/w 06/2012 by Gulfstream, Buton   - g-550 - 5019
              '48/96-Month c/w 06/13 by Gulfstream; Savannah, GA.


              ' can currently do-----------------------   
              '12-Year Major inspection c/w 08/03/07
              '30-Month inspection due 07/17


              ' future inspections---------------
              'Document 1, 3, 4, 8, 17, 28, 42, 47, 50 inspections & 12-Month items c/w 06/15 by Mather Aviation 
              '18-Month & Wing Center Section inspections c/w 03/12. 
              '300-Hour Check c/w at 2623TT; due at 2923TT. 
              'Major maintenance c/w by Cessna as reported 09/16/15.   
              ' 5000-Landings inspection c/w 08/03 by Duncan Aviation.   
              ' E1 - E12 inspections c/w 05/08 at 1215TT; due 05/09
              ' 12-Month F1 - F12 inspections c/w 10/30/09.  
              '24-Year Structural X-Ray inspections c/w 01/07.
              '96-Month X-ray inspections due 05/13
              'E inspection c/w 04/13 by Tag Engineering.
              'Group E inspections due 05/10
              '   Group F & G inspections due 05/11
              '15-Year Structural inspection due 05/14.  
              '24-Month Air Data & Transponder inspections due 05/11.
              '48-Month & 16-Year inspections c/w 05/12
              'E-Check due 03/11.  
              'F-Check due 12/11. 
              'C1 - C12 Inspection c/w 10/13. 
              '12-Month / E1- E12 Inspection due 04/15. 
              ' E1 - E12 inspections c/w 07/15.  
              'F1 - F12 inspections c/w 07/15.  
              'APU 125/250/500/1000/1500-Hour inspections c/w 07/15.
              '12/24/48 & 96-Month Structural inspections
              '18-Month & X-Ray inspections c/w as reported 08/18/04.
              'Engine Hot Section inspections due LE/RE: 05/16 / 05/16. 
              'Fuselage Penetration inspection c/w 06/04/09 at 6239TT & 3613 landings; due 06/30/21. 
              'Elevator Ballast inspection c/w 10/08/09 at 4469T & 2686 landings; due 10/31/13.  
              '5000-Landings inspection c/w 08/03 by Duncan Aviation
              'B1-B12 inspections due at 11160TT. 
              'C1-C12 inspections due at 11817TT. 
              'D1-D12 inspections due at 11817TT. E1-E12, 
              'E1 - E12 inspections c/w 09/08. 
              'F1-F12 inspections c/w as reported 03/09/16.
              'G inspection c/w 03/26/13; due 03/26/17  
              '48-Month inspection c/w in 1995.  
              '24-Month Progressive maintenance program. 
              'Fresh 15/30/60/180-Month inspections c/w by Penestar as reported 02/06/15  
              '72-Month inspection due end of 2015. 



              ' unsure of what to do with -------------
              '12/24-Month Document inspections c/w 07/15. 
              'E & F inspections c/w 10/30/09.
              ' D inspectio nc/w at 3166TT; due 6366TT
              'F1- F12 Inspection c/s 07/16.
              '48-Month /G Inspection due 09/18. 
              '48-Month & 800-Hour inspections c/w 02/14.  
              '48-Month & 8-Year X-Ray inspections c/w 11/98 by Midcoast.
              '48-Month & X-ray inspections c/w 04/07. 
              '12/24/27-Month inspections c/w as reported 09/24/13 
              '12-Month & 750-hour/12-Month Engine inspections c/w 07/15.  
              '1A, 1A+ Basic & 200/600-Hour Engine inspections c/w 02/15.  





              If found_any_info = False Then
                Response.Write("<br>---------NOTHING FOUND (AC: " & temp_ac_id & "): " & temp_details)
              End If





              'If found_any_info = True Then
              '  If Trim(temp_sale_flag) = "Y" Then
              '    ' g-550 (278), pilatus pc12ng (659), citation bravo (36), xls (287), kind air b200 (207), baron g58 (622), embrear phenom 100 (654), augsta westland aw139 (408)
              '    If temp_amod_id = 278 Or temp_amod_id = 659 Or temp_amod_id = 36 Or temp_amod_id = 287 Or temp_amod_id = 207 Or temp_amod_id = 622 Or temp_amod_id = 654 Or temp_amod_id = 408 Then
              '      clear_action_query &= "update Aircraft set ac_action_date = NULL where ac_journ_id = 0 and ac_id = " & temp_ac_id & "; "
              '    End If
              '  End If
              'End If







              ac_count = ac_count + 1
            Next
          End If
        End If



        '-- FIX DUE DATES
        update_query = ""
        update_query = " update Aircraft_Maintenance set acmaint_due_date = DATEADD(m,(select distinct mitem_duration from Maintenance_Item with (NOLOCK)"
        update_query &= " where mitem_name=acmaint_name),acmaint_complied_date)"
        update_query &= " from Aircraft_Maintenance with (NOLOCK)"
        update_query &= " where acmaint_due_date is NULL and acmaint_name in (select distinct mitem_name from Maintenance_Item with (NOLOCK)"
        update_query &= " where mitem_duration > 0)"
        'update_query &= " and acmaint_notes not like '%as reported%' "


        MySqlCommand_JETNET.CommandText = update_query
        MySqlCommand_JETNET.ExecuteNonQuery()
        MySqlCommand_JETNET.Dispose()

        '-- FIX COMPLIED WITH DATES
        update_query = ""
        update_query = " update Aircraft_Maintenance set acmaint_complied_date = DATEADD(m,-(select distinct mitem_duration from Maintenance_Item with (NOLOCK)"
        update_query &= " where mitem_name=acmaint_name),acmaint_due_date)"
        update_query &= " from Aircraft_Maintenance with (NOLOCK)"
        update_query &= " where acmaint_complied_date is NULL and acmaint_name in (select distinct mitem_name from Maintenance_Item with (NOLOCK)"
        update_query &= " where mitem_duration > 0)"
        ' update_query &= " and acmaint_notes not like '%as reported%' "

        MySqlCommand_JETNET.CommandText = update_query
        MySqlCommand_JETNET.ExecuteNonQuery()
        MySqlCommand_JETNET.Dispose()



        'update_query = ""

        ''update_query &= " update Aircraft_Maintenance set acmaint_name='100-Hour Inspection' where acmaint_name='100 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='1000-Hour Inspection' where acmaint_name='1000 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='10000-Hour Inspection' where acmaint_name='10000 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='1200-Hour Inspection' where acmaint_name='1200 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='150-Hour Inspection' where acmaint_name='150 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='1600-Hour Inspection' where acmaint_name='1600 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='200-Hour Inspection' where acmaint_name='200 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='2000-Hour Inspection' where acmaint_name='2000 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='300-Hour Inspection' where acmaint_name='300 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='400-Hour Inspection' where acmaint_name='400 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='4800-Hour Inspection' where acmaint_name='4800 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='500-Hour Inspection' where acmaint_name='500 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='5000-Hour Inspection' where acmaint_name='5000 Hour Inspection';"
        ''update_query &= " update Aircraft_Maintenance set acmaint_name='600-Hour Inspection' where acmaint_name='600 Hour Inspection';"
        ''  update_query &= " update Aircraft_Maintenance set acmaint_name='6000-Landings Inspection' where acmaint_name='6000 Landings Inspection';"


        'update_query &= " update Aircraft_Maintenance set acmaint_name='A-Check' where acmaint_name='A Check';"
        'update_query &= " update Aircraft_Maintenance set acmaint_name='B-Check' where acmaint_name='B Check';"
        'update_query &= " update Aircraft_Maintenance set acmaint_name='C-Check' where acmaint_name='C Check';"
        'update_query &= " update Aircraft_Maintenance set acmaint_name='D-Check' where acmaint_name='D Check';"
        'update_query &= " update Aircraft_Maintenance set acmaint_name='Major Corrosion Inspection' where acmaint_name='MCI Inspection';"
      End If





      'If only_run_this_section = True Then
      'Else
      '  ' If Trim(Request("clear")) = "Y" Then
      '  clear_action_query = clear_action_query
      '  clear_action_query &= "update Aircraft set ac_action_date = NULL where ac_journ_id = 0 and ac_forsale_flag = 'Y' and ac_amod_id in (278,659,36,287,207,622,654,408); "

      '  MySqlCommand_JETNET.CommandText = clear_action_query
      '  MySqlCommand_JETNET.ExecuteNonQuery()
      '  MySqlCommand_JETNET.Dispose()
      '  '   End If
      'End If


      Response.Write(mis_string)


      results = "<br>Ac With Details: " & CStr(ac_count) & "" & Chr(13) & Chr(10)
      results &= "<br>Details Accessable: " & CStr(ac_found_count) & "" & Chr(13) & Chr(10)
      results &= "<br>Details Inserted: " & CStr(ac_insert_count) & "" & Chr(13) & Chr(10)
      results &= "<br>Details Left to Convert: " & CStr(error_count) & "" & Chr(13) & Chr(10)
      results &= "<br>Percentage Converted: " & CStr(CDbl((100 - ((error_count / (ac_found_count + error_count))) * 100))) & "" & Chr(13) & Chr(10)
      results &= "<br>Details With TT: " & CStr(TT_Count) & "" & Chr(13) & Chr(10)
      results &= "<br>TT Records Found: " & CStr(afft_found) & "" & Chr(13) & Chr(10)


      Me.text_label.Text = results
      get_AC_Details = results

    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()
    End Try
  End Function
  'Public Function cut_out_entered_into_service(ByRef r As DataRow)

  '  Try 



  '    If Not IsDBNull(r.Item("into_service")) Then

  '      original_temp_details = temp_details
  '      temp_details = LCase(temp_details)

  '      temp_spot = InStr(temp_details, "entered into service")

  '      If temp_spot = 0 Then

  '        temp_spot = InStr(temp_details, "placed into service")
  '        If temp_spot > 0 Then
  '          length_plus_space = 18
  '        Else
  '          temp_spot = InStr(temp_details, "service date")
  '          If temp_spot > 0 Then
  '            length_plus_space = 11
  '          Else
  '            temp_spot = InStr(temp_details, "entry into service")
  '            If temp_spot > 0 Then
  '              length_plus_space = 17
  '            End If
  '          End If
  '        End If

  '      Else
  '        length_plus_space = 19
  '      End If

  '      temp_spot = temp_spot
  '      pass_in_string_to_strip_and_insert("Entered into Service")
  '    End If
  '  Catch ex As Exception

  '  End Try
  ' End Function
  Public Function find_details_in_string(ByVal temp_details As String, ByVal type_to_insert As String, ByVal term1 As String, ByVal term2 As String, ByVal term3 As String, ByVal term4 As String, ByVal term5 As String, ByVal term6 As String, ByVal term7 As String, ByVal try_le_re As Boolean) As Boolean

    Try
      before_text = ""

      term1 = LCase(Trim(term1))
      term2 = LCase(Trim(term2))
      term3 = LCase(Trim(term3))
      term4 = LCase(Trim(term4))
      term5 = LCase(Trim(term5))
      term6 = LCase(Trim(term6))
      term7 = LCase(Trim(term7))

      find_details_in_string = False

      original_temp_details = temp_details

      temp_spot = InStr(LCase(temp_details), term1)

      If temp_spot = 0 Or Trim(term1) = "" Then
        temp_spot = InStr(LCase(temp_details), term2)
        If temp_spot = 0 Or Trim(term2) = "" Then
          temp_spot = InStr(LCase(temp_details), term3)
          If temp_spot = 0 Or Trim(term3) = "" Then
            temp_spot = InStr(LCase(temp_details), term4)
            If temp_spot = 0 Or Trim(term4) = "" Then

              ' MAKE SURE THAT TERM 5 - MUST START WITH TERM 5, or else not a match 
              temp_spot = InStr(LCase(temp_details), term5)
              If temp_spot = 0 Or Trim(term5) = "" Then

                temp_spot = InStr(LCase(temp_details), term6)
                If temp_spot = 0 Or Trim(term6) = "" Then

                  temp_spot = InStr(LCase(temp_details), term7)
                  If temp_spot = 0 Or Trim(term7) = "" Then
                    details_found_in_string = False
                  Else
                    length_plus_space = Len(Trim(term7))
                    details_found_in_string = True
                  End If

                Else
                  length_plus_space = Len(Trim(term6))
                  details_found_in_string = True
                End If


              Else
                length_plus_space = Len(Trim(term5))
                details_found_in_string = True
              End If



            Else
              length_plus_space = Len(Trim(term4))
              details_found_in_string = True
            End If
          Else
            length_plus_space = Len(Trim(term3))
            details_found_in_string = True
          End If
        Else
          length_plus_space = Len(Trim(term2))
          details_found_in_string = True
        End If
      Else
        length_plus_space = Len(Trim(term1))
        details_found_in_string = True
      End If


      If temp_spot > 0 And details_found_in_string = True Then
        find_details_in_string = True
        found_any_info = True
        Call pass_in_string_to_strip_and_insert(type_to_insert, try_le_re, details_found_in_string, found_any_info)

        If details_found_in_string = True Then
          find_details_in_string = True
        End If
      End If

    Catch ex As Exception
      'Response.Write("<Br>ERROR: " & ex.ToString)
    End Try
  End Function

  Public Sub pass_in_string_to_strip_and_insert(ByVal type_to_insert As String, ByVal try_le_re As Boolean, ByRef details_found_in_string As Boolean, ByRef found_any_info As Boolean)

    Try

      found_le_re = False
      note_string = ""
      left_engine_string = ""
      right_engine_string = ""
      le_re_complied_or_due = ""
      left_aftt = 0
      right_aftt = 0
      found_possible_mismatch = False



      If temp_spot > 0 Then
        before_text = Left(temp_details, temp_spot - 1)
      End If

      If Trim(before_text) <> "" Then
        If Right(before_text, 1) <> " " Then
          If Right(before_text, 1) <> "/" Then
            found_possible_mismatch = True
            mis_count = mis_count + 1
            mis_string = mis_string & "<br/>" & before_text & "...." & temp_details & "..." & type_to_insert
          End If
        End If
      End If

      ' skip it.. if there is a mis-match found
      If found_possible_mismatch = True Then
        details_found_in_string = False ' then skip the rest 
        found_any_info = True
      Else
        If Trim(before_text) <> "" And Len(Trim(before_text)) > 3 Then
          If Trim(UCase(before_text)) = "AIRCRAFT" Then
            before_text = ""
          ElseIf Trim(UCase(before_text)) = "IN" Then
            before_text = ""
          Else
            before_text = before_text
          End If
        Else
          before_text = ""
        End If

        temp_string = Right(temp_details, Len(temp_details) - temp_spot - length_plus_space)
        note_string_to_replace = temp_string


        temp_spot = InStr(temp_string, ".")

        If temp_ac_id = 122924 Then
          temp_ac_id = temp_ac_id
        End If

        If temp_spot > 0 Then

        Else
          temp_spot = InStr(temp_string, ",")
        End If


        '---------------THIS SECTION IS ONLY USED FOR THE LEFT AND RIGHT ENGINE -----------------------------------------------------
        If try_le_re = True Then
          If InStr(temp_string, "LE/RE:") > 0 Then
            If temp_spot > 20 Then
              found_le_re = True
              Response.Write("<br>LE/RE: " & temp_string)

              le_re_complied_or_due = Left(Trim(temp_string), InStr(Trim(temp_string), "LE/RE:"))
              If InStr(le_re_complied_or_due, "cw") > 0 Then
                le_re_complied_or_due = "cw"
              ElseIf InStr(le_re_complied_or_due, "c/w") > 0 Then
                le_re_complied_or_due = "cw"
              ElseIf InStr(le_re_complied_or_due, "due") > 0 Then
                le_re_complied_or_due = "due"
              Else
                le_re_complied_or_due = "due"
              End If
              temp_string = Right(Trim(temp_string), Len(Trim(temp_string)) - InStr(Trim(temp_string), "LE/RE:") - 6)
              temp_spot = InStr(temp_string, ".")

              If temp_spot > 0 Then

              Else
                temp_spot = InStr(temp_string, ",")
              End If

              'see if there are two numbers
              If InStr(Trim(temp_string), " / ") > 0 Or InStr(Trim(temp_string), "/ ") > 0 Then
                If InStr(Trim(temp_string), " / ") > 0 Then
                  left_engine_string = Left(Trim(temp_string), InStr(Trim(temp_string), " / "))
                  right_engine_string = Right(Trim(temp_string), Len(Trim(temp_string)) - InStr(Trim(temp_string), " / ") - 2)
                Else
                  left_engine_string = Left(Trim(temp_string), InStr(Trim(temp_string), "/ "))
                  right_engine_string = Right(Trim(temp_string), Len(Trim(temp_string)) - InStr(Trim(temp_string), "/ ") - 1)
                End If


                temp_spot = InStr(right_engine_string, ".")

                If temp_spot > 0 Then

                Else
                  temp_spot = InStr(right_engine_string, ",")
                End If
                right_engine_string = Left(Trim(right_engine_string), temp_spot)

                ' MUST PUT IN A SECTION FOR IF THE TT IS ON THE LEFT TO SWITCH - TEMP HOLD
                If InStr(Trim(left_engine_string), "or at") > 0 Then
                  left_aftt = CInt(Trim(Replace(Replace(Replace(Right(Trim(left_engine_string), Len(Trim(left_engine_string)) - (InStr(Trim(left_engine_string), "or at") + 4)), "TT", ""), "or at", ""), "at", "")))
                  left_engine_string = Left(Trim(left_engine_string), InStr(Trim(left_engine_string), "or at") - 1)
                ElseIf InStr(Trim(left_engine_string), "or in") > 0 Then
                  left_aftt = CInt(Trim(Replace(Replace(Replace(Right(Trim(left_engine_string), Len(Trim(left_engine_string)) - (InStr(Trim(left_engine_string), "or in") + 4)), "TT", ""), "or in", ""), "in", "")))
                  left_engine_string = Left(Trim(left_engine_string), InStr(Trim(left_engine_string), "or in") - 1)
                ElseIf InStr(Trim(left_engine_string), " or ") > 0 Then
                  left_aftt = CInt(Trim(Replace(Replace(Right(Trim(left_engine_string), Len(Trim(left_engine_string)) - (InStr(Trim(left_engine_string), " or ") + 4)), "TT", ""), " or ", "")))
                  left_engine_string = Left(Trim(left_engine_string), InStr(Trim(left_engine_string), " or ") - 1)
                End If

                If InStr(Trim(right_engine_string), "or at") > 0 Then
                  right_aftt = CInt(Trim(Replace(Replace(Replace(Right(Trim(right_engine_string), Len(Trim(right_engine_string)) - (InStr(Trim(right_engine_string), "or at") + 4)), "TT", ""), "or at", ""), "at", "")))
                  right_engine_string = Left(Trim(right_engine_string), InStr(Trim(right_engine_string), "or at") - 1)
                ElseIf InStr(Trim(right_engine_string), "or in") > 0 Then
                  right_aftt = CInt(Trim(Replace(Replace(Replace(Right(Trim(right_engine_string), Len(Trim(right_engine_string)) - (InStr(Trim(right_engine_string), "or in") + 4)), "TT", ""), "or in", ""), "in", "")))
                  right_engine_string = Left(Trim(right_engine_string), InStr(Trim(right_engine_string), "or in") - 1)
                ElseIf InStr(Trim(right_engine_string), " or ") > 0 Then
                  right_aftt = CInt(Trim(Replace(Replace(Right(Trim(right_engine_string), Len(Trim(right_engine_string)) - (InStr(Trim(right_engine_string), " or ") + 4)), "TT", ""), " or ", "")))
                  right_engine_string = Left(Trim(right_engine_string), InStr(Trim(right_engine_string), " or ") - 1)
                End If

              End If

            End If
          End If
        End If
        '---------------THIS SECTION IS ONLY USED FOR THE LEFT AND RIGHT ENGINE -----------------------------------------------------



        If Trim(temp_string) <> "" And temp_spot > 0 And found_le_re = False Then
          If temp_spot > 20 Then
            temp_spot2 = InStr(temp_string, ";")
            temp_spot3 = InStr(temp_string, ":")
            If (((temp_spot3 < temp_spot2) And temp_spot3 > 0) Or (temp_spot2 = 0 And temp_spot3 > 0)) Then
              If (temp_spot3 < temp_spot) And (temp_spot3 > 0) Then
                temp_spot = temp_spot3
              End If
            Else
              If (temp_spot2 < temp_spot) And (temp_spot2 > 0) Then
                temp_spot = temp_spot2
              End If
            End If

          ElseIf InStr(Left(Trim(temp_string), temp_spot - 1), " due ") > 0 Then ' check if "due is in there for cases such as  72-Month inspection c/w 0/20/12; due 03/31/18.
            temp_spot2 = InStr(temp_string, ";")
            If (temp_spot2 < temp_spot) And (temp_spot2 > 0) Then
              temp_spot = temp_spot2
            End If
          ElseIf InStr(Left(Trim(temp_string), temp_spot - 1), " next due ") > 0 Then ' check if "due is in there for cases such as  72-Month inspection c/w 0/20/12; due 03/31/18.
            temp_spot2 = InStr(temp_string, ":")
            If (temp_spot2 < temp_spot) And (temp_spot2 > 0) Then
              temp_spot = temp_spot2
            End If
          End If
        End If


        is_due_date = False
        If temp_spot > 0 Then
          temp_string = Left(temp_string, temp_spot - 1)
          note_string = temp_string

          temp_spot = InStr(temp_string, "issued")
          If temp_spot > 0 Then
            is_due_date = False
            temp_string = Trim(Replace(temp_string, "issued", ""))
          Else

            temp_spot = InStr(temp_string, "renewal")
            If temp_spot > 0 Then
              is_due_date = False
              temp_string = Trim(Replace(temp_string, "renewal", ""))
            Else
              temp_spot = InStr(temp_string, "due")
              If temp_spot > 0 Then
                temp_string = Trim(Replace(temp_string, "due", ""))
                is_due_date = True
              Else
                temp_string = temp_string
              End If
            End If
          End If



          has_tt = False
          temp_spot = InStr(Trim(temp_string), "TT")
          If temp_spot > 0 Then
            temp_spot = temp_spot
            TT_Count = TT_Count + 1
            has_tt = True
          End If


          aftt_string = ""
          If IsDate(Right(Trim(temp_string), 8)) = True Then
            temp_string = Right(Trim(temp_string), 8)

          ElseIf IsDate(Right(Trim(temp_string), 5)) = True Then
            temp_string = Right(Trim(temp_string), 5)
          Else
            aftt_string = temp_string ' get the original string 
            aftt_orig = aftt_string
            'cut normal string out - should be down to a date
            Call replace_all_known_text(temp_string)
            aftt_temp = temp_string


            aftt_string = Replace(aftt_string, temp_string, "") ' get rid of what it finds, hopefully the date

            temp_spot = InStr(Trim(aftt_string), "TT")
            If temp_spot > 0 Then

              temp_spot2 = InStr(Trim(aftt_string), " at")
              If temp_spot2 > 0 And temp_spot > temp_spot2 Then
                aftt_string = Right(Trim(aftt_string), Len(Trim(aftt_string)) - temp_spot2)
              End If
              aftt_string = Replace(aftt_string, "or at ", "")
              aftt_string = Replace(aftt_string, "d at ", "")
              aftt_string = Replace(aftt_string, "at ", "")
              aftt_string = Replace(aftt_string, " AFTT", "TT")
              aftt_string = Replace(aftt_string, "AFTT", "TT")
              aftt_string = Replace(aftt_string, " ACTT", "TT")
              aftt_string = Replace(aftt_string, "ACTT", "TT")

              temp_spot = InStr(aftt_string, "TT")
              If temp_spot > 0 Then
                aftt_string = Left(aftt_string, temp_spot - 1)
                Call replace_all_known_text(aftt_string)

                If IsNumeric(aftt_string) = True Then
                  afft_found = afft_found + 1
                Else
                  aftt_orig = aftt_orig
                End If
              End If
            End If

          End If

          If IsNumeric(aftt_string) = False Then
            aftt_string = "0"
          End If

          If Trim(type_to_insert) = "Engine Overhaul" And InStr(temp_string, "LE/RE") > 0 And try_le_re = False Then
            ' will re-run this, and run it again with the true so that it checks for LE/RE in the above find_details_in_string function
            If find_details_in_string(temp_details, "Engine Overhaul", "Engine Overhaul", "Engine Overhaul/Hot Section inspections", "", "", "", "", "", True) Then

            End If
          End If


          If try_le_re = True And found_le_re = True Then
            '---------------THIS SECTION IS ONLY USED FOR THE LEFT AND RIGHT ENGINE -----------------------------------------------------

            ' if its just IN, or AIRCRAFT
            If Trim(UCase(note_string)) = "AIRCRAFT" Then
              note_string = ""
            ElseIf Trim(UCase(note_string)) = "IN" Then
              note_string = ""
            ElseIf Trim(UCase(note_string)) = "INSPECTION" Then
              note_string = ""
            ElseIf Len(Trim(note_string)) < 3 Then
              note_string = ""
            End If

            If IsDate(left_engine_string) = True Then
              If IsDate(Right(Trim(left_engine_string), 8)) = True Then
                left_engine_string = Right(Trim(left_engine_string), 8)
              ElseIf IsDate(Right(Trim(left_engine_string), 5)) = True Then
                left_engine_string = Right(Trim(left_engine_string), 5)
              ElseIf IsDate(Left(Trim(left_engine_string), 8)) = True Then
                left_engine_string = Left(Trim(left_engine_string), 8)
              ElseIf IsDate(Left(Trim(left_engine_string), 5)) = True Then
                left_engine_string = Left(Trim(left_engine_string), 5)
              ElseIf IsDate(Left(Trim(left_engine_string), 4)) = True Then
                left_engine_string = Left(Trim(left_engine_string), 4)
              End If
              left_engine_string = Replace(left_engine_string, ".", "")
              Call fix_date_format(left_engine_string, is_found, date_type, type_to_insert)
              If le_re_complied_or_due = "cw" Then
                Call insert_into_Aircraft_Maintenance(temp_ac_id, type_to_insert, "", left_engine_string, date_type, "LE", left_aftt, 0)
              Else
                Call insert_into_Aircraft_Maintenance(temp_ac_id, type_to_insert, "", left_engine_string, date_type, "LE", 0, left_aftt)
              End If
            ElseIf InStr(Trim(left_engine_string), "TT") > 0 Then
              left_engine_string = Replace(left_engine_string, "TT", "")
              If Trim(le_re_complied_or_due) = "cw" Then
                Call insert_into_Aircraft_Maintenance(temp_ac_id, type_to_insert, "", "", date_type, "LE", CInt(left_engine_string), 0)
              Else
                Call insert_into_Aircraft_Maintenance(temp_ac_id, type_to_insert, "", "", date_type, "LE", 0, CInt(left_engine_string))
              End If
            End If

            If IsDate(right_engine_string) = True Then
              If IsDate(Right(Trim(right_engine_string), 8)) = True Then
                right_engine_string = Right(Trim(right_engine_string), 8)
              ElseIf IsDate(Right(Trim(right_engine_string), 5)) = True Then
                right_engine_string = Right(Trim(right_engine_string), 5)
              ElseIf IsDate(Left(Trim(right_engine_string), 8)) = True Then
                right_engine_string = Left(Trim(right_engine_string), 8)
              ElseIf IsDate(Left(Trim(right_engine_string), 5)) = True Then
                right_engine_string = Left(Trim(right_engine_string), 5)
              ElseIf IsDate(Left(Trim(right_engine_string), 4)) = True Then
                right_engine_string = Left(Trim(right_engine_string), 4)
              End If
              right_engine_string = Replace(right_engine_string, ".", "")
              Call fix_date_format(right_engine_string, is_found, date_type, type_to_insert)
              If le_re_complied_or_due = "cw" Then
                Call insert_into_Aircraft_Maintenance(temp_ac_id, type_to_insert, "", right_engine_string, date_type, "RE", right_aftt, 0)
              Else
                Call insert_into_Aircraft_Maintenance(temp_ac_id, type_to_insert, "", right_engine_string, date_type, "RE", 0, right_aftt)
              End If
            ElseIf InStr(Trim(right_engine_string), "TT") > 0 Then
              right_engine_string = Replace(right_engine_string, "TT", "")
              If Trim(le_re_complied_or_due) = "cw" Then
                Call insert_into_Aircraft_Maintenance(temp_ac_id, type_to_insert, "", "", date_type, "RE", CInt(right_engine_string), 0)
              Else
                Call insert_into_Aircraft_Maintenance(temp_ac_id, type_to_insert, "", "", date_type, "RE", 0, CInt(right_engine_string))
              End If
            End If


            Response.Write("<br>LE--->" & left_engine_string & ", " & left_aftt & " AFTT")
            Response.Write("<br>RE--->" & right_engine_string & ", " & right_aftt & " AFTT")




            '---------------THIS SECTION IS ONLY USED FOR THE LEFT AND RIGHT ENGINE -----------------------------------------------------
          Else ' do the original section of information---------
            If IsDate(temp_string) = False Then
              If IsDate(Right(Trim(temp_string), 8)) = True Then
                temp_string = Right(Trim(temp_string), 8)
              ElseIf IsDate(Right(Trim(temp_string), 5)) = True Then
                temp_string = Right(Trim(temp_string), 5)
              ElseIf IsDate(Left(Trim(temp_string), 8)) = True Then
                temp_string = Left(Trim(temp_string), 8)
              ElseIf IsDate(Left(Trim(temp_string), 5)) = True Then
                temp_string = Left(Trim(temp_string), 5)
              End If
            End If


            ' if the aftt is in the notes, then remove it 
            If Trim(note_string) <> "" Then
              If Trim(aftt_string) <> "" And Trim(aftt_string) <> "0" Then
                If InStr(Trim(note_string), Trim(aftt_string)) > 0 Then
                  note_string = Replace(note_string, aftt_string, "")
                  note_string = Replace(note_string, "or at ", "")
                  note_string = Replace(note_string, "d at ", "")
                  note_string = Replace(note_string, "at ", "")
                  note_string = Replace(note_string, " AFTT", "")
                  note_string = Replace(note_string, "AFTT", "")
                  note_string = Replace(note_string, " ACTT", "")
                  note_string = Replace(note_string, "ACTT", "")
                  note_string = Replace(note_string, "TT", "")
                End If
              End If
            End If


            note_string = Trim(Replace(note_string, temp_string, ""))
            note_string = Trim(Replace(note_string, "c/w", ""))
            note_string = Trim(Replace(note_string, "due", ""))
            note_string = Trim(Replace(note_string, "renewal", ""))
            note_string = Trim(Replace(note_string, "issued", ""))

            ' if its just IN, or AIRCRAFT
            If Trim(UCase(note_string)) = "AIRCRAFT" Then
              note_string = ""
            ElseIf Trim(UCase(note_string)) = "IN" Then
              note_string = ""
            ElseIf Trim(UCase(note_string)) = "INSPECTION" Then
              note_string = ""
            ElseIf Len(Trim(note_string)) < 3 Then
              note_string = ""
            End If


            Call fix_date_format(temp_string, is_found, date_type, type_to_insert)

            If is_found = True Then

              If Trim(type_to_insert) = "36-Month Inspection" Or Trim(type_to_insert) = "96-Month Inspection" Then
                found_36_96 = found_36_96 & " " & temp_string
              End If

              If Trim(found_36_96) <> "" And Trim(type_to_insert) = "6-Month Inspection" And InStr(Trim(found_36_96), Trim(temp_string)) > 0 Then
                ' if there is a 96 or 36 month inspection and there is a "6-month" that could be one of those
                'check to see if date is the same, if so, dont do 6-month
              Else
                If is_due_date = True Then
                  Call insert_into_Aircraft_Maintenance(temp_ac_id, type_to_insert, "", temp_string, date_type, note_string, 0, aftt_string)
                Else
                  Call insert_into_Aircraft_Maintenance(temp_ac_id, type_to_insert, temp_string, "", date_type, note_string, aftt_string, 0)
                End If
              End If
            End If
          End If
        Else
          If Trim(type_to_insert) = "Phase 5 Inspection" Then
            type_to_insert = type_to_insert
          End If
          ' Response.Write("<br>NO DATE FOUND (" & type_to_insert & ") TT:" & has_tt & " -> " & original_temp_details)
          Response.Write("<br>NO DATE FOUND (" & type_to_insert & ") -> " & temp_string & "-->(ORIG)-->" & original_temp_details)
          error_count = error_count + 1
        End If
      End If

    Catch ex As Exception
      Response.Write("<Br>ERROR: " & ex.ToString)
    End Try
  End Sub
  Public Sub replace_all_known_text(ByRef temp_string As String)

    temp_string = Trim(Replace(temp_string, "are currently in progresss as reported", ""))
    temp_string = Trim(Replace(temp_string, "are currently in progress as reported", ""))
    temp_string = Trim(Replace(temp_string, "& gear detailed inspection", ""))
    temp_string = Trim(Replace(temp_string, "& landing gear overhaul scheduled for", ""))
    temp_string = Trim(Replace(temp_string, "& landing gear ndt", ""))
    temp_string = Trim(Replace(temp_string, "landing gear overhaul scheduled for", ""))
    temp_string = Trim(Replace(temp_string, "scheduled for completion", ""))
    temp_string = Trim(Replace(temp_string, "c/w by gulfstream as reported", ""))
    temp_string = Trim(Replace(temp_string, "is currently in progress as reported", ""))
    temp_string = Trim(Replace(temp_string, "scheduled to be", ""))
    temp_string = Trim(Replace(temp_string, "& ifis Mod", ""))
    temp_string = Trim(Replace(temp_string, "scheduled for ", ""))


    temp_string = Trim(Replace(temp_string, "items", ""))
    temp_string = Trim(Replace(temp_string, "as reporrted", ""))
    temp_string = Trim(Replace(temp_string, "as reported", ""))
    temp_string = Trim(Replace(temp_string, "due", ""))
    temp_string = Trim(Replace(temp_string, "renewal", ""))
    temp_string = Trim(Replace(temp_string, "issued", ""))
    temp_string = Trim(Replace(temp_string, "c/w", ""))
    If InStr(Trim(temp_string), " in ") > 4 Then
      temp_string = Left(temp_string, InStr(Trim(temp_string), " in ")) ' if its in .. wherever the date should be before
    Else
      temp_string = Trim(Replace(temp_string, " in ", ""))
    End If


    temp_string = Trim(Replace(temp_string, " are ", ""))
    temp_string = Trim(Replace(temp_string, "as ", ""))
    temp_string = Trim(Replace(temp_string, "progress", ""))

    If Left(temp_string, 2) = "n " Then
      temp_string = Trim(Replace(temp_string, "n ", ""))
    End If

    If Left(temp_string, 3) = "in " Then
      temp_string = Trim(Replace(temp_string, "in ", ""))
    End If

    If Left(temp_string, 3) = "is " Then
      temp_string = Trim(Replace(temp_string, "is ", ""))
    End If

    If Left(temp_string, 4) = "are " Then
      temp_string = Trim(Replace(temp_string, "are ", ""))
    End If

    temp_spot = InStr(Trim(temp_string), " at ")
    If temp_spot > 0 Then
      temp_string = Left(Trim(temp_string), temp_spot - 1)
    End If

    temp_spot = InStr(Trim(temp_string), " by ")
    If temp_spot > 0 Then
      temp_string = Left(Trim(temp_string), temp_spot - 1)
    End If

    temp_spot = InStr(Trim(temp_string), ";")
    If temp_spot > 0 Then
      temp_string = Left(Trim(temp_string), temp_spot - 1)
    End If

  End Sub
  Public Sub fix_date_format(ByRef temp_string As String, ByRef is_found As Boolean, ByRef date_type As String, ByVal type_to_insert As String)

    Dim temp_month As String = ""
    Dim this_spot As Integer = 0
    Dim temp_orig As String = ""

    Try

      temp_orig = Trim(temp_string)

      before_2000 = Right(Trim(temp_string), 2)
      If IsNumeric(before_2000) Then
        If CInt(before_2000) > 50 Then
          If Len(Trim(temp_string)) = 5 Then
            before_2000 = Left(Trim(temp_string), Len(temp_string) - 2) & "01/19" & before_2000
            temp_string = before_2000
          ElseIf Len(Trim(temp_string)) = 4 Then
            before_2000 = Left(Trim(temp_string), 2) & "01/19" & before_2000
            temp_string = before_2000
          End If
        ElseIf CInt(before_2000) = 0 Then
          If Len(Trim(temp_string)) = 5 Then
            before_2000 = Left(Trim(temp_string), Len(temp_string) - 2) & "01/20" & before_2000
            temp_string = before_2000
          End If
        End If
      End If


      If IsDate(temp_string) Then
        If Len(temp_string) >= 8 Then
          'Response.Write("<br>" & temp_string) 
          temp_string = CDate(temp_string)
          date_type = "D"
          is_found = True

        ElseIf Len(temp_string) = 5 Then
          'Response.Write("<br>YEAR/MONTH FOUND->" & temp_string) 
          is_found = True
          date_type = "M"

          temp_date_sep = Month(temp_string) & "/01/20" & Right(temp_string, 2)
          temp_string = temp_date_sep
        ElseIf Len(temp_string) = 7 Then
          ' if its a date, and its 7 chars - 08/2010 
          is_found = True
          date_type = "M"

          temp_date_sep = Month(temp_string) & "/01/20" & Right(temp_string, 2)
          temp_string = temp_date_sep
        ElseIf Len(temp_string) = 4 And InStr(Trim(temp_string), "/") > 0 Then
          ' if its a date, and its 7 chars - 08/2010 
          is_found = True
          date_type = "M"
          temp_date_sep = Month(temp_string) & "/01/20" & Right(temp_string, 2)
          temp_string = temp_date_sep
        End If

      Else

        If (Len(temp_string) = 4 And IsNumeric(temp_string)) Or (Len(temp_string) = 2 And IsNumeric(temp_string)) Then
          ' Response.Write("<br>YEAR FOUND->" & temp_string)
          is_found = True
          date_type = "Y"

          If Len(temp_string) = 4 Then
            temp_string = "01/01/" & Trim(temp_string)
          ElseIf Len(temp_string) = 2 Then
            temp_string = "01/01/20" & Trim(temp_string)
          End If

        Else

          temp_month = ""
          If InStr(Trim(temp_string), "january") > 0 Or InStr(Trim(temp_string), "jan") > 0 Then
            temp_month = "01"
          ElseIf InStr(Trim(temp_string), "february") > 0 Or InStr(Trim(temp_string), "feb") > 0 Then
            temp_month = "02"
          ElseIf InStr(Trim(temp_string), "march") > 0 Or InStr(Trim(temp_string), "mar") > 0 Then
            temp_month = "03"
          ElseIf InStr(Trim(temp_string), "april") > 0 Or InStr(Trim(temp_string), "apr") > 0 Then
            temp_month = "04"
          ElseIf InStr(Trim(temp_string), "may") > 0 Or InStr(Trim(temp_string), "may") > 0 Then
            temp_month = "05"
          ElseIf InStr(Trim(temp_string), "june") > 0 Or InStr(Trim(temp_string), "jun") > 0 Then
            temp_month = "06"
          ElseIf InStr(Trim(temp_string), "july") > 0 Or InStr(Trim(temp_string), "jul") > 0 Then
            temp_month = "07"
          ElseIf InStr(Trim(temp_string), "august") > 0 Or InStr(Trim(temp_string), "aug") > 0 Then
            temp_month = "08"
          ElseIf InStr(Trim(temp_string), "september") > 0 Or InStr(Trim(temp_string), "sep") > 0 Then
            temp_month = "09"
          ElseIf InStr(Trim(temp_string), "october") > 0 Or InStr(Trim(temp_string), "oct") > 0 Then
            temp_month = "10"
          ElseIf InStr(Trim(temp_string), "november") > 0 Or InStr(Trim(temp_string), "nov") > 0 Then
            temp_month = "11"
          ElseIf InStr(Trim(temp_string), "december") > 0 Or InStr(Trim(temp_string), "dec") > 0 Then
            temp_month = "12"
          End If

          'if we find a month, then try to see if it is something like August, 2005
          If Trim(temp_month) <> "" Then
            this_spot = InStr(temp_string, ", ")
            If this_spot > 0 Then
              temp_string = Right(Trim(temp_string), Len(Trim(temp_string)) - this_spot)
              If IsNumeric(Trim(temp_string)) Then
                temp_string = temp_month & "/01/" & Trim(temp_string)
                is_found = True
              Else
                is_found = False
              End If
            Else
              is_found = False
            End If
          Else
            is_found = False
          End If

          If is_found = False Then
            If Trim(type_to_insert) = "Phase 5 Inspection" Then
              type_to_insert = type_to_insert
            End If
            'Response.Write("<br>NO DATE FOUND (" & type_to_insert & ") TT:" & has_tt & " - -> " & original_temp_details)
            Response.Write("<br>NO DATE FOUND (" & type_to_insert & ") -> " & temp_string & "-->(ORIG)-->" & original_temp_details)
            error_count = error_count + 1
          End If
        End If

      End If

    Catch ex As Exception

    End Try
  End Sub

  Public Function insert_into_Aircraft_Maintenance(ByVal ac_id As Long, ByVal maint_name As String, ByVal compiled_date As String, ByVal due_date As String, ByVal data_type As String, ByVal note_string As String, ByVal maint_complied_tt As Integer, ByVal maint_due_tt As Integer) As Boolean
    insert_into_Aircraft_Maintenance = False
    Dim i As Integer = 0

    Try


      If Trim(maint_name) = "Phase 1 - 4 Inspection" Then
        Call check_and_insert_query(ac_id, "Phase 1 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Phase 2 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Phase 3 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Phase 4 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
      ElseIf Trim(maint_name) = "Phase 1 - 5 Inspection" Then
        Call check_and_insert_query(ac_id, "Phase 1 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Phase 2 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Phase 3 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Phase 4 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Phase 5 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
      ElseIf Trim(maint_name) = "Phase 3 - 5 inspections" Then
        Call check_and_insert_query(ac_id, "Phase 3 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Phase 4 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Phase 5 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)

      ElseIf Trim(maint_name) = "Phase 1 & 2 inspections" Then
        Call check_and_insert_query(ac_id, "Phase 1 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Phase 2 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)

      ElseIf Trim(maint_name) = "Phase 3 & 4 inspections" Then
        Call check_and_insert_query(ac_id, "Phase 3 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Phase 4 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)


      ElseIf Trim(maint_name) = "12/24-Month items" Then
        Call check_and_insert_query(ac_id, "12-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "24-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)

      ElseIf Trim(maint_name) = "24 & 48-Month inspection" Then
        Call check_and_insert_query(ac_id, "24-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "48-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)


      ElseIf Trim(maint_name) = "12/24/36/48/72/144-Month inspections" Then
        Call check_and_insert_query(ac_id, "12-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "24-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "36-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "48-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "72-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "144-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)

      ElseIf Trim(maint_name) = "12/24/36/72/144-Month inspections" Then
        Call check_and_insert_query(ac_id, "12-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "24-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "36-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "72-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "144-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)

      ElseIf Trim(maint_name) = "12/48/144-Month inspections" Then
        Call check_and_insert_query(ac_id, "12-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "48-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "144-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)

      ElseIf Trim(maint_name) = "72/144-Month inspections" Then
        Call check_and_insert_query(ac_id, "72-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "144-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)



      ElseIf Trim(maint_name) = "12/24/48/180-Month inspections" Then
        Call check_and_insert_query(ac_id, "12-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "24-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "48-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "180-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)

      ElseIf Trim(maint_name) = "180/240-Month Checks" Then
        Call check_and_insert_query(ac_id, "180-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "240-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)


      ElseIf Trim(maint_name) = "Documents 8 & 10" Then
        Call check_and_insert_query(ac_id, "Document 8 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Document 10 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)

      ElseIf Trim(maint_name) = "Document 1, 2, 4, 10, 11, 12, 13, 17, 36 " Then
        Call check_and_insert_query(ac_id, "Document 1 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Document 2 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Document 4 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Document 10 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Document 11 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Document 12 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Document 13 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Document 17 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "Document 36 Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)

      ElseIf Trim(maint_name) = "12, 24, 48-Month inspection" Then
        Call check_and_insert_query(ac_id, "12-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "24-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "48-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)

      ElseIf Trim(maint_name) = "12/36/72-Month Inspection" Then
        Call check_and_insert_query(ac_id, "12-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "36-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "72-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)


      ElseIf Trim(maint_name) = "Document 1, 2, 4, 10, 11, 12, 13, 17, 36 " Then

        For i = 1 To 12
          Call check_and_insert_query(ac_id, "Document " & i & " Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Next

      ElseIf Trim(maint_name) = "12/24/48 & 96-Month inspection" Then
        Call check_and_insert_query(ac_id, "12-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "24-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "48-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "96-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
      ElseIf Trim(maint_name) = "12/24/36/48/96-Month Inspection" Then
        Call check_and_insert_query(ac_id, "12-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "24-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "36-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "48-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
        Call check_and_insert_query(ac_id, "96-Month Inspection", compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)


      Else
        Call check_and_insert_query(ac_id, maint_name, compiled_date, due_date, data_type, note_string, maint_complied_tt, maint_due_tt)
      End If



    Catch ex As Exception
    Finally

    End Try

  End Function
  Public Function check_and_insert_query(ByVal ac_id As Long, ByVal maint_name As String, ByVal compiled_date As String, ByVal due_date As String, ByVal data_type As String, ByVal note_string As String, ByVal maint_complied_hours As Integer, ByVal maint_due_Hours As Integer) As Boolean


    Dim Insert_Query As String = ""
    Dim select_query As String = ""
    Dim atemptable As New DataTable
    Dim passed_test As Boolean = True


    Try

      If Trim(due_date) = "1/01/20/3" Or Trim(compiled_date) = "1/01/20/3" Then
      ElseIf Trim(due_date) = "2/01/20/4" Or Trim(compiled_date) = "2/01/20/4" Then
      ElseIf Trim(due_date) = "1/01/20/1" Or Trim(compiled_date) = "1/01/20/1" Then
      ElseIf Trim(due_date) = "4/01/20/1" Or Trim(compiled_date) = "4/01/20/1" Then
      ElseIf Trim(due_date) = "2/01/20/9" Or Trim(compiled_date) = "2/01/20/9" Then
      ElseIf Trim(due_date) = "10/01/20/2" Or Trim(compiled_date) = "10/01/20/2" Then

      Else
        select_query = " select acmaint_ac_id from Aircraft_Maintenance with (NOLOCK) where acmaint_ac_id = " & ac_id & ""

        If Trim(compiled_date) <> "" Then
          select_query &= " and acmaint_complied_date = '" & compiled_date & "' "
          If Year(compiled_date) < 1930 Then
            passed_test = False
          End If
        End If
        If Trim(due_date) <> "" Then
          select_query &= " and acmaint_due_date = '" & due_date & "' "
          If Year(due_date) < 1930 Then
            passed_test = False
          End If
        End If

        select_query &= " and acmaint_journ_id = 0 and acmaint_name = '" & maint_name & "' "


        If Trim(maint_complied_hours) > 0 Then
          select_query &= " and acmaint_complied_hrs = " & maint_complied_hours & " "
        End If

        If Trim(maint_due_Hours) > 0 Then
          select_query &= " and acmaint_due_hrs = " & maint_due_Hours & " "
        End If

        MySqlCommand_JETNET.CommandText = select_query
        MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()

        Try
          atemptable.Load(MyAircraftReader_JETNET)
        Catch constrExc As System.Data.ConstraintException
          Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
          ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
        End Try

        ac_found_count = ac_found_count + 1

        ' will now fail test for bad dates
        If atemptable.Rows.Count = 0 And passed_test = True Then

          Insert_Query = " INSERT INTO Aircraft_Maintenance"
          Insert_Query &= " (acmaint_ac_id"
          Insert_Query &= " ,acmaint_journ_id"
          Insert_Query &= " ,acmaint_name"
          Insert_Query &= " ,acmaint_complied_date"
          Insert_Query &= " ,acmaint_complied_hrs"
          Insert_Query &= " ,acmaint_due_date"
          Insert_Query &= " ,acmaint_due_hrs"
          Insert_Query &= " ,acmaint_notes"
          Insert_Query &= " ,acmaint_date_type)"
          Insert_Query &= " VALUES( "
          Insert_Query &= " " & ac_id & ""
          Insert_Query &= " ,0"
          Insert_Query &= " ,'" & maint_name & "'"

          If Trim(compiled_date) <> "" Then
            Insert_Query &= " ,'" & compiled_date & "'"
          Else
            Insert_Query &= " ,NULL"
          End If

          Insert_Query &= " ," & maint_complied_hours ' <acmaint_complied_hrs, int,>

          If Trim(due_date) <> "" Then
            Insert_Query &= " ,'" & due_date & "'"
          Else
            Insert_Query &= " ,NULL"
          End If


          Insert_Query &= " ," & maint_due_Hours  ' <acmaint_due_hrs, int,>"

          note_string = Replace(note_string, "'", "")
          If Trim(before_text) <> "" Then
            Insert_Query &= " ,'" & note_string & " (" & Replace(before_text, "'", "") & ")'" ' 
          Else
            Insert_Query &= " ,'" & note_string & "'" ' <acmaint_notes, varchar(300),>
          End If


          Insert_Query &= "  ,'" & data_type & "')" ' <acmaint_date_type, char(1)

          ' Response.Write("<Br>" & Insert_Query)
          Insert_Query = Insert_Query
          MySqlCommand_JETNET.CommandText = Insert_Query
          MySqlCommand_JETNET.ExecuteNonQuery()

          ac_insert_count = ac_insert_count + 1
          check_and_insert_query = True
          inserted_any_info = True
        Else
          check_and_insert_query = check_and_insert_query
        End If

      End If

    Catch ex As Exception
    Finally
      MyAircraftReader_JETNET.Close()
      MySqlCommand_JETNET.Dispose()
      atemptable = Nothing
    End Try
  End Function
  Public Function get_yacht_news_super_yachts() As String
    get_yacht_news_super_yachts = ""

    Dim we_have_count As Integer = 0
    Dim we_dont_have_count As Integer = 0
    Dim found_news As Integer = 0
    Dim didnt_find_yacht As String = ""
    Dim results As String = ""
    Dim query As String = ""
    Dim temp_link As String = ""
    Dim results_table As New DataTable
    Dim continue_search As Boolean = False
    Dim company_count As Integer = 0
    Dim company_string As String = ""
    Dim this_company_yachts As Integer = 0
    Dim total_found_temp As Integer = 0
    Dim i As Integer = 0
    Dim found_yacht As Integer = 0
    Dim connected_yachts As Integer = 0
    Dim found_companies As Integer = 0



    Try


      ' MySqlConn_JETNET.ConnectionString = "Data Source=www.jetnetsql2.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
      '' MySqlConn_JETNET2.ConnectionString = "Data Source=www.jetnetsql2.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
      '   MySqlConn_JETNET.ConnectionString = "Data Source=128.1.21.40;Initial Catalog=jetnet_ra_test;Persist Security Info=True;User ID=sa;Password=moejive"
      ' MySqlConn_JETNET.ConnectionString = "Data Source=128.1.21.200;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=sa;Password=moejive"
      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120

      Call insert_into_eventlog("Yacht News Started", "Research Assistant")

      results_table = GET_COMPANY_LIST("2")

      If Not IsNothing(results_table) Then
        If results_table.Rows.Count > 0 Then
          For Each r As DataRow In results_table.Rows
            If Not IsDBNull(r.Item("ytmap_src_comp_id")) Then
              If Not IsDBNull(r.Item("comp_id")) Then

                continue_search = Scrape_This_Page_Super_Yachts(r.Item("ytmap_web_address"), found_news, found_companies, found_yacht, didnt_find_yacht, r.Item("comp_id"), r.Item("comp_name"), results, i)

                company_count = company_count + 1
              End If
            End If
          Next
        End If
      End If



      'results = "<br><table cellspacing='0' cellpadding='0' border='0' valign='top'>"
      'results &= "<tr><Td align='left'><font color='black'>Super Yachts</font></td></tr>"
      'results &= "<tr><Td align='left'><font color='black'>Total Companies Searched: " & CStr(company_count) & "</font>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"
      'results &= "<tr><Td align='left'><font color='black'>News Entered Into Yacht-Spot: " & CStr(found_news) & "</font>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"
      'results &= "<tr><Td align='left'><font color='black'>Total Companies With New News: " & CStr(found_companies) & "</font>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"
      'results &= "<tr><Td align='left'><font color='black'>Yachts Connected to Articles: " & CStr(found_yacht) & "</font>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"


      results = Chr(13) & Chr(10) & "Super Yachts - www.superyachts.com:" & Chr(13) & Chr(10)
      results &= "Companies Searched: " & CStr(company_count) & "" & Chr(13) & Chr(10)
      results &= "News Entered Into Yacht-Spot: " & CStr(found_news) & "" & Chr(13) & Chr(10)
      results &= "Companies With New News: " & CStr(found_companies) & "" & Chr(13) & Chr(10)
      results &= "Yachts Connected to Articles: " & CStr(found_yacht) & "" & Chr(13) & Chr(10)

      TOTAL_NEWS = TOTAL_NEWS + found_news
      TOTAL_COMPANIES = TOTAL_COMPANIES + company_count
      TOTAL_COMPANIES_CONNECTED = TOTAL_COMPANIES_CONNECTED + found_companies
      TOTAL_YACHTS_CONNECTED = TOTAL_YACHTS_CONNECTED + found_yacht

      ' results &= "</table>"
      ' results &= "</td>"

      ' results &= "</td></tr></table>"

      Me.text_label.Text = results
      get_yacht_news_super_yachts = results

    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()
    End Try
  End Function

  Public Function get_yacht_news_super_yacht_times_NEW() As String
    get_yacht_news_super_yacht_times_NEW = ""

    Dim we_have_count As Integer = 0
    Dim we_dont_have_count As Integer = 0
    Dim found_news As Integer = 0
    Dim didnt_find_yacht As String = ""
    Dim results As String = ""
    Dim query As String = ""
    Dim temp_link As String = ""
    Dim results_table As New DataTable
    Dim continue_search As Boolean = False
    Dim company_count As Integer = 0
    Dim company_string As String = ""
    Dim this_company_yachts As Integer = 0
    Dim total_found_temp As Integer = 0
    Dim i As Integer = 0
    Dim connected_yachts As Integer = 0
    Dim found_yacht As Integer = 0
    Dim found_companies As Integer = 0




    Try


      ' MySqlConn_JETNET.ConnectionString = "Data Source=www.jetnetsql2.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
      '' MySqlConn_JETNET2.ConnectionString = "Data Source=www.jetnetsql2.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
      '  MySqlConn_JETNET.ConnectionString = "Data Source=128.1.21.40;Initial Catalog=jetnet_ra_test;Persist Security Info=True;User ID=sa;Password=moejive"
      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120


      results_table = GET_COMPANY_LIST("1")

      If Not IsNothing(results_table) Then
        If results_table.Rows.Count > 0 Then
          For Each r As DataRow In results_table.Rows

            If Not IsDBNull(r.Item("ytmap_src_comp_id")) Then
              If Not IsDBNull(r.Item("comp_id")) Then
 
                continue_search = Scrape_This_Page_Yacht_Times_NEW2("", found_news, found_companies, found_yacht, didnt_find_yacht, 0, "", results, i)

                ' continue_search = Scrape_This_Page_Yacht_Times_NEW(r.Item("ytmap_web_address"), found_news, found_companies, found_yacht, didnt_find_yacht, r.Item("comp_id"), r.Item("comp_name"), results, i)

                company_count = company_count + 1
              End If
            End If

          Next
        End If
      End If


      results = Chr(13) & Chr(10) & "Super Yacht Times - www.superyachttimes.net:" & Chr(13) & Chr(10)
      results &= "Companies Searched: " & CStr(company_count) & "" & Chr(13) & Chr(10)
      results &= "News Entered Into Yacht-Spot: " & CStr(found_news) & "" & Chr(13) & Chr(10)
      results &= "Companies With New News: " & CStr(found_companies) & "" & Chr(13) & Chr(10)
      results &= "Yachts Connected to Articles: " & CStr(found_yacht) & "" & Chr(13) & Chr(10)

      TOTAL_NEWS = TOTAL_NEWS + found_news
      TOTAL_COMPANIES = TOTAL_COMPANIES + company_count
      TOTAL_COMPANIES_CONNECTED = TOTAL_COMPANIES_CONNECTED + found_companies
      TOTAL_YACHTS_CONNECTED = TOTAL_YACHTS_CONNECTED + found_yacht



      Me.text_label.Text = results
      get_yacht_news_super_yacht_times_NEW = results


    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()

    End Try
  End Function
  Public Function get_yacht_news_super_yacht_times() As String
    get_yacht_news_super_yacht_times = ""

    Dim we_have_count As Integer = 0
    Dim we_dont_have_count As Integer = 0
    Dim found_news As Integer = 0
    Dim didnt_find_yacht As String = ""
    Dim results As String = ""
    Dim query As String = ""
    Dim temp_link As String = ""
    Dim results_table As New DataTable
    Dim continue_search As Boolean = False
    Dim company_count As Integer = 0
    Dim company_string As String = ""
    Dim this_company_yachts As Integer = 0
    Dim total_found_temp As Integer = 0
    Dim i As Integer = 0
    Dim connected_yachts As Integer = 0
    Dim found_yacht As Integer = 0
    Dim found_companies As Integer = 0




    Try


      ' MySqlConn_JETNET.ConnectionString = "Data Source=www.jetnetsql2.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
      '' MySqlConn_JETNET2.ConnectionString = "Data Source=www.jetnetsql2.com;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=crmexport;Password=d4gpt9f8"
      '  MySqlConn_JETNET.ConnectionString = "Data Source=128.1.21.40;Initial Catalog=jetnet_ra_test;Persist Security Info=True;User ID=sa;Password=moejive"
      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection

      MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120


      results_table = GET_COMPANY_LIST("1")

      If Not IsNothing(results_table) Then
        If results_table.Rows.Count > 0 Then
          For Each r As DataRow In results_table.Rows
            If Not IsDBNull(r.Item("ytmap_src_comp_id")) Then
              If Not IsDBNull(r.Item("comp_id")) Then


                continue_search = Scrape_This_Page_Yacht_Times("http://www.superyachttimes.com/shipyards/company/id/" & r.Item("ytmap_src_comp_id") & "/", found_news, found_companies, found_yacht, didnt_find_yacht, r.Item("comp_id"), r.Item("comp_name"), results, i)

                company_count = company_count + 1
              End If
            End If
          Next
        End If
      End If



      'results = "<br><table cellspacing='0' cellpadding='0' border='0' valign='top'>"
      'results &= "<tr><Td align='left'><font color='black'>Super Yacht Times</font></td></tr>"
      'results &= "<tr><Td align='left'><font color='black'>Total Companies Searched: " & CStr(company_count) & "</font>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"
      'results &= "<tr><Td align='left'><font color='black'>News Entered Into Yacht-Spot: " & CStr(found_news) & "</font>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"
      'results &= "<tr><Td align='left'><font color='black'>Total Companies With New News: " & CStr(found_companies) & "</font>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"
      'results &= "<tr><Td align='left'><font color='black'>Yachts Connected to Articles: " & CStr(found_yacht) & "</font>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"


      'results &= "</table>"


      'results &= "</td>"
      'results &= "</td></tr></table>"



      results = Chr(13) & Chr(10) & "Super Yacht Times - www.superyachttimes.net:" & Chr(13) & Chr(10)
      results &= "Companies Searched: " & CStr(company_count) & "" & Chr(13) & Chr(10)
      results &= "News Entered Into Yacht-Spot: " & CStr(found_news) & "" & Chr(13) & Chr(10)
      results &= "Companies With New News: " & CStr(found_companies) & "" & Chr(13) & Chr(10)
      results &= "Yachts Connected to Articles: " & CStr(found_yacht) & "" & Chr(13) & Chr(10)

      TOTAL_NEWS = TOTAL_NEWS + found_news
      TOTAL_COMPANIES = TOTAL_COMPANIES + company_count
      TOTAL_COMPANIES_CONNECTED = TOTAL_COMPANIES_CONNECTED + found_companies
      TOTAL_YACHTS_CONNECTED = TOTAL_YACHTS_CONNECTED + found_yacht



      Me.text_label.Text = results
      get_yacht_news_super_yacht_times = results


    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET.Dispose()

      MySqlCommand_JETNET.Dispose()

    End Try
  End Function
  Public Sub map_company_to_yacht(ByRef results_table As DataTable, ByRef connected_yachts As Integer)
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long
    Dim news_id(500) As String
    Dim news_title(500) As String
    Dim yacht_count As Long = 0
    Dim z As Integer = 0
    Dim x As Integer = 0
    Dim news_count As Long = 0
    Dim results_table2 As New DataTable
    Dim Update_Query As String = ""

    '----------------- BEFORE YOU SPLIT ON THE DIFFERENT NEWS ARTICLES FOR THIS COMPANY, GET ITS YACHTS-----------
    If Not IsNothing(results_table) Then
      If results_table.Rows.Count > 0 Then
        For Each r As DataRow In results_table.Rows
          If Not IsDBNull(r.Item("ytmap_src_comp_id")) Then    ' get company 1 at a time
            If Not IsDBNull(r.Item("comp_id")) Then


              If r.Item("ytmap_src_comp_id") = 43 Then
                news_count = news_count
              End If

              '--------------GET NEWS ARTICLES RELATED TO COMANY-------------------------
              news_count = 0
              results_table2 = GET_NEWS_FOR_COMPANY(r.Item("comp_id")) ' get yachts for that company
              If Not IsNothing(results_table2) Then
                If results_table2.Rows.Count > 0 Then
                  For Each k As DataRow In results_table2.Rows
                    If Not IsDBNull(k.Item("ytnews_id")) Then   'fill in arrays
                      news_id(news_count) = k.Item("ytnews_id")
                      news_title(news_count) = k.Item("ytnews_title")
                      news_count = news_count + 1
                    End If
                  Next
                End If
              End If
              '-------------------------------------------



              ' if there is news, try to connect yachts
              If news_count > 0 Then
                '--------------GET YACHTS RELATED TO COMANY-------------------------
                yacht_count = 0
                results_table2 = GET_YACHTS_FOR_COMPANY(r.Item("comp_id"), "1") ' get yachts for that company
                If Not IsNothing(results_table2) Then
                  If results_table2.Rows.Count > 0 Then
                    For Each k As DataRow In results_table2.Rows
                      If Not IsDBNull(k.Item("yt_yacht_name")) Then   'fill in arrays
                        yacht_name_array(yacht_count) = k.Item("yt_yacht_name")
                        yacht_id(yacht_count) = k.Item("yt_id")
                        yacht_count = yacht_count + 1
                      End If
                    Next
                  End If
                End If
                '------------------------------------------- 


                ' if it has yachts and news----try to connect---------
                If yacht_count > 0 Then
                  For x = 0 To news_count - 1
                    For z = 0 To yacht_count - 1
                      If Len(Trim(yacht_name_array(z))) > 3 Then
                        If InStr(news_title(x), Trim(yacht_name_array(z))) > 0 Then
                          Update_Query = ""
                          Update_Query = Update_Query & " Update Yacht_News set "
                          Update_Query = Update_Query & " ytnews_yt_id = '" & yacht_id(z) & "' "
                          Update_Query = Update_Query & " where ytnews_id = " & news_id(x)
                          Response.Write(Update_Query & "<Br>")
                          MySqlCommand_JETNET.CommandText = Update_Query
                          MySqlCommand_JETNET.ExecuteNonQuery()
                          MySqlCommand_JETNET.Dispose()
                          connected_yachts = connected_yachts + 1
                        End If
                      End If
                    Next
                  Next
                End If
                ' if it has yachts and news----try to connect---------
              End If



              ' go back through, look at the title, see if the articles match the titles

              '---------------------------------------

            End If

          End If
        Next
      End If
    End If
    '----------------------------------------

  End Sub

  Public Function Get_AC_Maint_Details() As DataTable
    Dim atemptable As New DataTable

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader


    Dim Query As String = ""
    Try




      Query &= " select ac_id, adet_data_description, ac_amod_id, ac_forsale_flag "
      'Query &= " (select distinct  top 1  adet_data_description from Aircraft_Details with (NOLOCK)"
      'Query &= " where adet_ac_id = ac_id and adet_journ_id = 0 and ("
      'Query &= " (adet_data_description like '%Certificate of Airworthiness%') or (adet_data_description like '%c of a%'))"
      'Query &= " ) as cert_of_a"

      'Query &= ", (select distinct  top 1  adet_data_description from Aircraft_Details with (NOLOCK)"
      'Query &= " where adet_ac_id = ac_id and adet_journ_id = 0 and ("
      'Query &= " (adet_data_description like '%entered into service%') or (adet_data_description like '%entry into service%') or (adet_data_description like '%placed into service%') or (adet_data_description like '%service date%'))"
      'Query &= " ) as into_service "

      Query &= " from Aircraft with (NOLOCK) "
      Query &= " inner join aircraft_model with (NOLOCK) on ac_amod_id = amod_id "
      Query &= " inner join aircraft_details with (NOLOCK) on adet_ac_id = ac_id and adet_journ_id = ac_journ_id and adet_data_type = 'Maintenance' and adet_data_name='Inspection' "
      Query &= " where  ac_journ_id = 0 "
      ' Query &= " and ac_id = 31918  "
      'Query &= " and ac_id in (16313, 16471, 11607, 11631, 11503) "


      '  Query &= " and amod_make_name in ('Gulfstream','Citation', 'Hawker', 'Falcon', 'Learjet', 'Global') "   '
      ' Query &= " and amod_make_name in ('Gulfstream','Citation','Challenger','Falcon') "

      ' Query &= " And ac_amod_id in (278) "   ' , , 34, 272, 262,277,278



      'Query &= " and ("
      'Query &= "  exists(select distinct  top 1  adet_data_description from Aircraft_Details with (NOLOCK)"
      'Query &= " where adet_ac_id = ac_id and adet_journ_id = 0 and ("
      'Query &= " (adet_data_description like '%Certificate of Airworthiness%') or (adet_data_description like '%c of a%'))) "

      'Query &= " or exists(select distinct  top 1  adet_data_description from Aircraft_Details with (NOLOCK)"
      'Query &= " where adet_ac_id = ac_id and adet_journ_id = 0 and ("
      'Query &= " (adet_data_description like '%entered into service%') or (adet_data_description like '%placed into service%') or (adet_data_description like '%service date%')))"

      'Query &= " ) "


      SqlConn.ConnectionString = MySqlConn_JETNET.ConnectionString
      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        atemptable.Load(SqlReader)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

    Return atemptable

  End Function
  Public Function GET_YACHT_MMSI_LIST() As DataTable
    Dim atemptable As New DataTable

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader


    Dim Query As String = ""
    Try



      Query = "  select yt_id, yt_mmsi_mobile_nbr from Yacht with (NOLOCK) where yt_mmsi_mobile_nbr <> '' and yt_journ_id = 0 "

      SqlConn.ConnectionString = MySqlConn_JETNET.ConnectionString
      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        atemptable.Load(SqlReader)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

    Return atemptable

  End Function
  Public Function GET_COMPANY_LIST(ByVal src_id As String) As DataTable
    Dim atemptable As New DataTable

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader


    Dim Query As String = ""
    Try



      Query = "select ytnewssrc_name, comp_name, comp_city, comp_state, comp_id, ytmap_src_comp_id, ytmap_web_address "
      Query &= " from Yacht_Source_Mapping "
      Query &= " inner join Yacht_News_Source on ytmap_ytnewssrc_id = ytnewssrc_id"
      Query &= " inner join Company on ytmap_jetnet_comp_id = comp_id and comp_journ_id = 0"

      Query &= " where ytmap_ytnewssrc_id = " & src_id & " "

      If Trim(src_id) = "1" Then
        Query &= " and ytmap_web_address <> '' and ytmap_web_address is not null "
      End If

      '     Query &= "  and ytmap_jetnet_comp_id = 333059 "
      '   Query &= " and comp_name >= 'Basimakopouloi Shipyard' "
      '   Query &= " and (ytmap_src_comp_id = 68 or ytmap_src_comp_id = 122)"   '
      '  Query &= "  and comp_id not in ( "
      '  Query &= "   select distinct yr_comp_id from yacht_reference where yr_contact_type='Y6') "

      Query &= " order by comp_name, comp_city, comp_state"


      SqlConn.ConnectionString = MySqlConn_JETNET.ConnectionString
      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        atemptable.Load(SqlReader)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

    Return atemptable

  End Function
  Function Scrape_This_Page_Super_Yacht_News(ByVal link As String, ByVal page_num As Integer) As Boolean
    Scrape_This_Page_Super_Yacht_News = False

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim temp_yacht_id As String = ""
    Dim company_string As String = ""
    Dim didnt_find_temp As Integer = 0
    Dim found_temp As Integer = 0
    Dim insert_strings As String = ""
    Dim results_table As New DataTable
    Dim Insert_Query As String = ""
    Dim original_string_text As String = ""
    Dim related_articles_text As String = ""
    Dim yacht_count As Integer = 0
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long





    Try


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text



      Call Find_Text_Loop_Top3("LATEST BUSINESS NEWS", string_text)
      'Call Find_Text_Loop_Top3("LATEST LAUNCH NEWS", string_text)
      'Call Find_Text_Loop_Top3("LATEST BROKERAGE / CHARTER NEWS", string_text)
      'Call Find_Text_Loop_Top3("LATEST PRODUCTS NEWS", string_text)
      'Call Find_Text_Loop_Top3("LATEST EVENTS NEWS", string_text)
      'Call Find_Text_Loop_Top3("LATEST PRODUCTS NEWS", string_text)



    Catch ex As Exception
      Response.Write("")
    Finally

    End Try
  End Function

  Public Function Find_Text_Loop_Top3(ByVal text_to_find As String, ByVal string_text As String)
    Find_Text_Loop_Top3 = ""
    Dim article_title As String = ""
    Dim article_date As String = ""
    Dim article_text As String = ""
    Dim array_split() As String
    Dim original_string_text As String = ""
    Dim article_link As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim Str2 As System.IO.Stream
    Dim srRead2 As System.IO.StreamReader
    Dim req2 As System.Net.WebRequest
    Dim resp2 As System.Net.WebResponse
    Dim original_string_text2 As String = ""
    Dim string_text2 As String = ""

    Try


      spot_to_find = InStr(string_text, text_to_find)
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - Len(text_to_find))


        array_split = Split(string_text, "<h4>")

        '  For i = 1 To array_split.Length - 1
        For i = 1 To 18
          string_text = array_split(i)
          original_string_text = string_text


          spot_to_find = InStr(string_text, "href=")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 5)
            spot_to_find = InStr(string_text, ">")
            If spot_to_find > 0 Then
              article_link = Left(string_text, spot_to_find - 2)
              string_text = Right(string_text, Len(string_text) - spot_to_find)

              spot_to_find = InStr(article_link, "alt=")
              If spot_to_find > 0 Then
                article_link = Left(article_link, spot_to_find - 3)
              End If


              spot_to_find = InStr(string_text, "</a>")
              If spot_to_find > 0 Then
                article_title = Left(string_text, spot_to_find - 1)


                string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

                spot_to_find = InStr(string_text, "<a class")
                If spot_to_find > 0 Then
                  string_text = Left(string_text, spot_to_find - 1)
                  string_text = Replace(string_text, "<p>", "")
                  string_text = Replace(string_text, "</h4>", "")
                  string_text = Replace(string_text, "&nbsp;", " ")
                  article_text = RTrim(LTrim(string_text))
                  '   If Left(Trim(article_text), 1) = " " Then
                  article_text = Right(Trim(article_text), Len(Trim(article_text)) - 1)
                  article_text = Right(Trim(article_text), Len(Trim(article_text)) - 1)
                  'End If
                End If


              End If

              article_link = "http://www.superyachtnews.com" & article_link

              Response.Write(article_link & "-" & article_title & "<br><br>")


              Try

                req2 = System.Net.WebRequest.Create(article_link)
                resp2 = req2.GetResponse

                Str2 = resp2.GetResponseStream
                srRead2 = New System.IO.StreamReader(Str2)
                string_text2 = srRead2.ReadToEnd().ToString
                string_text2 = string_text2
                original_string_text2 = string_text2

                If InStr(Trim(string_text2), "published_time") > 0 Then
                  string_text2 = Right(Trim(string_text2), Len(Trim(string_text2)) - InStr(Trim(string_text2), "published_time") - 14)
                  If InStr(Trim(string_text2), "/>") > 0 Then
                    string_text2 = Left(Trim(string_text2), InStr(Trim(string_text2), "/>") - 3)
                    string_text2 = Replace(string_text2, "content=""", "")
                    article_date = string_text2
                    article_date = fix_date(article_date, "")
                  End If
                End If

                If Not CHECK_IF_NEWS_EXISTS(0, article_date, Replace(article_title, "'", "''"), False, "") Then
                  '---------------------------------------------------
                  Call insert_into_news(article_date, article_title, article_text, article_link, 0, 0, 0, False, "7")
                  TOTAL_NEWS = TOTAL_NEWS + 1
                  '---------------------------------------------------
                End If


              Catch ex As Exception

              End Try


            End If
          End If


        Next

      End If

    Catch ex As Exception

    End Try
  End Function

  Function FIND_MARINE_MMSI(ByVal link As String) As Boolean
    FIND_MARINE_MMSI = False

    Try

      Dim Str As System.IO.Stream
      Dim srRead As System.IO.StreamReader
      Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
      Dim resp As System.Net.WebResponse = req.GetResponse
      Dim string_text As String = ""
      Dim string_text2 As String = ""
      Dim spot_to_find As Integer = 0
      Dim spot_to_find2 As Integer = 0
      Dim i As Integer = 0
      Dim final_string As String = ""
      Dim temp_yacht_id As String = ""
      Dim company_string As String = ""
      Dim didnt_find_temp As Integer = 0
      Dim found_temp As Integer = 0
      Dim insert_strings As String = ""
      Dim results_table As New DataTable
      Dim Insert_Query As String = ""
      Dim original_string_text As String = ""
      Dim related_articles_text As String = ""
      Dim yacht_count As Integer = 0
      Dim yacht_name_array(500) As String
      Dim yacht_id(500) As Long
      Dim array_split() As String
      Dim article_title As String = ""
      Dim article_date As String = ""
      Dim article_text As String = ""
      Dim yacht_to_connect_id As Long = 0
      Dim temp_comp_count As Integer = 0
      Dim link_to_article As String = ""
      Dim extra_text As String = ""





      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text

      If InStr(Trim(original_string_text), "We can find any ship in the world, but not this page.") > 0 Then
        FIND_MARINE_MMSI = False
      Else
        FIND_MARINE_MMSI = True
      End If

    Catch ex As Exception

    End Try


  End Function

  Function Scrape_This_Page_Super_Yachts(ByVal link As String, ByRef found_news As Integer, ByRef found_companies As Integer, ByRef found_yacht As Integer, ByRef didnt_find_yacht As String, ByVal jetnet_comp_id As Long, ByVal comp_name As String, ByRef results_string As String, ByVal page_num As Integer) As Boolean
    Scrape_This_Page_Super_Yachts = False

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim temp_yacht_id As String = ""
    Dim company_string As String = ""
    Dim didnt_find_temp As Integer = 0
    Dim found_temp As Integer = 0
    Dim insert_strings As String = ""
    Dim results_table As New DataTable
    Dim Insert_Query As String = ""
    Dim original_string_text As String = ""
    Dim related_articles_text As String = ""
    Dim yacht_count As Integer = 0
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long
    Dim array_split() As String
    Dim article_title As String = ""
    Dim article_date As String = ""
    Dim article_text As String = ""
    Dim yacht_to_connect_id As Long = 0
    Dim temp_comp_count As Integer = 0
    Dim link_to_article As String = ""
    Dim extra_text As String = ""



    Try



      '-------------------------------------------------------------
      yacht_count = 0
      results_table = GET_YACHTS_FOR_COMPANY(jetnet_comp_id, "2") ' get yachts for that company
      If Not IsNothing(results_table) Then
        If results_table.Rows.Count > 0 Then
          For Each k As DataRow In results_table.Rows
            If Not IsDBNull(k.Item("yt_yacht_name")) Then   'fill in arrays
              yacht_name_array(yacht_count) = k.Item("yt_yacht_name")
              yacht_id(yacht_count) = k.Item("yt_id")
              yacht_count = yacht_count + 1
            End If
          Next
        End If
      End If
      results_table.Dispose()
      '----------------------------------------------------------



      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text

      ' if the company id wasa put in wrong
      spot_to_find = InStr(string_text, "<ul id=""newsTab"">", CompareMethod.Text)
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find - 17)

        spot_to_find = InStr(string_text, "moreDetailFooter", CompareMethod.Text)
        If spot_to_find > 0 Then
          string_text = Left(string_text, spot_to_find - 1)


          array_split = Split(string_text, "<li>")

          For i = 1 To array_split.Length - 1
            string_text = array_split(i)
            original_string_text = string_text


            spot_to_find = InStr(string_text, "alt=", CompareMethod.Text)
            If spot_to_find > 0 Then
              string_text = Right(string_text, Len(string_text) - spot_to_find - 4)

              spot_to_find = InStr(string_text, "src=", CompareMethod.Text)
              If spot_to_find > 0 Then
                article_title = Left(string_text, spot_to_find - 3)
                article_title = clean_description(article_title)
              Else
                string_text = string_text
              End If

              spot_to_find = InStr(string_text, "<a href=", CompareMethod.Text)
              If spot_to_find > 0 Then
                extra_text = Right(string_text, Len(string_text) - spot_to_find - 9)
                spot_to_find = InStr(extra_text, ">", CompareMethod.Text)
                extra_text = Left(extra_text, spot_to_find - 2)
                link_to_article = extra_text
                If Left(Trim(link_to_article), 4) = "news" Then
                  link_to_article = "http://www.superyachts.com/" & link_to_article
                End If
              End If




              spot_to_find = InStr(string_text, "<p class=""date"">", CompareMethod.Text)
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 15)
                article_date = string_text
                spot_to_find = InStr(article_date, "</p>", CompareMethod.Text)
                If spot_to_find > 0 Then
                  article_date = Left(article_date, spot_to_find - 1)
                  article_date = fix_date(article_date, "superyachts")
                End If

                spot_to_find = InStr(string_text, "<p>", CompareMethod.Text)
                If spot_to_find > 0 Then
                  string_text = Right(string_text, Len(string_text) - spot_to_find - 2)
                  spot_to_find = InStr(string_text, "</p>", CompareMethod.Text)
                  If spot_to_find > 0 Then
                    article_text = Left(string_text, spot_to_find - 1)
                    article_text = clean_description(article_text)

                    ' if it has yachts and news----try to connect-----------------------------------------
                    yacht_to_connect_id = 0
                    yacht_news_name = ""
                    If yacht_count > 0 Then
                      For z = 0 To yacht_count - 1
                        If Len(Trim(yacht_name_array(z))) > 3 Then
                          If InStr(UCase(Replace(Replace(" " & Trim(article_title) & " ", "'", "''"), ",", "")), UCase(" " & Trim(yacht_name_array(z)) & " ")) > 0 Then
                            yacht_to_connect_id = yacht_id(z)
                            yacht_news_name = Trim(yacht_name_array(z))
                            z = yacht_count + 1  ' END IT  
                          End If
                        End If
                      Next
                    End If
                    ' if it has yachts and news----try to connect-----------------------------------------

                    '  If DateDiff(DateInterval.Day, CDate(article_date), Date.Now) < 7 Then
                    If Not CHECK_IF_NEWS_EXISTS(jetnet_comp_id, article_date, Replace(article_title, "'", "''"), False, "2") Then ' changed 
                      '---------------------------------------------------
                      Call insert_into_news(article_date, article_title, article_text, link_to_article, yacht_to_connect_id, jetnet_comp_id, 0, False, "2")

                      found_news = found_news + 1

                      If yacht_to_connect_id > 0 Then
                        found_yacht = found_yacht + 1
                      End If

                      temp_comp_count = temp_comp_count + 1
                      '---------------------------------------------------
                    End If
                    'End If

                  End If
                End If

              Else
                string_text = string_text
              End If


            Else
              string_text = string_text
            End If


          Next

        End If
      End If


      If temp_comp_count > 0 Then
        found_companies = found_companies + 1
      End If

    Catch ex As Exception
      Response.Write("")
    Finally

    End Try
  End Function
  Public Function Scrape_This_Page_Yacht_Times(ByVal link As String, ByRef found_news As Integer, ByRef found_companies As Integer, ByRef found_yacht As Integer, ByRef didnt_find_yacht As String, ByVal jetnet_comp_id As Long, ByVal comp_name As String, ByRef results_string As String, ByVal page_num As Integer) As Boolean
    Scrape_This_Page_Yacht_Times = False

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim temp_yacht_id As String = ""
    Dim company_string As String = ""
    Dim didnt_find_temp As Integer = 0
    Dim found_temp As Integer = 0
    Dim insert_strings As String = ""
    Dim Insert_Query As String = ""
    Dim original_string_text As String = ""
    Dim related_articles_text As String = ""
    Dim temp_companies As Integer = 0
    Dim found_comp As Boolean = False


    Try


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text

      ' if the company id wasa put in wrong
      spot_to_find = InStr(string_text, "&nbsp;News", CompareMethod.Text)
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

        spot_to_find = InStr(string_text, "sidebar", CompareMethod.Text)
        If spot_to_find > 0 Then
          string_text = Left(string_text, spot_to_find - 1)


          spot_to_find = InStr(string_text, "<h1>Related Articles</h1>", CompareMethod.Text)
          If spot_to_find > 0 Then
            related_articles_text = Right(string_text, Len(string_text) - spot_to_find)
            string_text = Left(string_text, spot_to_find - 25)
          End If

          temp_companies = found_companies

          cut_up_this_section(False, string_text, jetnet_comp_id, found_yacht, found_news, found_companies)

          'if there was a company found 
          found_comp = False
          If found_companies > temp_companies Then
            temp_companies = found_companies
            found_comp = True
          End If

          spot_to_find = InStr(related_articles_text, "GA_googleFillSlot", CompareMethod.Text)
          If spot_to_find > 0 Then
            related_articles_text = Left(related_articles_text, spot_to_find - 17)
          End If
          '----------------------------------- RELATED ARTICLES SECTION------------------------------
          cut_up_this_section(True, related_articles_text, jetnet_comp_id, found_yacht, found_news, found_companies)
          '----------------------------------- RELATED ARTICLES SECTION------------------------------

          'if you foun a company then no need to add another
          If found_comp Then
            found_companies = temp_companies
          End If

        End If
      End If

    Catch ex As Exception
      Response.Write("")
    Finally

    End Try

  End Function
  Public Function Scrape_This_Page_Yacht_Times_NEW(ByVal link As String, ByRef found_news As Integer, ByRef found_companies As Integer, ByRef found_yacht As Integer, ByRef didnt_find_yacht As String, ByVal jetnet_comp_id As Long, ByVal comp_name As String, ByRef results_string As String, ByVal page_num As Integer) As Boolean
    Scrape_This_Page_Yacht_Times_NEW = False

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim temp_yacht_id As String = ""
    Dim company_string As String = ""
    Dim didnt_find_temp As Integer = 0
    Dim found_temp As Integer = 0
    Dim insert_strings As String = ""
    Dim Insert_Query As String = ""
    Dim original_string_text As String = ""
    Dim related_articles_text As String = ""
    Dim temp_companies As Integer = 0
    Dim found_comp As Boolean = False

    Dim array_split() As String
    Dim yacht_to_connect_id As Long = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Dim href_link As String = ""
    Dim temp_comp_id As Long = 0
    Dim exists_count As Integer = 0
    Dim first_run_yacht As Boolean = True
    Dim yacht_count As Integer = 0
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long
    Dim results_table As New DataTable


    Try


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text

      ' if the company id wasa put in wrong
      spot_to_find = InStr(string_text, "<h3>Company news", CompareMethod.Text)
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

        spot_to_find = InStr(string_text, "</section>", CompareMethod.Text)
        If spot_to_find > 0 Then
          string_text = Left(string_text, spot_to_find - 1)

          array_split = Split(string_text, "article clearfix")

          For i = 1 To array_split.Length - 1
            string_text = array_split(i)

            spot_to_find = InStr(string_text, "row-article-item", CompareMethod.Text)
            string_text = Right(string_text, Len(string_text) - spot_to_find - 17)

            spot_to_find = InStr(string_text, "href=""", CompareMethod.Text)
            string_text = Right(string_text, Len(string_text) - spot_to_find - 6)

            spot_to_find = InStr(string_text, ">", CompareMethod.Text)
            href_link = Left(string_text, spot_to_find - 2)

            href_link = "http://www.superyachttimes.com/" & href_link
            href_link = href_link

            spot_to_find = InStr(string_text, "<p class=""title"">", CompareMethod.Text)
            string_text = Right(string_text, Len(string_text) - spot_to_find - 16)

            spot_to_find = InStr(string_text, "</p>", CompareMethod.Text)
            title_string = Left(string_text, spot_to_find - 1)

            title_string = title_string

            spot_to_find = InStr(string_text, "<p class=""intro"">", CompareMethod.Text)
            string_text = Right(string_text, Len(string_text) - spot_to_find - 16)

            spot_to_find = InStr(string_text, "</p>", CompareMethod.Text)
            description_string = Left(string_text, spot_to_find - 1)

            description_string = description_string

            spot_to_find = InStr(string_text, "<p class=""time"">", CompareMethod.Text)
            string_text = Right(string_text, Len(string_text) - spot_to_find - 15)

            spot_to_find = InStr(string_text, "</p>", CompareMethod.Text)
            date_string = Left(string_text, spot_to_find - 1)

            date_string = Replace(date_string, "<i class=""icon-clock""></i>", "")

            date_string = date_string

            date_string = fix_date(date_string, "new_super_yacht_times")

            ' if date is less then 30 days old, then do
            If DateDiff(DateInterval.Day, CDate(date_string), Date.Now) < 30 Then

              If first_run_yacht Then
                first_run_yacht = False
                '-------------------------------------------------------------
                yacht_count = 0
                results_table = GET_YACHTS_FOR_COMPANY(jetnet_comp_id, "1") ' get yachts for that company
                If Not IsNothing(results_table) Then
                  If results_table.Rows.Count > 0 Then
                    For Each k As DataRow In results_table.Rows
                      If Not IsDBNull(k.Item("yt_yacht_name")) Then   'fill in arrays
                        yacht_name_array(yacht_count) = k.Item("yt_yacht_name")
                        yacht_id(yacht_count) = k.Item("yt_id")
                        yacht_count = yacht_count + 1
                      End If
                    Next
                  End If
                End If
                results_table.Dispose()
                '----------------------------------------------------------   
              End If

              ' if it has yachts and news----try to connect-----------------------------------------
              yacht_to_connect_id = 0
              temp_comp_id = 0
              If yacht_count > 0 Then
                For z = 0 To yacht_count - 1
                  If Len(Trim(yacht_name_array(z))) > 3 Then
                    If InStr(UCase(Replace(Replace(" " & Trim(title_string) & " ", "'", "''"), ",", "")), UCase(" " & Trim(yacht_name_array(z)) & " ")) > 0 Then
                      yacht_to_connect_id = yacht_id(z)
                      temp_comp_id = jetnet_comp_id
                      z = yacht_count + 1  ' END IT 
                    End If
                  End If
                Next
              End If
              ' if it has yachts and news----try to connect-----------------------------------------



              If Not CHECK_IF_NEWS_EXISTS(jetnet_comp_id, date_string, Replace(title_string, "'", "''"), False, "1") Then

                Call insert_into_news(date_string, title_string, description_string, href_link, yacht_to_connect_id, jetnet_comp_id, jetnet_comp_id, True, "1")

                If yacht_to_connect_id > 0 Then
                  found_yacht = found_yacht + 1
                End If

                If temp_comp_id > 0 Then
                  temp_companies = temp_companies + 1
                End If

                found_news = found_news + 1


                exists_count = 0
              Else
                exists_count = exists_count + 1
                If exists_count >= 3 Then
                  i = array_split.Length
                End If
              End If


            End If














          Next

        End If
      End If

    Catch ex As Exception
      Response.Write("")
    Finally

    End Try

  End Function
  Public Function Scrape_This_Page_Yacht_Times_NEW2(ByVal link As String, ByRef found_news As Integer, ByRef found_companies As Integer, ByRef found_yacht As Integer, ByRef didnt_find_yacht As String, ByVal jetnet_comp_id As Long, ByVal comp_name As String, ByRef results_string As String, ByVal page_num As Integer) As Boolean
    Scrape_This_Page_Yacht_Times_NEW2 = False

    link = "https://www.superyachttimes.com/companies/abeking-rasmussen/news"
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim temp_yacht_id As String = ""
    Dim company_string As String = ""
    Dim didnt_find_temp As Integer = 0
    Dim found_temp As Integer = 0
    Dim insert_strings As String = ""
    Dim Insert_Query As String = ""
    Dim original_string_text As String = ""
    Dim related_articles_text As String = ""
    Dim temp_companies As Integer = 0
    Dim found_comp As Boolean = False

    Dim array_split() As String
    Dim yacht_to_connect_id As Long = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Dim href_link As String = ""
    Dim temp_comp_id As Long = 0
    Dim exists_count As Integer = 0
    Dim first_run_yacht As Boolean = True
    Dim yacht_count As Integer = 0
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long
    Dim results_table As New DataTable


    Try


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text

      ' if the company id wasa put in wrong
      spot_to_find = InStr(string_text, "<h2>News</h2>", CompareMethod.Text)
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

        spot_to_find = InStr(string_text, "</section>", CompareMethod.Text)
        If spot_to_find > 0 Then
          string_text = Left(string_text, spot_to_find - 1)

          array_split = Split(string_text, "<article class")

          For i = 1 To array_split.Length - 1
            string_text = array_split(i)
 

            spot_to_find = InStr(string_text, "<a title=""Read more"" href=""", CompareMethod.Text)
            string_text = Right(string_text, Len(string_text) - spot_to_find - 27)

            spot_to_find = InStr(string_text, "</a>", CompareMethod.Text)
            href_link = Left(string_text, spot_to_find - 3)

            If InStr(href_link, "superyachttimes") = 0 Then
              href_link = "https://www.superyachttimes.com/" & Trim(href_link)
            End If

            spot_to_find = InStr(string_text, "<h1>", CompareMethod.Text)
            string_text = Right(string_text, Len(string_text) - spot_to_find - 3)


            spot_to_find = InStr(string_text, "</h1>", CompareMethod.Text)
            title_string = Left(string_text, spot_to_find - 1)


            spot_to_find = InStr(string_text, "<time datetime=""", CompareMethod.Text)
            string_text = Right(string_text, Len(string_text) - spot_to_find - 15)


            spot_to_find = InStr(string_text, "data-local", CompareMethod.Text)
            date_string = Left(string_text, spot_to_find - 10)

            spot_to_find = InStr(date_string, "T")
            If spot_to_find > 0 Then
              date_string = Left(date_string, spot_to_find - 1)
            End If

            date_string = change_date_format(date_string, 1, 2, 0)
            'date_string = fix_date(date_string, "new_super_yacht_times")

 


            If Not CHECK_IF_NEWS_EXISTS(jetnet_comp_id, date_string, Replace(title_string, "'", "''"), False, "1") Then

              Call insert_into_news(date_string, title_string, description_string, href_link, yacht_to_connect_id, jetnet_comp_id, jetnet_comp_id, True, "1")

              If yacht_to_connect_id > 0 Then
                found_yacht = found_yacht + 1
              End If

              If temp_comp_id > 0 Then
                temp_companies = temp_companies + 1
              End If

              found_news = found_news + 1


              exists_count = 0
            Else
              exists_count = exists_count + 1
              If exists_count >= 3 Then
                i = array_split.Length
              End If
            End If 

 

          Next

        End If
      End If

    Catch ex As Exception
      Response.Write("")
    Finally

    End Try

  End Function

  Public Function cut_up_this_section(ByVal related As Boolean, ByVal string_text As String, ByVal jetnet_comp_id As Long, ByRef found_yacht As Integer, ByRef found_news As Integer, ByRef found_companies As Integer)
    Dim yacht_to_connect_id As Long = 0


    Dim array_split() As String
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Dim original_string_text As String = ""
    Dim spot_to_find As Integer = 0
    Dim href_link As String = ""
    Dim Insert_Query As String = ""
    Dim temp_comp_id As Long = 0
    Dim temp_date_diff As String = ""
    Dim exists_count As Integer = 0
    Dim first_run_yacht As Boolean = True
    Dim yacht_count As Integer = 0
    Dim yacht_name_array(500) As String
    Dim yacht_id(500) As Long
    Dim results_table As New DataTable
    Dim temp_companies As Integer = 0



    array_split = Split(string_text, "<p>")

    For i = 1 To array_split.Length - 1
      string_text = array_split(i)
      original_string_text = string_text

      spot_to_find = InStr(string_text, "</span>", CompareMethod.Text)
      If spot_to_find > 0 Then
        date_string = Left(string_text, spot_to_find - 1)
        date_string = Replace(date_string, "<span class=""tag"">", "")

        date_string = fix_date(date_string, "")

        ' if date is less then 7 days old, then do
        If DateDiff(DateInterval.Day, CDate(date_string), Date.Now) < 7 Then


          ' first time in, get list of yachts
          If first_run_yacht Then
            first_run_yacht = False
            '-------------------------------------------------------------
            yacht_count = 0
            results_table = GET_YACHTS_FOR_COMPANY(jetnet_comp_id, "1") ' get yachts for that company
            If Not IsNothing(results_table) Then
              If results_table.Rows.Count > 0 Then
                For Each k As DataRow In results_table.Rows
                  If Not IsDBNull(k.Item("yt_yacht_name")) Then   'fill in arrays
                    yacht_name_array(yacht_count) = k.Item("yt_yacht_name")
                    yacht_id(yacht_count) = k.Item("yt_id")
                    yacht_count = yacht_count + 1
                  End If
                Next
              End If
            End If
            results_table.Dispose()
            '---------------------------------------------------------- 
          End If


          string_text = Right(string_text, Len(string_text) - spot_to_find - 7)
          spot_to_find = InStr(string_text, "</a>", CompareMethod.Text)
          string_text = Left(string_text, spot_to_find - 1)

          spot_to_find = InStr(string_text, ">", CompareMethod.Text)

          href_link = Left(string_text, spot_to_find - 2)
          title_string = Right(string_text, Len(string_text) - spot_to_find)


          spot_to_find = InStr(href_link, "/editorial/", CompareMethod.Text)
          If spot_to_find > 0 Then
            href_link = Right(href_link, Len(href_link) - spot_to_find)
            href_link = "http://www.superyachttimes.com/" & href_link
          End If


          description_string = Scrape_This_Detail_Page(href_link)
          description_string = clean_description(description_string)



          ' if it has yachts and news----try to connect-----------------------------------------
          yacht_to_connect_id = 0
          temp_comp_id = 0
          If yacht_count > 0 Then
            For z = 0 To yacht_count - 1
              If Len(Trim(yacht_name_array(z))) > 3 Then
                If InStr(UCase(Replace(Replace(" " & Trim(title_string) & " ", "'", "''"), ",", "")), UCase(" " & Trim(yacht_name_array(z)) & " ")) > 0 Then
                  yacht_to_connect_id = yacht_id(z)
                  temp_comp_id = jetnet_comp_id
                  z = yacht_count + 1  ' END IT 
                End If
              End If
            Next
          End If
          ' if it has yachts and news----try to connect-----------------------------------------

          If Not CHECK_IF_NEWS_EXISTS(jetnet_comp_id, date_string, Replace(title_string, "'", "''"), related, "1") Then


            temp_date_diff = DateDiff(DateInterval.Month, CDate(date_string), Date.Now())

            ' if less than 6 months, go in
            ' or if its not a related, go in
            ' or if it has found a comp id, go in
            If temp_date_diff < 3 Or (Not related) Or temp_comp_id > 0 Then
              Call insert_into_news(date_string, title_string, description_string, href_link, yacht_to_connect_id, jetnet_comp_id, temp_comp_id, related, "1")

              If yacht_to_connect_id > 0 Then
                found_yacht = found_yacht + 1
              End If

              If temp_comp_id > 0 Then
                temp_companies = temp_companies + 1
              End If

              found_news = found_news + 1


            End If
            exists_count = 0
          Else
            exists_count = exists_count + 1
            If exists_count >= 3 Then
              i = array_split.Length
            End If
          End If
        Else
          i = array_split.Length
        End If   ' end of if for if date > 7 difference

      End If

    Next

    If temp_companies > 0 Then
      found_companies = found_companies + 1
    End If

  End Function
  Public Function Insert_EMail_Queue_Record(ByVal strBody As String) As Boolean

    Dim strService As String

    Dim strInsert As String
    Dim bResults As Boolean

    Try

      'MySqlConn_JETNET.ConnectionString = "Data Source=128.1.21.200;Initial Catalog=jetnet_ra;Persist Security Info=True;User ID=sa;Password=moejive"
      MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
      ' Response.Write("<br>Insert_EMail_Queue_Record - CONN")
      MySqlConn_JETNET.Open()

      ' Response.Write("<br>Insert_EMail_Queue_Record - CONN OPEN")

      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 120



            strService = "EventListener"


            strBody = Replace(strBody, "'", "''")

      strInsert = "INSERT INTO EMail_Queue ("
      strInsert = strInsert & "emailq_service, "
      strInsert = strInsert & "emailq_replyname, "
      strInsert = strInsert & "emailq_replyemail, "
      strInsert = strInsert & "emailq_smtp_server, "
      strInsert = strInsert & "emailq_smtp_username, "
      strInsert = strInsert & "emailq_smtp_password, "
      strInsert = strInsert & "emailq_to, "
      strInsert = strInsert & "emailq_cc, "
      strInsert = strInsert & "emailq_bcc, "
      strInsert = strInsert & "emailq_subject, "
      strInsert = strInsert & "emailq_body, "
      strInsert = strInsert & "emailq_attachment, "
      strInsert = strInsert & "emailq_status, "
      strInsert = strInsert & "emailq_errormsg, "
      strInsert = strInsert & "emailq_html_flag) "

      strInsert = strInsert & "VALUES ("
      strInsert = strInsert & "'" & strService & "', "
            'strInsert = strInsert & "'JETNET Listener', "
            'strInsert = strInsert & "'event@jetnet.com', "
            'strInsert = strInsert & "'smtp.jetnet.com', "
            'strInsert = strInsert & "'jetnet@jetnet.com', "
            '      strInsert = strInsert & "'tentej123', "

            strInsert = strInsert & "'JETNET LLC', "
            strInsert = strInsert & "'event@jetnet.com', "
            strInsert = strInsert & "'smtp.office365.com', "
            strInsert = strInsert & "'event@jetnet.com', "
            strInsert = strInsert & "'$38Lstnr!KPz#2', "


            strInsert = strInsert & "'matt@jetnet.com', "
            strInsert = strInsert & "'', "
      strInsert = strInsert & "'', "
            strInsert = strInsert & "'News Results: " & Date.Now() & "', "
            strInsert = strInsert & "'" & strBody & "', "
      strInsert = strInsert & "'', "
      strInsert = strInsert & "'Open', "
      strInsert = strInsert & "'', "
      strInsert = strInsert & "'N') "

      MySqlCommand_JETNET.CommandText = strInsert
      MySqlCommand_JETNET.ExecuteNonQuery()
      MySqlCommand_JETNET.Dispose()
      Response.Write(strInsert)

      bResults = True


      Insert_EMail_Queue_Record = bResults


    Catch ex As Exception
    Finally
      MySqlConn_JETNET.Dispose()
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET = Nothing

      MySqlCommand_JETNET.Dispose()
      MySqlCommand_JETNET = Nothing

      bResults = False
      Insert_EMail_Queue_Record = bResults
    End Try

  End Function ' End Function Insert_EMail_Queue_Record

  Public Function insert_into_yacht_ypl(ByVal source_id As String, ByVal yacht_name As String, ByVal status As String, ByVal yacht_to_connect_id As Long, ByVal details1 As String, ByVal ypl_link As String) As Boolean
    insert_into_yacht_ypl = False

    Try



      Dim Insert_Query As String = ""

      Insert_Query = " Insert into Yacht_Publication_Log(ypl_source,  ypl_source_date,"
      Insert_Query = Insert_Query & "ypl_process_status, ypl_research_status, ypl_yacht_info, ypl_other_info "
      Insert_Query = Insert_Query & ", ypl_yacht_id, ypl_source_url, ypl_acct_rep, ypl_action_date  "
      Insert_Query = Insert_Query & " ) VALUES ( "
      Insert_Query = Insert_Query & "" & source_id & ",'" & ypl_start_date & "',"

      Insert_Query = Insert_Query & "'" & status & "',"

      If Trim(status) = "For Sale Found  – Exact Match" Then
        Insert_Query = Insert_Query & "'N',"
      ElseIf Trim(status) = "MMSI Match - Yacht Found" Then
        Insert_Query = Insert_Query & "'R',"
      Else
        Insert_Query = Insert_Query & "'O',"
      End If

      Insert_Query = Insert_Query & "'" & yacht_name & "',"
      Insert_Query = Insert_Query & "'" & details1 & "',"

      '" & href_link & "', "
      Insert_Query = Insert_Query & "'" & yacht_to_connect_id & "',  "
      Insert_Query = Insert_Query & "'" & ypl_link & "',  "
      Insert_Query = Insert_Query & "'PUB1'  "

      Insert_Query = Insert_Query & " , '1/1/1900')"
      '  Response.Write(Insert_Query & "<br>")


      If source_id = 8 Then ' marrine traffic 
        SqlCommand_YPL.CommandText = Insert_Query
        SqlCommand_YPL.ExecuteNonQuery()
        SqlCommand_YPL.Dispose()
      Else
        MySqlCommand_JETNET.CommandText = Insert_Query
        MySqlCommand_JETNET.ExecuteNonQuery()
        MySqlCommand_JETNET.Dispose()
      End If



      insert_into_yacht_ypl = True


    Catch ex As Exception

    End Try

  End Function
  Public Function Update_yacht_ypl_fields(ByVal source_id As String, ByVal yacht_name As String, ByVal status As String, ByVal yacht_to_connect_id As Long, ByVal details1 As String, ByVal ypl_link As String, ByVal found_pub_id As Long) As Boolean
    Update_yacht_ypl_fields = False

    Try

      Dim Update_Query As String = ""

 

      Update_Query = " Update Yacht_Publication_Log set "
      Update_Query &= " ypl_source_date = '" & ypl_start_date & "' "
      Update_Query &= " where ypl_source_url = '" & ypl_link & "' "
      Update_Query &= " and ypl_id = '" & found_pub_id & "' "

      ' UPDATE THE SOURCE DATE NO MATTER WHAT 

      If found_pub_id = 390 Or found_pub_id = 89 Or found_pub_id = 28 Or found_pub_id = 64 Or found_pub_id = 69 Or found_pub_id = 51 Or found_pub_id = 109 Then
        source_id = source_id
      End If

      MySqlCommand_JETNET.CommandText = Update_Query
      MySqlCommand_JETNET.ExecuteNonQuery()
      MySqlCommand_JETNET.Dispose()


      If asking_within_range = True Then  ' dont run second update if the asking is still within range 
        ' if its still within range 
      Else
        Update_Query = " Update Yacht_Publication_Log set "

        Update_Query &= "  ypl_process_status = '" & status & "' "

        If Trim(status) = "For Sale Found  – Exact Match" Then
          Update_Query &= " , ypl_research_status = 'N' "
        Else
          Update_Query &= " , ypl_research_status = 'O' "
        End If

        Update_Query &= ", ypl_other_info = '" & Trim(details1) & "' "
        Update_Query &= ", ypl_source_date = '" & ypl_start_date & "' "

        Update_Query &= " where ypl_source_url = '" & ypl_link & "' "
        Update_Query &= " and ypl_id = '" & found_pub_id & "' "

        ' and its cleared and the info has changed , or its or its still open
        ' if its cleared and the info hasnt changed it wont be updated
        If InStr(Trim(details1), "vs. $0") > 0 Then ' if it is vs 0 and has been cleared, then do not update if it was previously vs 0 
          Update_Query &= " and ( (ypl_research_status in ('C','D','N','I') and ypl_other_info <> '" & Trim(details1) & "' and ypl_other_info not like '%vs. $0%') or (ypl_research_status = 'O') ) "
        Else
          Update_Query &= " and ( (ypl_research_status in ('C','D','N','I') and ypl_other_info <> '" & Trim(details1) & "') or (ypl_research_status = 'O') ) "
        End If


        MySqlCommand_JETNET.CommandText = Update_Query
        MySqlCommand_JETNET.ExecuteNonQuery()
        MySqlCommand_JETNET.Dispose()

        Update_yacht_ypl_fields = True
      End If

    Catch ex As Exception

    End Try

  End Function
  Public Function update_yacht_ypl(ByVal status1 As String) As Boolean
    update_yacht_ypl = False

    Try

      Dim update_Query As String = ""

      If ypl_id > 0 Then

        update_Query = " Update Yacht_Publication_Log set ypl_source_date = '" & ypl_start_date & "' "

        If Trim(status1) <> "" Then
          update_Query &= " , ypl_process_status = '" & status1 & "' "
        End If

        update_Query &= " where ypl_id = " & ypl_id


        If ypl_id = 390 Or ypl_id = 89 Then
          ypl_id = ypl_id
        End If


        '  Response.Write(update_Query & "<br>")
        MySqlCommand_JETNET.CommandText = update_Query
        MySqlCommand_JETNET.ExecuteNonQuery()
        MySqlCommand_JETNET.Dispose()

        update_yacht_ypl = True
      End If

    Catch ex As Exception

    End Try

  End Function
  Public Function update_yacht_ypl_change_date(ByVal source_id As Long) As Boolean
    update_yacht_ypl_change_date = False

    Try

      Dim update_Query As String = ""

      update_Query = " Update Yacht_Publication_Log set ypl_process_status = 'Off Market', ypl_other_info = 'as of " & ypl_start_date & "', ypl_research_status = 'O' where ypl_source_date <> '" & ypl_start_date & "' and ypl_source = '" & source_id & "' and ypl_process_status not like '%News%' "
      update_Query &= " and ypl_process_status <> 'Off Market' and ypl_research_status not in ('D')  " ' cause if its already off market, dont want to put back to open

      '  Response.Write(update_Query & "<br>")
      MySqlCommand_JETNET.CommandText = update_Query
      MySqlCommand_JETNET.ExecuteNonQuery()
      MySqlCommand_JETNET.Dispose()

      update_yacht_ypl_change_date = True

    Catch ex As Exception

    End Try

  End Function

  Public Function insert_into_news(ByVal date_string As String, ByVal title_string As String, ByVal description_string As String, ByVal href_link As String, ByVal yacht_to_connect_id As Long, ByVal jetnet_comp_id As Long, ByVal temp_comp_id As Long, ByVal related As Boolean, ByVal src_id As String) As Boolean
    insert_into_news = False
    Dim insert_pub_record As Boolean = False

    Try

      If InStr(description_string, "sold") > 0 Or InStr(description_string, "sells") > 0 Or InStr(description_string, "sale") > 0 Then
        If InStr(description_string, "sales director") = 0 And InStr(description_string, "for sale") = 0 And InStr(description_string, "sales") = 0 And InStr(description_string, "tenders") = 0 And InStr(description_string, "berths") = 0 Then
          insert_pub_record = True
        End If
      End If


      If insert_pub_record = True Then
        If Trim(yacht_news_name) = "" Then
          yacht_news_name = Replace(title_string, "'", "''")
        End If
        If CHECK_IF_IN_PUB(href_link, yacht_to_connect_id, 0, "") = False Then
          Call insert_into_yacht_ypl(src_id, yacht_news_name, "Sold News Article", yacht_to_connect_id, "", href_link)
        End If
      End If

      ' if it sold, assume it might say for sale in it, dont need double 
      If insert_pub_record = False Then
        If InStr(title_string, "for sale") > 0 Then
          If InStr(description_string, "shipyard") = 0 And InStr(description_string, "berth") = 0 Then
            insert_pub_record = True
          End If
        End If


        If insert_pub_record = True Then
          If Trim(yacht_news_name) = "" Then
            yacht_news_name = Replace(title_string, "'", "''")
          End If
          If CHECK_IF_IN_PUB(href_link, yacht_to_connect_id, 0, "") = False Then
            Call insert_into_yacht_ypl(src_id, yacht_news_name, "For Sale News Article", yacht_to_connect_id, "", href_link)
          End If
        End If
      End If


      Dim Insert_Query As String = ""

      Insert_Query = " Insert into Yacht_News(ytnews_date,  ytnews_title,"
      Insert_Query = Insert_Query & "ytnews_description, ytnews_web_address, ytnews_source_id "
      Insert_Query = Insert_Query & ", ytnews_brand_name, ytnews_model_id, ytnews_yt_id, ytnews_comp_id, ytnews_action_date  "
      Insert_Query = Insert_Query & " ) VALUES ( "
      Insert_Query = Insert_Query & "'" & date_string & "','" & Replace(title_string, "'", "''") & "',"
      Insert_Query = Insert_Query & "'" & Replace(description_string, "'", "''") & "','" & href_link & "', "
      Insert_Query = Insert_Query & "'" & src_id & "', NULL, NULL, " & yacht_to_connect_id & ", "
      ' if its related news, only connect to company if you find companies yacht
      If related Then
        Insert_Query = Insert_Query & "'" & temp_comp_id & "'"
      Else
        Insert_Query = Insert_Query & "'" & jetnet_comp_id & "'"
      End If

      Insert_Query = Insert_Query & " , '1/1/1900')"
      '  Response.Write(Insert_Query & "<br>")
      MySqlCommand_JETNET.CommandText = Insert_Query
      MySqlCommand_JETNET.ExecuteNonQuery()
      MySqlCommand_JETNET.Dispose()

      insert_into_news = True


    Catch ex As Exception

    End Try

  End Function

  Public Function clean_description(ByVal description_string As String) As String

    description_string = Replace(description_string, "<p>", "")
    description_string = Replace(description_string, "</p>", "")
    description_string = Replace(description_string, "<p align=""center"">", "")
    description_string = Replace(description_string, "<p align=""left"">", "")
    description_string = Replace(description_string, "<p align=""right"">", "")
    description_string = Replace(description_string, "<strong>", "")
    description_string = Replace(description_string, "</strong>", "")
    description_string = Replace(description_string, "<em>", "")
    description_string = Replace(description_string, "</em>", "")
    description_string = Replace(description_string, "&nbsp;", " ")
    description_string = Replace(description_string, "&amp;", " ")
    description_string = Replace(description_string, "amp;", " ")
    description_string = Replace(description_string, "<br/>", " ")
    description_string = Replace(description_string, "&rsquo;", " ")

    clean_description = description_string

  End Function
  Public Function Scrape_This_Detail_Page(ByVal html_string As String) As String
    Scrape_This_Detail_Page = ""

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(html_string)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim spot_to_find As Integer = 0
    Dim i As Integer = 0
    Dim href_link As String = ""
    Dim original_string_text As String = ""


    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      original_string_text = string_text

      ' if the company id wasa put in wrong
      spot_to_find = InStr(string_text, "><strong>", CompareMethod.Text)
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find - 8)

        spot_to_find = InStr(string_text, "</strong>", CompareMethod.Text)
        If spot_to_find > 0 Then
          string_text = Left(string_text, spot_to_find - 1)
        End If

        string_text = Replace(string_text, "<strong>", "")

        Scrape_This_Detail_Page = string_text
      End If


      If InStr(Scrape_This_Detail_Page, "<img ") > 0 Then
        string_text = original_string_text
        spot_to_find = InStr(string_text, "><strong>", CompareMethod.Text)
        If spot_to_find > 0 Then
          string_text = Right(string_text, Len(string_text) - spot_to_find - 9)

          spot_to_find = InStr(string_text, "</strong>", CompareMethod.Text)
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 8)

            spot_to_find = InStr(string_text, "</div>", CompareMethod.Text)
            If spot_to_find > 0 Then
              string_text = Left(string_text, spot_to_find - 1)
              spot_to_find = InStr(string_text, "<p>", CompareMethod.Text)
              If spot_to_find > 0 Then
                string_text = Right(string_text, Len(string_text) - spot_to_find - 2)
                spot_to_find = InStr(string_text, "</p>", CompareMethod.Text)
                If spot_to_find > 0 Then
                  string_text = Left(string_text, spot_to_find - 1)
                End If


              End If
            End If

          End If
        End If

        Scrape_This_Detail_Page = string_text
      End If




    Catch ex As Exception
      Response.Write(string_text)
    Finally

    End Try

  End Function

  Public Function fix_date(ByVal date_string As String, ByVal site As String) As String
    Dim array_split() As String
    Dim i As Integer = 0
    Dim day1 As String = ""
    Dim month1 As String = ""
    Dim year1 As String = ""
    fix_date = ""


    date_string = Replace(date_string, "Monday", "")
    date_string = Replace(date_string, "Tuesday", "")
    date_string = Replace(date_string, "Wednesday", "")
    date_string = Replace(date_string, "Thursday", "")
    date_string = Replace(date_string, "Friday", "")
    date_string = Replace(date_string, "Saturday", "")
    date_string = Replace(date_string, "Sunday", "")
    date_string = Replace(date_string, ",", "")
    date_string = Trim(date_string)


    fix_date = Trim(date_string)
    array_split = Split(fix_date, " ")

    If Trim(site) = "superyachts" Then
      If array_split.Length = 3 Then
        For i = 0 To array_split.Length - 1
          If i = 0 Then
            month1 = array_split(i)
          ElseIf i = 1 Then
            day1 = array_split(i)
          ElseIf i = 2 Then
            year1 = array_split(i)
          End If
        Next
      End If
    ElseIf Trim(site) = "yacht_paging" Then
      If array_split.Length = 3 Then
        For i = 0 To array_split.Length - 1
          If i = 0 Then
            day1 = array_split(i)
          ElseIf i = 1 Then
            month1 = array_split(i)
          ElseIf i = 2 Then
            year1 = array_split(i)
          End If
        Next
      End If

    ElseIf Trim(site) = "new_super_yacht_times" Then
      array_split = Split(fix_date, "-")

      If array_split.Length = 3 Then
        For i = 0 To array_split.Length - 1
          If i = 0 Then
            day1 = array_split(i)
          ElseIf i = 1 Then
            month1 = array_split(i)
          ElseIf i = 2 Then
            year1 = array_split(i)
          End If
        Next
      End If
    Else
      If array_split.Length = 3 Then
        For i = 0 To array_split.Length - 1
          If i = 0 Then
            day1 = array_split(i)
          ElseIf i = 1 Then
            month1 = array_split(i)
          ElseIf i = 2 Then
            year1 = array_split(i)
          End If
        Next
      End If
    End If

    day1 = Replace(day1, "th", "")
    day1 = Replace(day1, "st", "")
    day1 = Replace(day1, "nd", "")
    If Len(Trim(day1)) = 1 Then
      day1 = "0" & Trim(day1)
    End If


    If Trim(month1) = "Jan" Or Trim(month1) = "January" Then
      month1 = "01"
    ElseIf Trim(month1) = "Feb" Or Trim(month1) = "February" Then
      month1 = "02"
    ElseIf Trim(month1) = "Mar" Or Trim(month1) = "March" Then
      month1 = "03"
    ElseIf Trim(month1) = "Apr" Or Trim(month1) = "April" Then
      month1 = "04"
    ElseIf Trim(month1) = "May" Then
      month1 = "05"
    ElseIf Trim(month1) = "Jun" Or Trim(month1) = "June" Then
      month1 = "06"
    ElseIf Trim(month1) = "Jul" Or Trim(month1) = "July" Then
      month1 = "07"
    ElseIf Trim(month1) = "Aug" Or Trim(month1) = "August" Then
      month1 = "08"
    ElseIf Trim(month1) = "Sep" Or Trim(month1) = "September" Then
      month1 = "09"
    ElseIf Trim(month1) = "Oct" Or Trim(month1) = "October" Then
      month1 = "10"
    ElseIf Trim(month1) = "Nov" Or Trim(month1) = "November" Then
      month1 = "11"
    ElseIf Trim(month1) = "Dec" Or Trim(month1) = "December" Then
      month1 = "12"
    End If


    fix_date = Trim(month1) & "/" & Trim(day1) & "/" & Trim(year1)

  End Function
  Public Function CHECK_IF_NEWS_EXISTS(ByVal comp_id As Long, ByVal date_temp As String, ByVal title_temp As String, ByVal related As Boolean, ByVal src_id As String) As Boolean
    CHECK_IF_NEWS_EXISTS = False

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader


    Dim Query As String = ""
    Try



      Query = "select * "
      Query &= " from Yacht_News  with (NOLOCK)  "
      Query &= " where ytnews_title = '" & title_temp & "' "

      'If Trim(src_id) <> "" Then
      '  Query &= " and ytnews_source_id = " & src_id & " "
      'End If

      'If related Then
      'Else
      '  Query &= " and ytnews_comp_id = " & comp_id & "  "
      'End If

      'If Trim(date_temp) = "" And src_id = "" Then
      '  Query &= " and (ytnews_date = '" & date_temp & "' or ytnews_date <= '" & FormatDateTime(Date.Now, DateFormat.ShortDate) & "' or ytnews_date is null)  "
      'Else
      '  Query &= " and ytnews_date = '" & date_temp & "' "
      'End If



      SqlConn.ConnectionString = MySqlConn_JETNET.ConnectionString
      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        If SqlReader.HasRows Then
          CHECK_IF_NEWS_EXISTS = True
        Else
          CHECK_IF_NEWS_EXISTS = False
        End If
      Catch constrExc As System.Data.ConstraintException
        CHECK_IF_NEWS_EXISTS = False
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

    Return CHECK_IF_NEWS_EXISTS

  End Function
  Public Sub CHECK_PUB_PRICE_RANGE(ByVal ypl_id As String, ByVal current_asking As String)

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader
    Dim Query As String = ""
    Try

      asking_within_range = False
      ypl_asking_price = ""

      Query = "select ypl_other_info "
      Query &= " from Yacht_Publication_Log  with (NOLOCK)  "
      Query &= " where ypl_id = '" & ypl_id & "' "

      SqlConn.ConnectionString = MySqlConn_JETNET.ConnectionString

      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        If SqlReader.HasRows Then
          SqlReader.Read()
          ypl_asking_price = SqlReader.Item("ypl_other_info")

          If Trim(ypl_asking_price) <> "" Then
            If InStr(Trim(ypl_asking_price), "vs") > 0 Then
              ypl_asking_price = Left(Trim(ypl_asking_price), InStr(Trim(ypl_asking_price), "vs") - 1)
              ypl_asking_price = Replace(ypl_asking_price, "$", "")
              ypl_asking_price = Replace(ypl_asking_price, ",", "")
            End If
          Else
          End If
        End If

        current_asking = Replace(current_asking, "$", "")
        current_asking = Replace(current_asking, ",", "")

        If Trim(current_asking) <> "" And Trim(ypl_asking_price) <> "" Then
          If IsNumeric(current_asking) = True And IsNumeric(ypl_asking_price) = True Then
            If CInt(ypl_asking_price) > CInt(current_asking) Then ' if the pub is bigger than the current asking on the site
              If (CInt(current_asking) + CInt(CInt(current_asking) * 0.05)) >= CInt(ypl_asking_price) Then ' if its within 5 percent its fine
                asking_within_range = True
              End If
            ElseIf CInt(ypl_asking_price) < CInt(current_asking) Then ' if the pub is less than than the current asking on the site
              If (CInt(ypl_asking_price) + CInt(CInt(ypl_asking_price) * 0.05)) >= CInt(current_asking) Then ' if its within 5 percent its fine
                asking_within_range = True
              End If
            Else
              ' they r equal 
            End If
          End If
        End If

      Catch constrExc As System.Data.ConstraintException
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

    Catch ex As Exception
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

  End Sub
  Public Function CHECK_IF_PUB_EXISTS(ByVal yacht_name As String, ByVal yt_id As Long, ByVal asking_price As String, ByVal src_id As String) As Boolean
    CHECK_IF_PUB_EXISTS = False

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader


    Dim Query As String = ""
    Try

      ypl_id = 0

      Query = "select * "
      Query &= " from Yacht_Publication_Log  with (NOLOCK)  "
      Query &= " where ypl_yacht_info = '" & yacht_name & "' "

      If src_id = "9" Then
        Query &= " and ypl_source = '9' "
      End If

      SqlConn.ConnectionString = MySqlConn_JETNET.ConnectionString


      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        If SqlReader.HasRows Then
          SqlReader.Read()
          ypl_id = SqlReader.Item("ypl_id")
          CHECK_IF_PUB_EXISTS = True
        Else
          CHECK_IF_PUB_EXISTS = False
        End If
      Catch constrExc As System.Data.ConstraintException
        CHECK_IF_PUB_EXISTS = False
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

    Return CHECK_IF_PUB_EXISTS
  End Function
  Public Function CHECK_IF_Yacht_EXISTS(ByVal yacht_name As String, ByVal imo_number As String, ByVal mmsi_number As String, ByVal yacht_id As Long) As Boolean
    CHECK_IF_Yacht_EXISTS = False


    Dim Query As String = ""
    Dim tcount As Integer = 0

    Try

      If yacht_id > 0 Then
      Else
        yacht_id_sy = 0
        ys_mmsi = ""
        ys_imo = ""
        ' if there is a yacht name then clear the dups
        If Trim(yacht_name) <> "" Then
          ys_dups = ""
        End If
      End If


      Query = "select * "
      Query &= " from Yacht with (NOLOCK)  "

      If Trim(yacht_name) <> "" And Trim(imo_number) <> "" Then
        Query &= " where yt_yacht_name = '" & yacht_name & "' and  yt_imo_nbr = '" & imo_number & "'  "
      ElseIf Trim(imo_number) <> "" Then
        Query &= " where yt_imo_nbr = '" & imo_number & "' "
      ElseIf Trim(yacht_name) <> "" Then
        Query &= " where yt_yacht_name = '" & yacht_name & "'  "
      ElseIf Trim(mmsi_number) <> "" Then
        Query &= " where yt_mmsi_mobile_nbr = '" & mmsi_number & "'  "
      ElseIf yacht_id > 0 Then
        Query &= " where yt_mmsi_mobile_nbr <> '' and yt_id = " & yacht_id & " "
      End If

      Query &= " and yt_journ_id = 0 "
      Query &= " and yt_lifecycle_id <> '4' "

      SqlCommand_YPL.CommandText = Query.ToString
      SqlReader_YPL = SqlCommand_YPL.ExecuteReader()

      Try
        If SqlReader_YPL.HasRows Then

          Do While SqlReader_YPL.Read()
            ' if we r doing yacht id search, dont set the variables
            If yacht_id > 0 Then
            Else
              yacht_id_sy = SqlReader_YPL.Item("yt_id")

              If Not IsDBNull(SqlReader_YPL.Item("yt_mmsi_mobile_nbr")) Then
                If Trim(SqlReader_YPL.Item("yt_mmsi_mobile_nbr")) = Trim(mmsi_number) Then
                  match_has_mmsi = True
                End If
                ys_mmsi = SqlReader_YPL.Item("yt_mmsi_mobile_nbr")
              End If

              If Not IsDBNull(SqlReader_YPL.Item("yt_imo_nbr")) Then
                ' If Trim(SqlReader.Item("yt_imo_nbr")) = Trim(imo_number) Then
                ys_imo = SqlReader_YPL.Item("yt_imo_nbr")
                ' End If
              End If
            End If

            tcount = tcount + 1
          Loop
          CHECK_IF_Yacht_EXISTS = True
        Else
          CHECK_IF_Yacht_EXISTS = False
        End If

        If tcount > 1 Then
          CHECK_IF_Yacht_EXISTS = False ' too many 
          ys_dups = "Duplicates Found"
        End If

      Catch constrExc As System.Data.ConstraintException
        CHECK_IF_Yacht_EXISTS = False
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      Finally
        SqlReader_YPL.Close()
      End Try

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlCommand_YPL.Dispose()
    End Try

    Return CHECK_IF_Yacht_EXISTS
  End Function

  Public Function GET_NEWS_FOR_COMPANY(ByVal comp_id As Long) As DataTable
    Dim atemptable As New DataTable

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader


    Dim Query As String = ""
    Try



      Query = "select ytnews_id, ytnews_title "
      Query &= " from Yacht_News  with (NOLOCK)  "
      Query &= " inner join Yacht_News_Source with (NOLOCK) on  ytnewssrc_id = ytnews_source_id "
      Query &= " where ytnews_source_id = 1 "
      Query &= " and ytnews_comp_id = " & comp_id & "  "
      ' Query &= " and (ytnews_yt_id  is null and ytnews_yt_id <> 0 )   "  changed from this, should only get articles with no yacht attached
      Query &= " and (ytnews_yt_id  is null or ytnews_yt_id = 0 )   "  ' 

      Query &= " order by ytnews_id, ytnews_title "


      SqlConn.ConnectionString = MySqlConn_JETNET.ConnectionString
      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        atemptable.Load(SqlReader)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

    Return atemptable

  End Function


  Public Function GET_YACHTS_FOR_COMPANY(ByVal comp_id As Long, ByVal src_id As String) As DataTable
    Dim atemptable As New DataTable

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader


    Dim Query As String = ""
    Try



      Query = "select distinct yt_yacht_name, yt_id"
      Query &= " from Yacht_Source_Mapping  with (NOLOCK)  "
      Query &= " inner join Yacht_News_Source with (NOLOCK) on ytmap_ytnewssrc_id = ytnewssrc_id"
      Query &= " inner join Company  with (NOLOCK)  on ytmap_jetnet_comp_id = comp_id and comp_journ_id = 0"
      Query &= " inner join yacht_reference  with (NOLOCK)  on yr_comp_id = comp_id and yr_journ_id = 0  "
      Query &= " inner join yacht  with (NOLOCK)  on yr_yt_id = yt_id and yt_journ_id = 0 "
      Query &= " inner join Yacht_Model with (NOLOCK) on ym_model_id = yt_model_id "
      Query &= " where ytmap_ytnewssrc_id = " & src_id & " "
      Query &= " and (comp_id = " & comp_id & " and yr_contact_type in ('Y6'))  or (ym_mfr_comp_id = " & comp_id & ")  "
      Query &= " and len(yt_yacht_name) > 1 "
      Query &= " order by yt_yacht_name desc, yt_id "


      SqlConn.ConnectionString = MySqlConn_JETNET.ConnectionString
      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        atemptable.Load(SqlReader)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

    Return atemptable

  End Function
  Public Function GET_LAST_DATE(ByVal table_name As String, ByVal field_name As String, ByVal where_statement As String, ByVal select_internal_ip As String) As String
    GET_LAST_DATE = ""

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader


    Dim Query As String = ""
    Try

      ypl_id = 0

      Query = "select top 1 " & field_name & " from " & table_name & "  with (NOLOCK)   "

      If Trim(where_statement) <> "" Then
        Query &= " Where " & where_statement
      End If

      Query &= " order by " & field_name & " desc"

      If Trim(select_internal_ip) = "Y" Then
        SqlConn.ConnectionString = Inhouse_Live_Connection
      Else
        SqlConn.ConnectionString = JETNET_LIVE_SQL_CONN
      End If

      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        If SqlReader.HasRows Then
          SqlReader.Read()
          GET_LAST_DATE = SqlReader.Item("" & field_name & "")
        Else
        End If
      Catch constrExc As System.Data.ConstraintException
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

  End Function
  Public Function GET_LOG_DATE(ByVal run_type As String, ByVal select_internal_ip As String) As String
    GET_LOG_DATE = ""

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader


    Dim Query As String = ""
    Try

      ypl_id = 0

      Query = "select top 1 evtl_date from EventLog  with (NOLOCK)   "

      Query &= " Where evtl_message= '" & run_type & "' "

      Query &= " order by evtl_date desc"

      If Trim(select_internal_ip) = "Y" Then
        SqlConn.ConnectionString = Inhouse_Live_Connection
      Else
        SqlConn.ConnectionString = JETNET_LIVE_SQL_CONN
      End If

      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      SqlCommand.CommandText = Query.ToString
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        If SqlReader.HasRows Then
          SqlReader.Read()
          GET_LOG_DATE = SqlReader.Item("evtl_date")
        Else
        End If
      Catch constrExc As System.Data.ConstraintException
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

  End Function
  Public Function GET_DATA_INTEGRITY_INFO() As String
    GET_DATA_INTEGRITY_INFO = ""

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader
    Dim atemptable As New DataTable
    Dim atemptable2 As New DataTable
    Dim temp_name As String = ""
    Dim left_names(1000) As String
    Dim left_count(1000) As String
    Dim left_date(1000) As String
    Dim unique_date(1000) As String
    Dim temp_count As String = ""
    Dim array_count As Integer = 0
    Dim i As Integer = 0
    Dim k As Integer = 0
    Dim temp_table As String = ""
    Dim last_date As String = ""
    Dim unique_date_count As Integer = 0
    Dim found_spot As Boolean = False

    Dim Query As String = ""
    Try

      ypl_id = 0

      Query = "select evtl_message, evtl_date from EventLog  with (NOLOCK)   "

      Query &= " Where evtl_type = 'Data Integrity' "

      Query &= " and evtl_date > getdate() - 90 "
      Query &= " order by evtl_date asc "
 
      SqlConn.ConnectionString = JETNET_LIVE_SQL_CONN 

      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 60

      SqlCommand.CommandText = Query
      SqlReader = SqlCommand.ExecuteReader()

      Try
        atemptable.Load(SqlReader)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

      SqlReader.Close()

      If atemptable.Rows.Count > 0 Then
        For Each r As DataRow In atemptable.Rows

          temp_name = r.Item("evtl_message")
          If InStr(Trim(temp_name), "(") > 0 Then
            temp_count = Right(Trim(temp_name), Len(Trim(temp_name)) - InStr(Trim(temp_name), "("))
            temp_count = Replace(temp_count, ")", "")
            temp_name = Left(Trim(temp_name), InStr(Trim(temp_name), "(") - 1) 
          End If
          left_names(array_count) = temp_name
          left_count(array_count) = temp_count
          left_date(array_count) = FormatDateTime(r.Item("evtl_date"), DateFormat.ShortDate)
          array_count = array_count + 1 
        Next
      End If



      Query = "  SELECT distinct sqlrep_title, sqlrep_id  FROM SQL_Report WITH(NOLOCK)"
      Query &= " WHERE sqlrep_level = 'JETNET' AND sqlrep_sub_id = 0 "  
      Query &= " and sqlrep_type = 'Data Integrity' "
      Query &= " order by sqlrep_title asc "

      SqlCommand.CommandText = Query
      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        atemptable2.Load(SqlReader)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        ' aError = "Error in Get_CRM_VIEW_Prospects load datatable" + constrExc.Message
      End Try

      temp_table = ("<table cellpadding='4' cellspacing='0' border='1'>")

      temp_table &= ("<tr>")
      temp_table &= ("<td>Data Integrity Report</td>")

      For i = 0 To array_count - 1
        If Trim(last_date) = "" Or Trim(last_date) <> Trim(left_date(i)) Then
          temp_table &= ("<td>" & left_date(i) & "</td>")
          unique_date(unique_date_count) = left_date(i)
          unique_date_count = unique_date_count + 1
        End If
        last_date = Trim(left_date(i))
      Next

      temp_table &= ("<tr>")

      If atemptable2.Rows.Count > 0 Then
        For Each r As DataRow In atemptable2.Rows

          temp_table &= ("<tr>")

          If Len(Trim(r.Item("sqlrep_title"))) > 60 Then
            temp_table &= ("<td><A href='default.aspx?rep_id=" & Trim(r.Item("sqlrep_id")) & "'><font size='-1'>" & Left(Trim(r.Item("sqlrep_title")), 60))
            temp_table &= ("<Br/>")
            temp_table &= (Right(Trim(r.Item("sqlrep_title")), Len(Trim(r.Item("sqlrep_title"))) - 60))
            temp_table &= ("</font></a></td>")
          Else
            temp_table &= ("<td><A href='default.aspx?rep_id=" & Trim(r.Item("sqlrep_id")) & "'><font size='-1'>" & Trim(r.Item("sqlrep_title")) & "</font></a></td>")
          End If


          last_date = ""
          found_spot = False
          For i = 0 To unique_date_count - 1
            found_spot = False
            For k = 0 To array_count - 1
              If Trim(left_date(k)) = Trim(unique_date(i)) And Trim(left_names(k)) = Trim(r.Item("sqlrep_title")) Then
                temp_table &= ("<td align='right'>" & FormatNumber(left_count(k), 0) & "</td>")
                k = array_count
                found_spot = True
              End If
            Next
            If found_spot = False Then
              temp_table &= ("<td align='right'>-&nbsp;</td>")
            End If
          Next


          temp_table &= ("<tr>")

        Next
      End If




    


      temp_table &= ("</table>")

      Me.integrity_label.Text = temp_table

    Catch ex As Exception
      Return Nothing
      '  aError = "Error in Get_CRM_VIEW_Prospects()" + ex.Message
    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

  End Function





#Region "JETNET_NEWS_FUNCTIONS"

  Function ABI_News_Scraper_Gulfstream(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_Gulfstream = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text



      array_split = Split(string_text, "<td width=" & Chr("34") & "130" & Chr("34") & "><a href=")

      For i = 1 To array_split.Length - 1
        string_text = array_split(i)
        spot_to_find = InStr(string_text, Chr("34") & ">", CompareMethod.Text)
        link_to_go = Left(string_text, spot_to_find - 1)
        link_to_go = Right(link_to_go, Len(link_to_go) - 1)

        string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
        spot_to_find = InStr(string_text, "</a></td>", CompareMethod.Text)

        date_string = Left(string_text, spot_to_find - 1)

        string_text = Right(string_text, Len(string_text) - spot_to_find - 9)

        spot_to_find = InStr(string_text, "</td>", CompareMethod.Text)

        string_text = Left(string_text, spot_to_find - 1)

        spot_to_find = InStr(string_text, "<td>", CompareMethod.Text)

        title_string = Right(string_text, Len(string_text) - spot_to_find - 3)
        title_string = Replace(title_string, "'", "")

        If InStr(link_to_go, "http") = 0 Then
          link_to_go = "http://" & link_to_go
        End If


        If record_exists(date_string, title_string, link_to_go) = False Then



          description_string = Get_Description_From_Gulfstream(link_to_go, title_string)
          description_string = replace_all_chars(description_string)
          description_string = Left(description_string, 500)



          insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
          insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
          insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
          insert_string = insert_string & " ) VALUES ( "
          insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
          insert_string = insert_string & link_to_go & "','" & id & "','GULFSTREAM','',"
          insert_string = insert_string & "'',0"
          insert_string = insert_string & " ) "

          Response.Write(insert_string & "<br><br>")




          'SETUP AND EXECUTE THE SQL INSERT COMMAND
          Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

          sqlComm.ExecuteNonQuery()
          sqlComm.Dispose()

          ABI_News_Scraper_Gulfstream = ABI_News_Scraper_Gulfstream + 1

        End If
      Next

    Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try

  End Function
  Function Get_Description_From_Gulfstream(ByVal link As String, ByVal temp_title As String) As String
    Get_Description_From_Gulfstream = ""
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim orig_string_text As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim i As Integer = 0

    Try


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      orig_string_text = string_text

      spot_to_find = InStr(string_text, "<p><strong>")
      string_text = Right(string_text, Len(string_text) - spot_to_find + 1)

      If Left(string_text, 26) = "<p><strong>NOTE TO EDITORS" Then
        spot_to_find = 0
        string_text = orig_string_text
        spot_to_find = InStr(string_text, "<p><b>")
        If spot_to_find = 0 Then
          spot_to_find = InStr(string_text, "<h1>")
        End If
        string_text = Right(string_text, Len(string_text) - spot_to_find + 1)
        string_text = Replace(string_text, "<br/>", "")
        string_text = Replace(string_text, temp_title, "")
      End If


      spot_to_find = InStr(string_text, "<div id=")
      string_text = Left(string_text, spot_to_find)


      Get_Description_From_Gulfstream = string_text
    Catch ex As Exception
      Response.Write(ex)
    End Try
  End Function
  Function ABI_News_Scraper_Hawker_Beach(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_Hawker_Beach = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text

      string_text = Right(string_text, Len(string_text) - InStr(string_text, "News &amp; Press"))

      string_text = Right(string_text, Len(string_text) - InStr(string_text, "News &amp; Press"))

      array_split = Split(string_text, "<div class=" & Chr("34") & "news_link" & Chr("34") & ">")

      For i = 1 To array_split.Length - 1

        string_text = array_split(i)
        spot_to_find = InStr(string_text, "</div>")
        date_string = Left(string_text, spot_to_find - 1)

        spot_to_find = InStr(string_text, "<div class=" & Chr("34") & "news_title_link" & Chr("34") & "><a href=" & Chr("34"))
        string_text = Right(string_text, Len(string_text) - spot_to_find)


        spot_to_find = InStr(string_text, "title=")
        title_string = Right(string_text, Len(string_text) - spot_to_find - 6)
        link_to_go = Left(string_text, spot_to_find - 3)

        spot_to_find = InStr(title_string, ">")
        title_string = Left(title_string, spot_to_find - 1)
        title_string = replace_all_chars(title_string)



        spot_to_find = InStr(link_to_go, "href=")
        link_to_go = Right(link_to_go, Len(link_to_go) - spot_to_find - 5)



        If InStr(link_to_go, "http") = 0 Then
          link_to_go = "http://" & link_to_go
        End If


        If record_exists(date_string, title_string, link_to_go) = False Then


          description_string = Get_Description_From_Hawker_Beach(link_to_go, title_string, date_string)

          date_string = change_date_format(date_string, 0, 1, 2)

          description_string = replace_all_chars(description_string)
          description_string = Left(description_string, 500)


          insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
          insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
          insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
          insert_string = insert_string & " ) VALUES ( "
          insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
          insert_string = insert_string & link_to_go & "','" & id & "','HAWKER','',"
          insert_string = insert_string & "'',0"
          insert_string = insert_string & " ) "

          Response.Write(insert_string & "<br><br>")




          'SETUP AND EXECUTE THE SQL INSERT COMMAND
          Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

          sqlComm.ExecuteNonQuery()
          sqlComm.Dispose()
          ABI_News_Scraper_Hawker_Beach = ABI_News_Scraper_Hawker_Beach + 1
        End If
      Next

    Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try

  End Function
  Function Get_Description_From_Hawker_Beach(ByVal link As String, ByVal temp_title As String, ByRef date_string As String) As String
    Get_Description_From_Hawker_Beach = ""
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim orig_string_text As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim i As Integer = 0

    Try


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      orig_string_text = string_text

      spot_to_find = InStr(string_text, "<div id=" & Chr("34") & "single_news" & Chr("34") & ">")
      spot_to_find2 = InStr(string_text, "###")

      string_text = Left(string_text, spot_to_find2 - 1)
      string_text = Right(string_text, Len(string_text) - spot_to_find - 22)

      If InStr(string_text, " – ") > 0 Then
        date_string = Left(string_text, InStr(string_text, " – "))
        date_string = Trim(date_string)

        If InStr(Trim(date_string), ".") > 0 Then
          date_string = Right(Trim(date_string), Len(Trim(date_string)) - InStr(Trim(date_string), "."))
        End If

        If InStr(Trim(date_string), " ") > 0 Then
          date_string = Right(Trim(date_string), Len(Trim(date_string)) - InStr(Trim(date_string), " "))
        End If

      End If


      Get_Description_From_Hawker_Beach = string_text
    Catch ex As Exception
      Response.Write(ex)
    End Try
  End Function
  Function ABI_News_Scraper_Cessna(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_Cessna = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text



      array_split = Split(string_text, "<td class=" & Chr("34") & "date last" & Chr("34") & ">")

      For i = 1 To array_split.Length - 1

        string_text = array_split(i)

        spot_to_find = InStr(string_text, "</td>")
        date_string = Left(string_text, spot_to_find - 1)
        date_string = Replace(date_string, ".", "/")

        string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

        spot_to_find = InStr(string_text, "href=" & Chr("34"))

        string_text = Right(string_text, Len(string_text) - spot_to_find - 5)


        spot_to_find = InStr(string_text, ">")
        link_to_go = Left(string_text, spot_to_find - 2)


        string_text = Right(string_text, Len(string_text) - spot_to_find)

        spot_to_find = InStr(string_text, "</a>")
        title_string = Left(string_text, spot_to_find - 1)
        title_string = Replace(title_string, "'", "")

        spot_to_find = InStr(title_string, ";")
        If (spot_to_find > 0 And Len(title_string) > 119) Then
          title_string = Left(title_string, spot_to_find)
        End If
        title_string = Left(title_string, 119)


        If InStr(link_to_go, "../") > 0 Then
          link_to_go = "http://www.cessna.com/" & Replace(link_to_go, "../", "")
        End If

        If Right(Trim(date_string), 4) = "2012" Then


          If record_exists(date_string, title_string, link_to_go) = False Then


            description_string = Get_Description_From_Cessna(link_to_go, title_string)
            description_string = replace_all_chars(description_string)
            description_string = Left(description_string, 500)

            insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
            insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
            insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
            insert_string = insert_string & " ) VALUES ( "
            insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
            insert_string = insert_string & link_to_go & "','" & id & "','CESSNA','',"
            insert_string = insert_string & "'',0"
            insert_string = insert_string & " ) "

            ' Response.Write(insert_string & "<br><br>")




            'SETUP AND EXECUTE THE SQL INSERT COMMAND
            Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

            sqlComm.ExecuteNonQuery()
            sqlComm.Dispose()

            ABI_News_Scraper_Cessna = ABI_News_Scraper_Cessna + 1

          End If
        End If
      Next

    Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try

  End Function
  Function Get_Description_From_Cessna(ByVal link As String, ByVal temp_title As String) As String
    Get_Description_From_Cessna = ""
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim orig_string_text As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim i As Integer = 0

    Try


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      orig_string_text = string_text

      spot_to_find = InStr(string_text, "<div class=" & Chr("34") & "FCKContent" & Chr("34") & ">")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 21)

      spot_to_find = InStr(string_text, "<p>###</p>")
      If spot_to_find > 0 Then
        string_text = Left(string_text, spot_to_find - 1)
      Else
        spot_to_find = InStr(string_text, "</div>")
        string_text = Left(string_text, spot_to_find - 1)
      End If




      Get_Description_From_Cessna = string_text

    Catch ex As Exception
      Response.Write(ex)
    Finally

    End Try
  End Function

  'Function Scrape_Aircraft_Exchange(ByVal link As String, ByVal id As Integer) As Long
  '  Scrape_Aircraft_Exchange = 0
  '  Dim Str As System.IO.Stream
  '  Dim srRead As System.IO.StreamReader
  '  Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
  '  Dim resp As System.Net.WebResponse = req.GetResponse
  '  Dim string_text As String = ""
  '  Dim string_text2 As String = ""
  '  Dim spot_to_find As Integer = 0
  '  Dim spot_to_find2 As Integer = 0
  '  Dim array_split() As String
  '  Dim i As Integer = 0
  '  Dim date_string As String = ""
  '  Dim description_string As String = ""
  '  Dim title_string As String = ""
  '  Dim insert_string As String = ""
  '  Dim link_to_go As String = ""
  '  Try


  '    Str = resp.GetResponseStream
  '    srRead = New System.IO.StreamReader(Str)
  '    ' read all the text 
  '    string_text = srRead.ReadToEnd().ToString
  '    string_text = string_text


  '  Catch ex As Exception
  '    Response.Write(insert_string)
  '  Finally
  '  End Try
  'End Function
  Function scrape_controller_html()

        'Dim applicationDirectory = Path.GetDirectoryName("C:\Users\matt\Desktop\Work Documents\")
        'Dim myFile As String = Path.Combine(applicationDirectory, "Page53.html")
        'Dim temp = New Uri("file:///" + myFile)

        'Response.Write(temp.ToString)
        '  yt_table &= "<tr><td><b>Aircraft</b></td><td><b>AC ID</b></td><td><b>Owner</b></td><td><b>Status</b></td><td><b>Action</b></td></tr>"



        'https://www.controller.com/listings/aircraft/for-sale/list/category/3/jet-aircraft/?sortorder=27&SCF=False%2f&page=1          
        'https://www.controller.com/listings/aircraft/for-sale/list/category/8/turboprop-aircraft/?sortorder=27&SCF=False%2f&page=1         
        'https://www.controller.com/listings/aircraft/for-sale/list/category/9/piston-twin-aircraft/?sortorder=27&SCF=False%2f&page=1      
        'https://www.controller.com/listings/aircraft/for-sale/list/category/5/piston-helicopters/?sortorder=27&SCF=False%2f&page=1          
        'https://www.controller.com/listings/aircraft/for-sale/list/category/7/turbine-helicopters/?sortorder=27&SCF=False%2f&page=1       



        Dim skip_this As String = "Y"
        Dim skip_3_ina_row As Integer = 0

        Try


            MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
            'MySqlConn_JETNET.ConnectionString = Inhouse_Test_Connection
            'MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN
            MySqlConn_JETNET.Open()
      MySqlCommand_JETNET.Connection = MySqlConn_JETNET
      MySqlCommand_JETNET.CommandType = CommandType.Text
      MySqlCommand_JETNET.CommandTimeout = 60

      Call insert_into_eventlog("Aircraft Pubs Started", "Research Assistant")

      ypl_start_date = Date.Now

      Call Find_Naughty_Models()

            'Dim user As New Dictionary(Of String, String)
            'user.Add("username", "Admiinistrator")
            'user.Add("password", "moejive")

            'Dim ww As New Web
            'WebRequest(Request = WebRequest.Create("http://server/file.xml"))


            For i = 1 To 150

                Try

                    Response.Write("<br/>Started Page: " & i)

                    Using sr As New StreamReader("D:\jetnetassistant\CONTROLLER\" & i & ".htm")
                        '  Using sr As New StreamReader("C:\Controller\" & i & ".htm")
                        scrape_for_controller(sr, i)
                    End Using



                    Response.Write("<br/>Finished Page: " & i)



                    non_zero_Count = non_zero_Count
                    zero_Count = zero_Count


                Catch ex As Exception
                    skip_3_ina_row = skip_3_ina_row + 1
                    Response.Write("<br/>")
                    Response.Write(ex)
                    If skip_3_ina_row = 3 Then
                        i = 150
                    End If
                End Try
            Next

            skip_3_ina_row = 0

            For i = 1 To 150

                Try

                    Response.Write("<br/>Started Page: " & i)

                    Using sr As New StreamReader("D:\jetnetassistant\CONTROLLER\" & i & ".html")
                        '      Using sr As New StreamReader("C:\Controller\" & i & ".html")
                        scrape_for_controller(sr, i)
                    End Using



                    Response.Write("<br/>Finished Page: " & i)



                    non_zero_Count = non_zero_Count
                    zero_Count = zero_Count


                Catch ex As Exception
                    skip_3_ina_row = skip_3_ina_row + 1
                    Response.Write("<br/>")
                    Response.Write(ex)
                    If skip_3_ina_row = 3 Then
                        i = 150
                    End If
                End Try
            Next



            Call Insert_EMail_Queue_Record(yt_table)

      Call insert_into_eventlog("Aircraft Pubs Finished", "Research Assistant")

            Response.Write("<br/>Finished")

        Catch ex As Exception
      'Response.Write(ex)
    Finally

      MySqlConn_JETNET.Dispose()
      MySqlConn_JETNET.Close()
      MySqlConn_JETNET = Nothing

    End Try


    ' string applicationDirectory = Path.GetDirectoryName(Application.ExecutablePath);
    'string myFile = Path.Combine(applicationDirectory, "Sample.html");
    'webMain.Url = new Uri("file:///" + myFile);


  End Function
    Function Scrape_Ebay_Page(ByVal link As String, ByVal id As Integer) As Long
        Scrape_Ebay_Page = 0
        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader
        Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(link)
        ' Dim req As System.Net.WebClient
        '  req.Method = "PUT"
        Dim resp As System.Net.HttpWebResponse = req.GetResponse
        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim i As Integer = 0
        Dim date_string As String = ""
        Dim description_string As String = ""
        Dim title_string As String = ""
        Dim insert_string As String = ""
        Dim link_to_go As String = ""
        Try


            ' Dim webClient As New System.Net.WebClient
            ' Dim result As String = webClient.DownloadString(link)

            'Try
            '  Dim fr As System.Net.HttpWebRequest
            '  Dim targetURI As New Uri(link)

            '  fr = DirectCast(System.Net.HttpWebRequest.Create(targetURI), System.Net.HttpWebRequest)
            '  If (fr.GetResponse().ContentLength > 0) Then
            '    Dim str As New System.IO.StreamReader(fr.GetResponse().GetResponseStream())
            '    Response.Write(str.ReadToEnd())
            '  str.Close(); 
            '  End If
            'Catch ex As System.Net.WebException
            '  'Error in accessing the resource, handle it
            'End Try


            '  Dim result As String = req.DownloadString(link)


            'Using client As New Net.WebClient
            '  Dim reqparm As New Specialized.NameValueCollection
            '  ' reqparm.Add("page", "1")
            '  Dim responsebytes = client.UploadValues(link, "POST", reqparm)
            'End Using


            Str = resp.GetResponseStream
            srRead = New System.IO.StreamReader(Str)
            ' read all the text 
            string_text = srRead.ReadToEnd().ToString
            string_text = string_text


        Catch ex As Exception
            Response.Write(insert_string)
        Finally
        End Try
    End Function
    Function Scraper_Template(ByVal link As String) As Long
        Scraper_Template = 0

        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader

        System.Net.ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType)
        Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(link)
        ' Dim req As System.Net.WebClient
        '  req.Method = "PUT"
        Dim resp As System.Net.HttpWebResponse = req.GetResponse
        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim i As Integer = 0
        Dim date_string As String = ""
        Dim description_string As String = ""
        Dim title_string As String = ""
        Dim insert_string As String = ""
        Dim link_to_go As String = ""
        Dim comp_name As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_ac_name As String = ""
        Dim temp_make As String = ""
        Dim temp_temp As String = ""
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split2() As String
        Dim temp_manu As String = ""

        Str = resp.GetResponseStream
        srRead = New System.IO.StreamReader(Str)
        ' read all the text 
        string_text = srRead.ReadToEnd().ToString
        string_text = string_text

        spot_to_find = InStr(string_text, "<td class='featured_ads'>", CompareMethod.Text)

        string_text = Right(string_text, Len(string_text) - spot_to_find - 25)

        array_split = Split(string_text, "'listing_header'")

        For i = 0 To array_split.Length - 2

            pub_reg_no = ""
            pub_ser_no = ""
            pub_desc = ""
            pub_price = ""
            pub_aftt = ""
            pub_seller_info = ""
            pub_picture = ""
            pub_status = ""
            pub_url = ""
            has_pics = False
            aftt_different = ""
            acpub_status = ""
            acpub_process_status = ""
            comp_name = ""
            temp_ac_name = ""
            landings_different = ""
            pub_comp_id = 0
            pub_seller_info = ""
            Response.Flush()
            Response.Flush()

            temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
            If temp_ac_id = 0 Then
                temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
                If temp_ac_id = 0 Then
                    temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                    If temp_ac_id = 0 Then
                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                        If temp_ac_id = 0 Then
                            temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                            If temp_ac_id = 0 Then
                                If Trim(pub_reg_no) <> "" Then
                                    temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                                    If temp_ac_id = 0 Then
                                        ' one last try 
                                        temp_ac_id = find_ac_global_search("", temp_make, temp_model, pub_reg_no)
                                        If temp_ac_id = 0 Then
                                            temp_ac_id = temp_ac_id
                                        End If
                                    End If
                                End If
                            End If

                        End If
                    End If
                End If
            End If

            If temp_ac_id = 0 Then
                temp_ac_id = temp_ac_id
            End If


            acpub_price_details = ""
            acpub_process_status = ""
            If On_Naughty_List(temp_ac_name) = True Or pub_comp_id = 230818 Then
                ' if its on naughtly list then excldue
            Else
                If temp_ac_id > 0 Then
                    Call find_ac_data(temp_ac_id)
                Else
                    acpub_process_status = "For Sale Not Found – No AC Match"
                    acpub_status = "O"
                End If

                If Trim(aftt_different) <> "" Then
                    pub_desc = pub_desc & aftt_different
                End If

                If Trim(landings_different) <> "" Then
                    pub_desc = pub_desc & " " & landings_different
                End If

                If Trim(acpub_price_details) <> "" Then
                    pub_desc = pub_desc & " " & acpub_price_details
                End If

                Call check_insert_ac_pub(temp_ac_id, 26)
            End If

        Next



    End Function
    Function Scrape_For_Business_Air(ByVal link As String) As Long
        Scrape_For_Business_Air = 0

        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader


        '
        System.Net.ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType)
        System.Net.ServicePointManager.Expect100Continue = True

        '   System.Net.ServicePointManager.

        Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(link)


        '   System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls

        ' Dim req As System.Net.WebClient
        '  req.Method = "PUT"
        Dim resp As System.Net.HttpWebResponse = req.GetResponse
        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim i As Integer = 0
        Dim date_string As String = ""
        Dim description_string As String = ""
        Dim title_string As String = ""
        Dim insert_string As String = ""
        Dim link_to_go As String = ""
        Dim comp_name As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_ac_name As String = ""
        Dim temp_make As String = ""
        Dim temp_temp As String = ""
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split2() As String
        Dim temp_manu As String = ""

        Str = resp.GetResponseStream
        srRead = New System.IO.StreamReader(Str)
        ' read all the text 
        string_text = srRead.ReadToEnd().ToString
        string_text = string_text

        spot_to_find = InStr(string_text, "<tr class=""odd views-row-first"">", CompareMethod.Text)

        string_text = Right(string_text, Len(string_text) - spot_to_find - 25)

        array_split = Split(string_text, "'listing_header'")

        For i = 0 To array_split.Length - 2

            pub_reg_no = ""
            pub_ser_no = ""
            pub_desc = ""
            pub_price = ""
            pub_aftt = ""
            pub_seller_info = ""
            pub_picture = ""
            pub_status = ""
            pub_url = ""
            has_pics = False
            aftt_different = ""
            acpub_status = ""
            acpub_process_status = ""
            comp_name = ""
            temp_ac_name = ""
            landings_different = ""
            pub_comp_id = 0
            pub_seller_info = ""
            Response.Flush()
            Response.Flush()




            'spot_to_find = InStr(string_text, "'>", CompareMethod.Text)
            'pub_url = Left(string_text, spot_to_find - 1)
            'pub_url = Replace(pub_url, "href='", "")
            'pub_url = Replace(pub_url, " ", "")
            'pub_url = "https://www.barnstormers.com" & Trim(pub_url)


            'string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
            'spot_to_find = InStr(string_text, "</a>", CompareMethod.Text)
            'temp_ac_name = Left(string_text, spot_to_find - 1)

            'temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
            'If temp_ac_id = 0 Then
            '    temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
            '    If temp_ac_id = 0 Then
            '        temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
            '        If temp_ac_id = 0 Then
            '            temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
            '            If temp_ac_id = 0 Then
            '                temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
            '                If temp_ac_id = 0 Then
            '                    If Trim(pub_reg_no) <> "" Then
            '                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
            '                        If temp_ac_id = 0 Then
            '                            ' one last try 
            '                            temp_ac_id = find_ac_global_search("", temp_make, temp_model, pub_reg_no)
            '                            If temp_ac_id = 0 Then
            '                                temp_ac_id = temp_ac_id
            '                            End If
            '                        End If
            '                    End If
            '                End If

            '            End If
            '        End If
            '    End If
            'End If

            'If temp_ac_id = 0 Then
            '    temp_ac_id = temp_ac_id
            'End If


            'acpub_price_details = ""
            'acpub_process_status = ""
            'If temp_ac_id = 0 And On_Naughty_List(temp_ac_name) = True Then
            '    ' if its on naughtly list then excldue
            'Else
            '    If temp_ac_id > 0 Then
            '        Call find_ac_data(temp_ac_id)
            '    Else
            '        acpub_process_status = "For Sale Not Found – No AC Match"
            '        acpub_status = "O"
            '    End If

            '    If Trim(aftt_different) <> "" Then
            '        pub_desc = pub_desc & aftt_different
            '    End If

            '    If Trim(landings_different) <> "" Then
            '        pub_desc = pub_desc & " " & landings_different
            '    End If

            '    If Trim(acpub_price_details) <> "" Then
            '        pub_desc = pub_desc & " " & acpub_price_details
            '    End If

            '    Call check_insert_ac_pub(temp_ac_id, 26)
            'End If

        Next



    End Function

    Function Scrape_For_global_plane_search(ByVal link As String) As Long
        Scrape_For_global_plane_search = 0

        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader


        '
      '  System.Net.ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType)
        '   System.Net.ServicePointManager.Expect100Continue = True

        '
        System.Net.ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType)
        System.Net.ServicePointManager.Expect100Continue = True

        '   System.Net.ServicePointManager.

        Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(link)
        ' req.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials

        '   System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls

        ' Dim req As System.Net.WebClient
        req.Method = "PUT"
        req.ContentLength = 0
        Dim resp As System.Net.HttpWebResponse = req.GetResponse


        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim i As Integer = 0
        Dim date_string As String = ""
        Dim description_string As String = ""
        Dim title_string As String = ""
        Dim insert_string As String = ""
        Dim link_to_go As String = ""
        Dim comp_name As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_ac_name As String = ""
        Dim temp_make As String = ""
        Dim temp_temp As String = ""
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split2() As String
        Dim temp_manu As String = ""

        Str = resp.GetResponseStream
        srRead = New System.IO.StreamReader(Str)
        ' read all the text 
        string_text = srRead.ReadToEnd().ToString
        string_text = string_text

        spot_to_find = InStr(string_text, "DetailList", CompareMethod.Text)

        string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

        array_split = Split(string_text, "<li class=""standard"">")

        For i = 0 To array_split.Length - 2

            pub_reg_no = ""
            pub_ser_no = ""
            pub_desc = ""
            pub_price = ""
            pub_aftt = ""
            pub_seller_info = ""
            pub_picture = ""
            pub_status = ""
            pub_url = ""
            has_pics = False
            aftt_different = ""
            acpub_status = ""
            acpub_process_status = ""
            comp_name = ""
            temp_ac_name = ""
            landings_different = ""
            pub_comp_id = 0
            pub_seller_info = ""
            Response.Flush()
            Response.Flush()




            'spot_to_find = InStr(string_text, "'>", CompareMethod.Text)
            'pub_url = Left(string_text, spot_to_find - 1)
            'pub_url = Replace(pub_url, "href='", "")
            'pub_url = Replace(pub_url, " ", "")
            'pub_url = "https://www.barnstormers.com" & Trim(pub_url)


            'string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
            'spot_to_find = InStr(string_text, "</a>", CompareMethod.Text)
            'temp_ac_name = Left(string_text, spot_to_find - 1)

            'temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
            'If temp_ac_id = 0 Then
            '    temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
            '    If temp_ac_id = 0 Then
            '        temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
            '        If temp_ac_id = 0 Then
            '            temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
            '            If temp_ac_id = 0 Then
            '                temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
            '                If temp_ac_id = 0 Then
            '                    If Trim(pub_reg_no) <> "" Then
            '                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
            '                        If temp_ac_id = 0 Then
            '                            ' one last try 
            '                            temp_ac_id = find_ac_global_search("", temp_make, temp_model, pub_reg_no)
            '                            If temp_ac_id = 0 Then
            '                                temp_ac_id = temp_ac_id
            '                            End If
            '                        End If
            '                    End If
            '                End If

            '            End If
            '        End If
            '    End If
            'End If

            'If temp_ac_id = 0 Then
            '    temp_ac_id = temp_ac_id
            'End If


            'acpub_price_details = ""
            'acpub_process_status = ""
            'If temp_ac_id = 0 And On_Naughty_List(temp_ac_name) = True Then
            '    ' if its on naughtly list then excldue
            'Else
            '    If temp_ac_id > 0 Then
            '        Call find_ac_data(temp_ac_id)
            '    Else
            '        acpub_process_status = "For Sale Not Found – No AC Match"
            '        acpub_status = "O"
            '    End If

            '    If Trim(aftt_different) <> "" Then
            '        pub_desc = pub_desc & aftt_different
            '    End If

            '    If Trim(landings_different) <> "" Then
            '        pub_desc = pub_desc & " " & landings_different
            '    End If

            '    If Trim(acpub_price_details) <> "" Then
            '        pub_desc = pub_desc & " " & acpub_price_details
            '    End If

            '    Call check_insert_ac_pub(temp_ac_id, 26)
            'End If

        Next



    End Function

    Function Scrape_For_flightmarket(ByVal link As String) As Long
        Scrape_For_flightmarket = 0

        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader


        '
        '   System.Net.ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType)
        '   System.Net.ServicePointManager.Expect100Continue = True

        '   System.Net.ServicePointManager.

        Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(link)


        '   System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls

        ' Dim req As System.Net.WebClient
        '  req.Method = "PUT"
        Dim resp As System.Net.HttpWebResponse = req.GetResponse
        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim i As Integer = 0
        Dim date_string As String = ""
        Dim description_string As String = ""
        Dim title_string As String = ""
        Dim insert_string As String = ""
        Dim link_to_go As String = ""
        Dim comp_name As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_ac_name As String = ""
        Dim temp_make As String = ""
        Dim temp_temp As String = ""
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split2() As String
        Dim temp_manu As String = ""
        Dim Str2 As System.IO.Stream
        Dim srRead2 As System.IO.StreamReader
        Dim req2 As System.Net.WebRequest
        Dim resp2 As System.Net.WebResponse
        Dim string_text3 As String = ""

        Str = resp.GetResponseStream
        srRead = New System.IO.StreamReader(Str)
        ' read all the text 
        string_text = srRead.ReadToEnd().ToString
        string_text = string_text

        spot_to_find = InStr(string_text, "VENDA</h1> - DESTAQUES", CompareMethod.Text)

        string_text = Right(string_text, Len(string_text) - spot_to_find - 20)

        array_split = Split(string_text, "<li>")

        For i = 1 To array_split.Length - 2

            pub_reg_no = ""
            pub_ser_no = ""
            pub_desc = ""
            pub_price = ""
            pub_aftt = ""
            pub_seller_info = ""
            pub_picture = ""
            pub_status = ""
            pub_url = ""
            has_pics = False
            aftt_different = ""
            acpub_status = ""
            acpub_process_status = ""
            comp_name = ""
            temp_ac_name = ""
            landings_different = ""
            pub_comp_id = 0
            pub_seller_info = ""
            Response.Flush()
            Response.Flush()



            string_text = array_split(i)

            spot_to_find = InStr(string_text, "<a href=", CompareMethod.Text)

            string_text = Right(string_text, Len(string_text) - spot_to_find - 8)

            spot_to_find = InStr(string_text, ">", CompareMethod.Text)
            pub_url = "https://www.flightmarket.com.br" & Left(string_text, spot_to_find - 2)


            Try
                pub_reg_no = ""
                pub_ser_no = ""
                pub_aftt = ""
                temp_year = ""
                temp_model = ""
                temp_make = ""
                acpub_original_name = ""

                req2 = System.Net.WebRequest.Create(pub_url)
                resp2 = req2.GetResponse

                Str2 = resp2.GetResponseStream
                srRead2 = New System.IO.StreamReader(Str2)
                string_text3 = srRead2.ReadToEnd().ToString


                spot_to_find = InStr(string_text3, "data-text=""", CompareMethod.Text)
                acpub_original_name = Right(string_text3, Len(string_text3) - spot_to_find - 10)

                spot_to_find = InStr(acpub_original_name, " data-url=", CompareMethod.Text)
                acpub_original_name = Left(acpub_original_name, spot_to_find - 2)


                spot_to_find = InStr(string_text3, "Fabricante</td>", CompareMethod.Text)
                If spot_to_find > 0 Then
                    temp_make = Right(string_text3, Len(string_text3) - spot_to_find - 15)

                    spot_to_find = InStr(temp_make, "</td>", CompareMethod.Text)
                    temp_make = Left(temp_make, spot_to_find - 1)
                    temp_make = replace_fm(temp_make)
                End If


                spot_to_find = InStr(string_text3, "Modelos</td>", CompareMethod.Text)
                If spot_to_find > 0 Then
                    temp_model = Right(string_text3, Len(string_text3) - spot_to_find - 12)

                    spot_to_find = InStr(temp_model, "</td>", CompareMethod.Text)
                    temp_model = Left(temp_model, spot_to_find - 1)
                    temp_model = replace_fm(temp_model)
                End If

                spot_to_find = InStr(string_text3, "Ano</td>", CompareMethod.Text)
                If spot_to_find > 0 Then
                    temp_year = Right(string_text3, Len(string_text3) - spot_to_find - 8)

                    spot_to_find = InStr(temp_year, "</td>", CompareMethod.Text)
                    temp_year = Left(temp_year, spot_to_find - 1)
                    temp_year = replace_fm(temp_year)
                End If


                spot_to_find = InStr(string_text3, "Horas totais</td>", CompareMethod.Text)
                If spot_to_find > 0 Then
                    pub_aftt = Right(string_text3, Len(string_text3) - spot_to_find - 17)

                    spot_to_find = InStr(pub_aftt, "</td>", CompareMethod.Text)
                    pub_aftt = Left(pub_aftt, spot_to_find - 1)
                    pub_aftt = replace_fm(pub_aftt)
                End If

                spot_to_find = InStr(string_text3, "de Série</td>", CompareMethod.Text)
                If spot_to_find > 0 Then
                    pub_ser_no = Right(string_text3, Len(string_text3) - spot_to_find - 12)

                    spot_to_find = InStr(pub_ser_no, "</td>", CompareMethod.Text)
                    pub_ser_no = Left(pub_ser_no, spot_to_find - 1)
                    pub_ser_no = replace_fm(pub_ser_no)
                    pub_ser_no = replace_fm(pub_ser_no)
                End If

                spot_to_find = InStr(string_text3, "Matrícula</td>", CompareMethod.Text)
                If spot_to_find > 0 Then
                    pub_reg_no = Right(string_text3, Len(string_text3) - spot_to_find - 14)

                    spot_to_find = InStr(pub_reg_no, "</td>", CompareMethod.Text)
                    pub_reg_no = Left(pub_reg_no, spot_to_find - 1)
                    pub_reg_no = replace_fm(pub_reg_no)
                End If




                'pub_url = Replace(pub_url, "href='", "")
                'pub_url = Replace(pub_url, " ", "")
                'pub_url = "https://www.barnstormers.com" & Trim(pub_url)


                'string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
                'spot_to_find = InStr(string_text, "</a>", CompareMethod.Text)
                'temp_ac_name = Left(string_text, spot_to_find - 1)

                temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
                If temp_ac_id = 0 Then
                    temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
                    If temp_ac_id = 0 Then
                        temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                        If temp_ac_id = 0 Then
                            temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                            If temp_ac_id = 0 Then
                                temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                                If temp_ac_id = 0 Then
                                    If Trim(pub_reg_no) <> "" Then
                                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                                        If temp_ac_id = 0 Then
                                            ' one last try 
                                            temp_ac_id = find_ac_global_search("", temp_make, temp_model, pub_reg_no)
                                            If temp_ac_id = 0 Then
                                                temp_ac_id = temp_ac_id
                                            End If
                                        End If
                                    End If
                                End If

                            End If
                        End If
                    End If
                End If

                If temp_ac_id = 0 Then
                    temp_ac_id = temp_ac_id
                End If


                acpub_price_details = ""
                acpub_process_status = ""
                If On_Naughty_List(temp_ac_name) = True Then
                    ' if its on naughtly list then excldue
                Else
                    If temp_ac_id > 0 Then
                        Call find_ac_data(temp_ac_id)
                    Else
                        acpub_process_status = "For Sale Not Found – No AC Match"
                        acpub_status = "O"
                    End If

                    If Trim(aftt_different) <> "" Then
                        pub_desc = pub_desc & aftt_different
                    End If

                    If Trim(landings_different) <> "" Then
                        pub_desc = pub_desc & " " & landings_different
                    End If

                    If Trim(acpub_price_details) <> "" Then
                        pub_desc = pub_desc & " " & acpub_price_details
                    End If

                    Call check_insert_ac_pub(temp_ac_id, 28)
                End If

            Catch ex As Exception

            End Try

        Next



    End Function
    Public Function replace_fm(ByVal temp_String As String)

        replace_fm = ""
        replace_fm = temp_String
        replace_fm = Replace(replace_fm, "<td>", "")
        replace_fm = Replace(replace_fm, "vbLf & """, "")
        replace_fm = Replace(replace_fm, vbLf, "")
        replace_fm = Replace(replace_fm, vbCr, "")
        replace_fm = Replace(replace_fm, vbCrLf, "")
        replace_fm = Replace(replace_fm, """", "")
        replace_fm = LTrim(RTrim(replace_fm))
        cutme(replace_fm)   'gets rid of starting and ending spaces 

    End Function

    Function Scrape_For_aviapages(ByVal link As String) As Long
        Scrape_For_aviapages = 0

        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader


        '
        '   System.Net.ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType)
        '   System.Net.ServicePointManager.Expect100Continue = True

        '   System.Net.ServicePointManager.

        Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(link)


        '   System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls

        ' Dim req As System.Net.WebClient
        '  req.Method = "PUT"
        Dim resp As System.Net.HttpWebResponse = req.GetResponse
        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim i As Integer = 0
        Dim date_string As String = ""
        Dim description_string As String = ""
        Dim title_string As String = ""
        Dim insert_string As String = ""
        Dim link_to_go As String = ""
        Dim comp_name As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_ac_name As String = ""
        Dim temp_make As String = ""
        Dim temp_temp As String = ""
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split2() As String
        Dim temp_manu As String = ""

        Str = resp.GetResponseStream
        srRead = New System.IO.StreamReader(Str)
        ' read all the text 
        string_text = srRead.ReadToEnd().ToString
        string_text = string_text

        spot_to_find = InStr(string_text, "vt-fcn", CompareMethod.Text)

        string_text = Right(string_text, Len(string_text) - spot_to_find - 25)

        array_split = Split(string_text, "'listing_header'")

        For i = 0 To array_split.Length - 2

            pub_reg_no = ""
            pub_ser_no = ""
            pub_desc = ""
            pub_price = ""
            pub_aftt = ""
            pub_seller_info = ""
            pub_picture = ""
            pub_status = ""
            pub_url = ""
            has_pics = False
            aftt_different = ""
            acpub_status = ""
            acpub_process_status = ""
            comp_name = ""
            temp_ac_name = ""
            landings_different = ""
            pub_comp_id = 0
            pub_seller_info = ""
            Response.Flush()
            Response.Flush()




            'spot_to_find = InStr(string_text, "'>", CompareMethod.Text)
            'pub_url = Left(string_text, spot_to_find - 1)
            'pub_url = Replace(pub_url, "href='", "")
            'pub_url = Replace(pub_url, " ", "")
            'pub_url = "https://www.barnstormers.com" & Trim(pub_url)


            'string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
            'spot_to_find = InStr(string_text, "</a>", CompareMethod.Text)
            'temp_ac_name = Left(string_text, spot_to_find - 1)

            'temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
            'If temp_ac_id = 0 Then
            '    temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
            '    If temp_ac_id = 0 Then
            '        temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
            '        If temp_ac_id = 0 Then
            '            temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
            '            If temp_ac_id = 0 Then
            '                temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
            '                If temp_ac_id = 0 Then
            '                    If Trim(pub_reg_no) <> "" Then
            '                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
            '                        If temp_ac_id = 0 Then
            '                            ' one last try 
            '                            temp_ac_id = find_ac_global_search("", temp_make, temp_model, pub_reg_no)
            '                            If temp_ac_id = 0 Then
            '                                temp_ac_id = temp_ac_id
            '                            End If
            '                        End If
            '                    End If
            '                End If

            '            End If
            '        End If
            '    End If
            'End If

            'If temp_ac_id = 0 Then
            '    temp_ac_id = temp_ac_id
            'End If


            'acpub_price_details = ""
            'acpub_process_status = ""
            'If temp_ac_id = 0 And On_Naughty_List(temp_ac_name) = True Then
            '    ' if its on naughtly list then excldue
            'Else
            '    If temp_ac_id > 0 Then
            '        Call find_ac_data(temp_ac_id)
            '    Else
            '        acpub_process_status = "For Sale Not Found – No AC Match"
            '        acpub_status = "O"
            '    End If

            '    If Trim(aftt_different) <> "" Then
            '        pub_desc = pub_desc & aftt_different
            '    End If

            '    If Trim(landings_different) <> "" Then
            '        pub_desc = pub_desc & " " & landings_different
            '    End If

            '    If Trim(acpub_price_details) <> "" Then
            '        pub_desc = pub_desc & " " & acpub_price_details
            '    End If

            '    Call check_insert_ac_pub(temp_ac_id, 26)
            'End If

        Next



    End Function
    Function Scrape_Barnstormers(ByVal link As String) As Long
        Scrape_Barnstormers = 0

        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader
        System.Net.ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType)
        'The underlying connection was closed: An unexpected error occurred on a send

        Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(link)
        ' Dim req As System.Net.WebClient
        '  req.Method = "PUT"
        Dim resp As System.Net.HttpWebResponse = req.GetResponse
        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim i As Integer = 0
        Dim date_string As String = ""
        Dim description_string As String = ""
        Dim title_string As String = ""
        Dim insert_string As String = ""
        Dim link_to_go As String = ""
        Dim comp_name As String = ""
        Dim temp_ac_id As Long = 0
        Dim temp_ac_name As String = ""
        Dim temp_make As String = ""
        Dim temp_temp As String = ""
        Dim temp_model As String = ""
        Dim temp_year As String = ""
        Dim array_split2() As String
        Dim temp_manu As String = ""

        Str = resp.GetResponseStream
        srRead = New System.IO.StreamReader(Str)
        ' read all the text 
        string_text = srRead.ReadToEnd().ToString
        string_text = string_text

        Try


            spot_to_find = InStr(string_text, "<td class='featured_ads'>", CompareMethod.Text)
            If spot_to_find > 0 Then

                string_text = Right(string_text, Len(string_text) - spot_to_find - 25)


                spot_to_find = InStr(string_text, "</ul>", CompareMethod.Text)
                If spot_to_find > 0 Then
                    string_text = Left(string_text, spot_to_find - 1)
                End If


                array_split = Split(string_text, "'listing_header'")

                For i = 1 To array_split.Length - 1
                    string_text = array_split(i)
                    pub_reg_no = ""
                    pub_ser_no = ""
                    pub_desc = ""
                    pub_price = ""
                    pub_aftt = ""
                    pub_seller_info = ""
                    pub_picture = ""
                    pub_status = ""
                    pub_url = ""
                    has_pics = False
                    aftt_different = ""
                    acpub_status = ""
                    acpub_process_status = ""
                    comp_name = ""
                    temp_ac_name = ""
                    landings_different = ""
                    temp_ac_id = 0
                    pub_comp_id = 0
                    pub_seller_info = ""
                    Response.Flush()
                    Response.Flush()




                    spot_to_find = InStr(string_text, "'>", CompareMethod.Text)
                    pub_url = Left(string_text, spot_to_find - 1)
                    pub_url = Replace(pub_url, "href='", "")
                    pub_url = Replace(pub_url, " ", "")
                    pub_url = "https://www.barnstormers.com" & Trim(pub_url)


                    string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
                    spot_to_find = InStr(string_text, "</a>", CompareMethod.Text)
                    temp_ac_name = Left(string_text, spot_to_find - 1)

                    temp_ac_name = Replace(temp_ac_name, "&#039;", "")
                    temp_ac_name = Replace(temp_ac_name, "&amp;", "")
                    temp_ac_name = Replace(temp_ac_name, "&#039;", "")
                    temp_ac_name = Replace(temp_ac_name, "&#039;", "")
                    temp_ac_name = Replace(temp_ac_name, "&#039;", "")
                    temp_ac_name = Replace(temp_ac_name, "'", "")


                    acpub_original_name = temp_ac_name

                    'temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
                    'If temp_ac_id = 0 Then
                    '    temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
                    '    If temp_ac_id = 0 Then
                    '        temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
                    '        If temp_ac_id = 0 Then
                    '            temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                    '            If temp_ac_id = 0 Then
                    '                temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                    '                If temp_ac_id = 0 Then
                    '                    If Trim(pub_reg_no) <> "" Then
                    '                        temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                    '                        If temp_ac_id = 0 Then
                    '                            ' one last try 
                    '                            temp_ac_id = find_ac_global_search("", temp_make, temp_model, pub_reg_no)
                    '                            If temp_ac_id = 0 Then
                    '                                temp_ac_id = temp_ac_id
                    '                            End If
                    '                        End If
                    '                    End If
                    '                End If

                    '            End If
                    '        End If
                    '    End If
                    'End If

                    'If temp_ac_id = 0 Then
                    '    temp_ac_id = temp_ac_id
                    'End If


                    'acpub_price_details = ""
                    'acpub_process_status = ""
                    If On_Naughty_List(temp_ac_name) = True Then
                        ' if its on naughtly list then excldue
                    Else
                        If temp_ac_id > 0 Then
                            Call find_ac_data(temp_ac_id)
                        Else
                            acpub_process_status = "For Sale Not Found – No AC Match"
                            acpub_status = "O"
                        End If

                        If Trim(aftt_different) <> "" Then
                            pub_desc = pub_desc & aftt_different
                        End If

                        If Trim(landings_different) <> "" Then
                            pub_desc = pub_desc & " " & landings_different
                        End If

                        If Trim(acpub_price_details) <> "" Then
                            pub_desc = pub_desc & " " & acpub_price_details
                        End If

                        Call check_insert_ac_pub(temp_ac_id, 27)
                    End If

                Next
            Else
                pub_desc = pub_desc
            End If


        Catch ex As Exception

        End Try

    End Function

    Function Scrape_Aircraft_Exchange_Detail(ByVal link As String) As Long
    Scrape_Aircraft_Exchange_Detail = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(link)
    ' Dim req As System.Net.WebClient
    '  req.Method = "PUT"
    Dim resp As System.Net.HttpWebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Dim comp_name As String = ""
    Dim temp_ac_id As Long = 0
    Dim temp_ac_name As String = ""
    Dim temp_make As String = ""
    Dim temp_temp As String = ""
    Dim temp_model As String = ""
    Dim temp_year As String = ""
    Dim array_split2() As String
    Dim temp_manu As String = ""
 
      Try


       
 


        Str = resp.GetResponseStream
        srRead = New System.IO.StreamReader(Str)
        ' read all the text 
        string_text = srRead.ReadToEnd().ToString
        string_text = string_text

        spot_to_find = InStr(string_text, "<div class=""aircraft"">", CompareMethod.Text)

        string_text = Right(string_text, Len(string_text) - spot_to_find - 22)

        array_split = Split(string_text, "<div class=""aircraft"">")

      For i = 0 To array_split.Length - 2
        string_text = array_split(i)

        '- to be moved inside --
        acpub_count = acpub_count + 1

        Response.Flush()
        Response.Flush()
        pub_reg_no = ""
        pub_ser_no = ""
        pub_desc = ""
        pub_price = ""
        pub_aftt = ""
        pub_seller_info = ""
        pub_picture = ""
        pub_status = ""
        pub_url = ""
        has_pics = False
        aftt_different = ""
        acpub_status = ""
        acpub_process_status = ""
        comp_name = ""
        temp_ac_name = ""
                landings_different = ""
                pub_comp_id = 0
                pub_seller_info = ""
                Response.Flush()
        Response.Flush()

        pub_comp_id = 0
        If InStr(LCase(string_text), LCase("Sojourn Aviation Company")) > 0 Then
          pub_comp_id = 290871
        ElseIf InStr(LCase(string_text), LCase("Exclusive Aircraft Sales")) > 0 Then
          pub_comp_id = 221839
        ElseIf InStr(LCase(string_text), LCase("Meisinger Aviation")) > 0 Then
          pub_comp_id = 287118
        ElseIf InStr(LCase(string_text), LCase("QS Partners")) > 0 Then
          pub_comp_id = 378059
        ElseIf InStr(LCase(string_text), LCase("Aerolineas Ejecutivas")) > 0 Then
          pub_comp_id = 173754
        ElseIf InStr(LCase(string_text), LCase("Jack Prewitt & Associates")) > 0 Then
          pub_comp_id = 10843
        ElseIf InStr(LCase(string_text), LCase("Asian Sky Group")) > 0 Then
          pub_comp_id = 310707
        ElseIf InStr(LCase(string_text), LCase("ACASS Canada")) > 0 Then
          pub_comp_id = 151069
        ElseIf InStr(LCase(string_text), LCase("AVPRO")) > 0 Then
          pub_comp_id = 12142
        ElseIf InStr(LCase(string_text), LCase("Axiom Aviation")) > 0 Then
          pub_comp_id = 141512
        ElseIf InStr(LCase(string_text), LCase("Axis Aviation")) > 0 Then
          pub_comp_id = 317611
        ElseIf InStr(LCase(string_text), LCase("Banyan Air Service")) > 0 Then
          pub_comp_id = 133650
        ElseIf InStr(LCase(string_text), LCase("Business Aircraft Leasing")) > 0 Then
          pub_comp_id = 1801
        ElseIf InStr(LCase(string_text), LCase("Corporate Fleet Service")) > 0 Then
          pub_comp_id = 16663
        ElseIf InStr(LCase(string_text), LCase("Cutter Aviation")) > 0 Then
          pub_comp_id = 1581
        ElseIf InStr(LCase(string_text), LCase("Dallas Jet Internationa")) > 0 Then
          pub_comp_id = 141851
        ElseIf InStr(LCase(string_text), LCase("Dassault Falcon Jet")) > 0 Then
          pub_comp_id = 14082
        ElseIf InStr(LCase(string_text), LCase("Duncan Aviation")) > 0 Then
          pub_comp_id = 7223
        ElseIf InStr(LCase(string_text), LCase("Eagle Aviation")) > 0 Then
          pub_comp_id = 2473
        ElseIf InStr(LCase(string_text), LCase("Eagle Creek")) > 0 Then
          pub_comp_id = 23849
        ElseIf InStr(LCase(string_text), LCase("Elliott Jets")) > 0 Then
          pub_comp_id = 1770
        ElseIf InStr(LCase(string_text), LCase("Embraer Executive")) > 0 Then
          pub_comp_id = 311813
        ElseIf InStr(LCase(string_text), LCase("Gantt Aviation")) > 0 Then
          pub_comp_id = 17884
        ElseIf InStr(LCase(string_text), LCase("General Aviation Services")) > 0 Then
          pub_comp_id = 16520
        ElseIf InStr(LCase(string_text), LCase("Global Flight")) > 0 Then
          pub_comp_id = 15104
        ElseIf InStr(LCase(string_text), LCase("Global Wings")) > 0 Then
          pub_comp_id = 102406
        ElseIf InStr(LCase(string_text), LCase("Guardian Jet")) > 0 Then
          pub_comp_id = 136677
        ElseIf InStr(LCase(string_text), LCase("Gulfstream Aerospace Corporation")) > 0 Then
          pub_comp_id = 9487
        ElseIf InStr(LCase(string_text), LCase("Hatt &amp; Associates")) > 0 Then
          pub_comp_id = 366875
        ElseIf InStr(LCase(string_text), LCase("International Jet Trade")) > 0 Then
          pub_comp_id = 1862
        ElseIf InStr(LCase(string_text), LCase("JBA Aviation")) > 0 Then
          pub_comp_id = 26353
        ElseIf InStr(LCase(string_text), LCase("Jet Sense")) > 0 Then
          pub_comp_id = 338030
        ElseIf InStr(LCase(string_text), LCase("Jetcraft")) > 0 Then
          pub_comp_id = 9400
        ElseIf InStr(LCase(string_text), LCase("Jeteffect")) > 0 Then
          pub_comp_id = 26512
        ElseIf InStr(LCase(string_text), LCase("Leading Edge Aviation Solutio")) > 0 Then
          pub_comp_id = 11648
        ElseIf InStr(LCase(string_text), LCase("Meisner Aircraft")) > 0 Then
          pub_comp_id = 26377
        ElseIf InStr(LCase(string_text), LCase("MENTE Group")) > 0 Then
          pub_comp_id = 279825
        ElseIf InStr(LCase(string_text), LCase("OGARAJET")) > 0 Then
          pub_comp_id = 10574
        ElseIf InStr(LCase(string_text), LCase("Southern Cross Aircraf")) > 0 Then
          pub_comp_id = 7268
        ElseIf InStr(LCase(string_text), LCase("Textron Aviation")) > 0 Then
          pub_comp_id = 1122
        ElseIf InStr(LCase(string_text), LCase("Western Aircraft")) > 0 Then
          pub_comp_id = 20848
        ElseIf InStr(LCase(string_text), LCase("Wetzel Aviation")) > 0 Then
          pub_comp_id = 152650
        ElseIf InStr(LCase(string_text), LCase("Skyservice Business Aviati")) > 0 Then
          pub_comp_id = 221839
        ElseIf InStr(LCase(string_text), LCase("SOLJETS")) > 0 Then
          pub_comp_id = 377610
        Else
          pub_comp_id = 0
        End If





        spot_to_find = InStr(string_text, "<a class=""no-underline"" href=""", CompareMethod.Text)
        string_text = Right(string_text, Len(string_text) - spot_to_find - 29)

        spot_to_find = InStr(string_text, ">", CompareMethod.Text)
        pub_url = Left(string_text, spot_to_find - 2)


        spot_to_find = InStr(string_text, "<h4 class=""font-semibold"">", CompareMethod.Text)
        string_text = Right(string_text, Len(string_text) - spot_to_find - 25)

        spot_to_find = InStr(string_text, "</h4>", CompareMethod.Text)
        temp_ac_name = Left(string_text, spot_to_find - 1)



        Try

          temp_year = Left(Trim(temp_ac_name), 4)
          ' temp_ac_name = Replace(temp_ac_name, temp_year & " ", "")

          array_split2 = Split(Replace(temp_ac_name, temp_year & " ", ""), " ")

          If array_split2.Length > 0 Then
            If array_split2.Length > 5 Then
              temp_temp = temp_temp
            ElseIf array_split2.Length = 5 Then

              If UCase(Trim(array_split2(3))) = "EASY" Then
                temp_manu = array_split2(0)
                temp_make = array_split2(1)
                temp_model = array_split2(2)
                temp_model = temp_model & " " & array_split2(3)
                temp_model = temp_model & " " & array_split2(4)
              Else
                temp_manu = array_split2(0)
                temp_make = array_split2(1)
                temp_model = array_split2(2)
                temp_model = temp_model & " " & array_split2(3)
                temp_model = temp_model & " " & array_split2(4)
              End If


            ElseIf array_split2.Length = 4 Then

              If UCase(Trim(array_split2(3))) = "EASY" Then
                temp_manu = array_split2(0)
                temp_make = array_split2(1)
                temp_model = array_split2(2)
                temp_model = temp_model & " " & array_split2(3)
              ElseIf UCase(Trim(array_split2(1))) = "CITATION" And UCase(Trim(array_split2(2))) = "JET" Then
                temp_manu = array_split2(0)
                temp_make = array_split2(1)
                temp_model = array_split2(3)
              ElseIf UCase(Trim(array_split2(1))) = "GRAND" And UCase(Trim(array_split2(2))) = "CARAVAN" Then
                temp_manu = array_split2(0)
                temp_make = array_split2(1)
                temp_make = temp_make & " " & array_split2(2)
                temp_model = array_split2(3)
              ElseIf UCase(Trim(array_split2(0))) = "ASTRA/GULFSTREAM" And UCase(Trim(array_split2(1))) = "1125" Then
                temp_manu = "ASTRA"
                temp_make = "ASTRA"
                temp_model = "1125"
                temp_model = temp_model & " " & array_split2(3)
              Else
                temp_manu = array_split2(0)
                temp_make = array_split2(1)
                temp_model = array_split2(2)
                temp_model = temp_model & " " & array_split2(3)
              End If
         


            ElseIf array_split2.Length = 3 Then

              temp_manu = array_split2(0)
              temp_make = array_split2(1)
              temp_model = array_split2(2)

            ElseIf array_split2.Length = 2 Then

              temp_make = array_split2(0)
              temp_model = array_split2(1)

            ElseIf array_split2.Length = 1 Then
              temp_temp = temp_temp
            End If
          End If

        Catch ex As Exception

        End Try




        spot_to_find = InStr(string_text, "Price:</span>", CompareMethod.Text)
        string_text = Right(string_text, Len(string_text) - spot_to_find - 13)

        spot_to_find = InStr(string_text, "</p>", CompareMethod.Text)
        pub_price = Left(string_text, spot_to_find - 1)
        pub_price = Replace(Replace(pub_price, "$", ""), ",", "")
        cutme(pub_price)
        pub_price = RTrim(LTrim(pub_price))

        If Trim(pub_price) = "Make an Offer" Then
          pub_price = "Make Offer"
        End If

        Response.Flush()
        Response.Flush()
        Response.Flush()

        spot_to_find = InStr(string_text, "Serial Number:</span>", CompareMethod.Text)
        If spot_to_find > 0 Then
          string_text = Right(string_text, Len(string_text) - spot_to_find - 21)

          spot_to_find = InStr(string_text, "</li>", CompareMethod.Text)
          pub_ser_no = Left(string_text, spot_to_find - 1)
          cutme(pub_ser_no)
          pub_ser_no = RTrim(LTrim(pub_ser_no))
        End If


        spot_to_find = InStr(string_text, "Tail Number:</span>", CompareMethod.Text)
        If spot_to_find > 0 Then
          string_text = Right(string_text, Len(string_text) - spot_to_find - 19)

          spot_to_find = InStr(string_text, "</li>", CompareMethod.Text)
          pub_reg_no = Left(string_text, spot_to_find - 1)
          cutme(pub_reg_no)
          pub_reg_no = RTrim(LTrim(pub_reg_no))
        End If


                spot_to_find = InStr(string_text, "Hours:</span>", CompareMethod.Text)
                If spot_to_find > 0 Then
                    string_text = Right(string_text, Len(string_text) - spot_to_find - 13)

                    spot_to_find = InStr(string_text, "</li>", CompareMethod.Text)
                    pub_aftt = Left(string_text, spot_to_find - 1)
                    cutme(pub_aftt)
                    pub_aftt = Replace(pub_aftt, ",", "")



                    spot_to_find = InStr(pub_aftt, "(", CompareMethod.Text)
                    If spot_to_find > 0 Then
                        pub_aftt = Left(string_text, spot_to_find - 1)   ' found 1 example
                    End If

                    pub_aftt = Replace(pub_aftt, "DeliveryTimeOnly", "")
                    pub_aftt = RTrim(LTrim(pub_aftt))


                End If

                spot_to_find = InStr(string_text, "Cycles:</span>", CompareMethod.Text)
        If spot_to_find > 0 Then
          string_text = Right(string_text, Len(string_text) - spot_to_find - 14)

          spot_to_find = InStr(string_text, "</li>", CompareMethod.Text)
          pub_landings = Left(string_text, spot_to_find - 1)
          cutme(pub_landings)
          pub_landings = Replace(pub_landings, ",", "")
          pub_landings = RTrim(LTrim(pub_landings))
        End If

        acpub_original_name = temp_ac_name & " " & pub_ser_no

        pub_desc = ""
        Response.Flush()
        System.Threading.Thread.Sleep(10) 
        Response.Flush()
        Response.Flush()
        Response.Flush()
        Response.Flush()
        Response.Flush()

        '  If InStr(acpub_original_name, "2002 Bombardier Learjet 45") > 0 Then
        ' acpub_original_name = acpub_original_name
        '  End If


        temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
        If temp_ac_id = 0 Then
          temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, "", "")
          If temp_ac_id = 0 Then
            temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
            If temp_ac_id = 0 Then
              temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
              If temp_ac_id = 0 Then
                temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                If temp_ac_id = 0 Then
                  If Trim(pub_reg_no) <> "" Then
                    temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                    If temp_ac_id = 0 Then
                      ' one last try 
                      temp_ac_id = find_ac_global_search("", temp_make, temp_model, pub_reg_no)
                      If temp_ac_id = 0 Then
                        temp_ac_id = temp_ac_id
                      End If
                    End If
                  End If
                End If

              End If
            End If
          End If
        End If

        If temp_ac_id = 0 Then
          temp_ac_id = temp_ac_id
        End If

        If Trim(temp_ac_name) = "1996 Hawker 800XP" Then
          temp_ac_name = temp_ac_name
        End If


        acpub_price_details = ""
        acpub_process_status = ""
                If On_Naughty_List(temp_ac_name) = True Then
                    ' if its on naughtly list then excldue
                Else
                    If temp_ac_id > 0 Then
            Call find_ac_data(temp_ac_id)
          Else
            acpub_process_status = "For Sale Not Found – No AC Match"
            acpub_status = "O"
          End If

          If Trim(aftt_different) <> "" Then
            pub_desc = pub_desc & aftt_different
          End If

          If Trim(landings_different) <> "" Then
            pub_desc = pub_desc & " " & landings_different
          End If

          If Trim(acpub_price_details) <> "" Then
            pub_desc = pub_desc & " " & acpub_price_details
          End If

          Call check_insert_ac_pub(temp_ac_id, 26)
        End If




        '  spot_to_find = InStr(string_text, "ArticleStartDateOWSDATE"":""", CompareMethod.Text)
        '  string_text = Right(string_text, Len(string_text) - spot_to_find - 25)




      Next






      Catch ex As Exception
        'Response.Write(ex)
      Finally


      End Try
 
  End Function
  Function Scrape_Aircraft_Exchange(ByVal link As String, ByVal id As Integer) As Long
    Scrape_Aircraft_Exchange = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(link)
    ' Dim req As System.Net.WebClient
    '  req.Method = "PUT"
    Dim resp As System.Net.HttpWebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Dim comp_name As String = ""

    Try

      Try

        MySqlConn_JETNET.ConnectionString = Inhouse_Live_Connection
        'MySqlConn_JETNET.ConnectionString = Inhouse_Test_Connection
        ' MySqlConn_JETNET.ConnectionString = JETNET_LIVE_SQL_CONN
        MySqlConn_JETNET.Open()
        MySqlCommand_JETNET.Connection = MySqlConn_JETNET
        MySqlCommand_JETNET.CommandType = CommandType.Text
        MySqlCommand_JETNET.CommandTimeout = 60

        Call insert_into_eventlog("Aircraft Pubs Started", "Research Assistant")

        ypl_start_date = Date.Now

        Call Find_Naughty_Models()

        ' Dim webClient As New System.Net.WebClient
        ' Dim result As String = webClient.DownloadString(link)

        'Try
        '  Dim fr As System.Net.HttpWebRequest
        '  Dim targetURI As New Uri(link)

        '  fr = DirectCast(System.Net.HttpWebRequest.Create(targetURI), System.Net.HttpWebRequest)
        '  If (fr.GetResponse().ContentLength > 0) Then
        '    Dim str As New System.IO.StreamReader(fr.GetResponse().GetResponseStream())
        '    Response.Write(str.ReadToEnd())
        '  str.Close(); 
        '  End If
        'Catch ex As System.Net.WebException
        '  'Error in accessing the resource, handle it
        'End Try


        '  Dim result As String = req.DownloadString(link)


        'Using client As New Net.WebClient
        '  Dim reqparm As New Specialized.NameValueCollection
        '  ' reqparm.Add("page", "1")
        '  Dim responsebytes = client.UploadValues(link, "POST", reqparm)
        'End Using


        Str = resp.GetResponseStream
        srRead = New System.IO.StreamReader(Str)
        ' read all the text 
        string_text = srRead.ReadToEnd().ToString
        string_text = string_text


        spot_to_find = InStr(string_text, "<div class=""broker-box leading-normal text-sm flex flex-col justify-between"">", CompareMethod.Text)

        string_text = Right(string_text, Len(string_text) - spot_to_find - 77)

        array_split = Split(string_text, "<div class=""broker-box leading-normal text-sm flex flex-col justify-between"">")

        For i = 0 To array_split.Length - 1
          string_text = array_split(i)

          '- to be moved inside --
          acpub_count = acpub_count + 1

          pub_reg_no = ""
          pub_ser_no = ""
          pub_desc = ""
          pub_price = ""
          pub_aftt = ""
          pub_seller_info = ""
          pub_picture = ""
          pub_status = ""
          pub_url = ""
          has_pics = False
          aftt_different = ""
          acpub_status = ""
          acpub_process_status = ""
          comp_name = ""



          spot_to_find = InStr(string_text, "<div class=""logo"">", CompareMethod.Text)
          string_text = Right(string_text, Len(string_text) - spot_to_find - 18)


          If InStr(string_text, "ASIAN SKY GROUP") > 0 Then
            string_text = string_text 
 
            spot_to_find = InStr(string_text, "<a href=""", CompareMethod.Text)
            string_text = Right(string_text, Len(string_text) - spot_to_find - 8)


            spot_to_find = InStr(string_text, ">", CompareMethod.Text)
            link_to_go = Left(string_text, spot_to_find - 2)


            spot_to_find = InStr(string_text, "<h4 class=""leading-tight uppercase"">", CompareMethod.Text)
            string_text = Right(string_text, Len(string_text) - spot_to_find - 36)

            spot_to_find = InStr(string_text, "</h4>", CompareMethod.Text)
            pub_seller_info = Left(string_text, spot_to_find - 1)

            Call Scrape_Aircraft_Exchange_Detail(link_to_go)

          End If
          '  spot_to_find = InStr(string_text, "ArticleStartDateOWSDATE"":""", CompareMethod.Text)
          '  string_text = Right(string_text, Len(string_text) - spot_to_find - 25)




        Next



        '  Call Insert_EMail_Queue_Record(yt_table)

        '  Call insert_into_eventlog("Aircraft Pubs Finished", "Research Assistant")

      Catch ex As Exception
        'Response.Write(ex)
      Finally

        MySqlConn_JETNET.Dispose()
        MySqlConn_JETNET.Close()
        MySqlConn_JETNET = Nothing

      End Try

    Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try
  End Function


  Function ABI_News_Scraper_Dassult_Falcon(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_Dassult_Falcon = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text

      'spot_to_find = InStr(string_text, "<ul class=" & Chr("34") & "art-list" & Chr("34") & ">", CompareMethod.Text)

      'string_text = Right(string_text, Len(string_text) - spot_to_find)

      'spot_to_find = InStr(string_text, "</ul>", CompareMethod.Text)

      'string_text = Left(string_text, spot_to_find)

      'array_split = Split(string_text, "<li>")



      spot_to_find = InStr(string_text, """Rank"":0", CompareMethod.Text)

      string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

      array_split = Split(string_text, """Rank"":0")


      For i = 0 To array_split.Length - 1
        string_text = array_split(i)


        spot_to_find = InStr(string_text, """Path"":", CompareMethod.Text)
        string_text = Right(string_text, Len(string_text) - spot_to_find - 7)


        spot_to_find = InStr(string_text, "Description", CompareMethod.Text)
        link_to_go = Left(string_text, spot_to_find - 3)
        link_to_go = Replace(link_to_go, "\", "/")
        link_to_go = Replace(link_to_go, "u002f", "")
        link_to_go = Replace(link_to_go, """", "")


        spot_to_find = InStr(string_text, "ArticleStartDateOWSDATE"":""", CompareMethod.Text)
        string_text = Right(string_text, Len(string_text) - spot_to_find - 25)

        spot_to_find = InStr(string_text, "ListShortSummaryOWSTEXT"":", CompareMethod.Text)
        date_string = Left(string_text, spot_to_find - 3)

        spot_to_find = InStr(date_string, "T", CompareMethod.Text)
        date_string = Left(date_string, spot_to_find - 1)

        date_string = change_date_format(date_string, 1, 2, 0)


        string_text = Right(string_text, Len(string_text) - spot_to_find - 25)


        spot_to_find = InStr(string_text, "ArticleTitleOWSTEXT"":""", CompareMethod.Text)

        description_string = Left(string_text, spot_to_find - 3)
        description_string = Replace(description_string, """", "")
        description_string = Replace(description_string, "aryOWSTEXT", ":")

        string_text = Right(string_text, Len(string_text) - spot_to_find - 21)

        spot_to_find = InStr(string_text, """", CompareMethod.Text)

        title_string = Left(string_text, spot_to_find - 1)
        title_string = Replace(title_string, """", "")


        If record_exists(date_string, title_string, link_to_go) = False Then


          link_to_go = Replace(link_to_go, "&#39;", "'")
          description_string = replace_all_chars(description_string)
          description_string = Left(description_string, 500)


          link_to_go = Replace(link_to_go, "'", "&#39;")




          insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
          insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
          insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
          insert_string = insert_string & " ) VALUES ( "
          insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & Trim(description_string) & "','"
          insert_string = insert_string & link_to_go & "','" & id & "','FALCON','',"
          insert_string = insert_string & "'',0"
          insert_string = insert_string & " ) "

          ' Response.Write(insert_string & "<br><br>")




          'SETUP AND EXECUTE THE SQL INSERT COMMAND
          Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

          sqlComm.ExecuteNonQuery()
          sqlComm.Dispose()

          ABI_News_Scraper_Dassult_Falcon = ABI_News_Scraper_Dassult_Falcon + 1

        End If
      Next

    Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try

  End Function
  Function Get_Description_From_Dassult_Falcon(ByVal link As String, ByVal temp_title As String) As String
    Get_Description_From_Dassult_Falcon = ""
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim orig_string_text As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim i As Integer = 0

    Try


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      orig_string_text = string_text

      spot_to_find = InStr(string_text, "<div class=" & Chr("34") & "df-intro" & Chr("34") & ">")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 21)


      spot_to_find = InStr(string_text, "Container_label")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 14)


      spot_to_find = InStr(string_text, ">")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 1)

      spot_to_find = InStr(string_text, "Container_label")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 14)

      spot_to_find = InStr(string_text, ">")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 1)

      spot_to_find = InStr(string_text, "publishingReusableFragmentIdSection")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find - 34)
      End If

      spot_to_find = InStr(string_text, "publishingReusableFragmentIdSection")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find - 34)
      End If

      spot_to_find = InStr(string_text, "publishingReusableFragmentIdSection")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find - 34)
      End If

      spot_to_find = InStr(string_text, "color:window" & Chr("34") & ">")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find - 13)
      End If


      spot_to_find = InStr(string_text, "Related Articles")

      If spot_to_find = 0 Then
        spot_to_find = spot_to_find
      End If

      string_text = Left(string_text, spot_to_find - 1)


      If Left(string_text, 3) = "img" Then
        spot_to_find = InStr(string_text, ">")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 1)
      ElseIf Left(string_text, 3) = "div" Then
        string_text = Right(string_text, Len(string_text) - 3)
      End If



      Get_Description_From_Dassult_Falcon = string_text



    Catch ex As Exception
      Response.Write(ex)
    Finally

    End Try
  End Function

    Function ABI_News_Scraper_Flight_Global(ByVal link As String, ByVal id As Integer) As Long
        ABI_News_Scraper_Flight_Global = 0
        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
        '   System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Ssl3
        Dim resp As System.Net.WebResponse = req.GetResponse
        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim i As Integer = 0
        Dim date_string As String = ""
        Dim description_string As String = ""
        Dim title_string As String = ""
        Dim insert_string As String = ""
        Dim link_to_go As String = ""
        Dim Str2 As System.IO.Stream
        Dim srRead2 As System.IO.StreamReader
        Dim req2 As System.Net.WebRequest
        Dim resp2 As System.Net.WebResponse
        Dim string_text3 As String = ""


        Try

            Str = resp.GetResponseStream
            srRead = New System.IO.StreamReader(Str)
            ' read all the text 
            string_text = srRead.ReadToEnd().ToString
            string_text = string_text

            spot_to_find = InStr(string_text, "<h1>All news</h1>")

            string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

            spot_to_find = InStr(string_text, "<li class="""">")
            string_text = Right(string_text, Len(string_text) - spot_to_find + 13)


            array_split = Split(string_text, "<li class="""">")

            For i = 1 To array_split.Length - 1

                string_text = array_split(i)


                spot_to_find = InStr(string_text, "href=")
                link_to_go = Right(string_text, Len(string_text) - spot_to_find - 5)
                spot_to_find = InStr(link_to_go, "class=")
                link_to_go = Left(link_to_go, spot_to_find - 3)

                If InStr(link_to_go, "http") = 0 Then
                    link_to_go = "https://www.flightglobal.com" & link_to_go
                End If


                spot_to_find = InStr(string_text, "<h3>")
                string_text = Right(string_text, Len(string_text) - spot_to_find - 5)
                spot_to_find = InStr(string_text, "</a>")
                title_string = Left(string_text, spot_to_find - 1)

                spot_to_find = InStr(title_string, ">")
                title_string = Right(title_string, Len(title_string) - spot_to_find - 1)


                title_string = Replace(title_string, "'", "")
                title_string = Replace(title_string, """", "")


                link_to_go = Replace(link_to_go, "'", "")
                link_to_go = Replace(link_to_go, """", "")




                '---get DATE STRING -- 11/12/2020 - MSW ----
                spot_to_find = InStr(string_text, "<p Class=""meta"">")
                string_text = Right(string_text, Len(string_text) - spot_to_find - 16)

                spot_to_find = InStr(string_text, ">")
                date_string = Right(date_string, Len(date_string) - spot_to_find - 1)

                spot_to_find = InStr(date_string, "</span>")
                date_string = Left(date_string, spot_to_find - 1)

                spot_to_find = InStr(date_string, "T")
                If spot_to_find > 0 Then
                    date_string = Left(date_string, spot_to_find - 1)
                Else
                    spot_to_find = spot_to_find
                End If

                date_string = change_date_format(date_string, 1, 2, 0)


                spot_to_find = InStr(string_text, "<p>")
                description_string = Right(string_text, Len(string_text) - spot_to_find - 3)

                spot_to_find = InStr(description_string, "</p>")
                description_string = Left(description_string, spot_to_find - 1)


                If record_exists(date_string, title_string, link_to_go) = False Then


                    description_string = replace_all_chars(description_string)
                    description_string = Left(description_string, 500)

                    If InStr(description_string, "flightglobal.com") Then
                        description_string = ""
                    End If

                    insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
                    insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
                    insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
                    insert_string = insert_string & " ) VALUES ( "
                    insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
                    insert_string = insert_string & link_to_go & "','" & id & "','','',"
                    insert_string = insert_string & "'',0"
                    insert_string = insert_string & " ) "

                    ' Response.Write(insert_string & "<br><br>")




                    'SETUP AND EXECUTE THE SQL INSERT COMMAND
                    Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                    ABI_News_Scraper_Flight_Global = ABI_News_Scraper_Flight_Global + 1

                End If
            Next

        Catch ex As Exception
            Response.Write(insert_string)
        Finally
        End Try

    End Function
    Function ABI_News_Scraper_Flight_Global_Old(ByVal link As String, ByVal id As Integer) As Long
        ABI_News_Scraper_Flight_Global_Old = 0
        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
        Dim resp As System.Net.WebResponse = req.GetResponse
        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim i As Integer = 0
        Dim date_string As String = ""
        Dim description_string As String = ""
        Dim title_string As String = ""
        Dim insert_string As String = ""
        Dim link_to_go As String = ""
        Dim Str2 As System.IO.Stream
        Dim srRead2 As System.IO.StreamReader
        Dim req2 As System.Net.WebRequest
        Dim resp2 As System.Net.WebResponse
        Dim string_text3 As String = ""


        Try

            Str = resp.GetResponseStream
            srRead = New System.IO.StreamReader(Str)
            ' read all the text 
            string_text = srRead.ReadToEnd().ToString
            string_text = string_text

            spot_to_find = InStr(string_text, "<h3>")

            string_text = Right(string_text, Len(string_text) - spot_to_find - 3)

            spot_to_find = InStr(string_text, "fgc-article-preview__description"">")
            string_text = Right(string_text, Len(string_text) - spot_to_find + 3)


            array_split = Split(string_text, "fgc-article-preview__description"">")

            For i = 1 To array_split.Length - 1

                string_text = array_split(i)


                spot_to_find = InStr(string_text, "href=")
                link_to_go = Right(string_text, Len(string_text) - spot_to_find - 5)
                spot_to_find = InStr(link_to_go, "class=")
                link_to_go = Left(link_to_go, spot_to_find - 3)

                If InStr(link_to_go, "http") = 0 Then
                    link_to_go = "https://www.flightglobal.com" & link_to_go
                End If


                spot_to_find = InStr(string_text, "fgc-article-preview_title"">")
                string_text = Right(string_text, Len(string_text) - spot_to_find - 26)
                spot_to_find = InStr(string_text, "</a>")
                title_string = Left(string_text, spot_to_find - 1)




                title_string = Replace(title_string, "'", "")
                title_string = Replace(title_string, """", "")


                link_to_go = Replace(link_to_go, "'", "")
                link_to_go = Replace(link_to_go, """", "")

                Try

                    req2 = System.Net.WebRequest.Create(link_to_go)
                    resp2 = req2.GetResponse

                    Str2 = resp2.GetResponseStream
                    srRead2 = New System.IO.StreamReader(Str2)
                    string_text3 = srRead2.ReadToEnd().ToString
                    string_text2 = string_text3

                    If InStr(Trim(string_text2), "article:published_time"" content=""") > 0 Then
                        string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), "article:published_time"" content=""") - 32)

                        If InStr(Trim(string_text2), "/>") > 0 Then
                            string_text2 = Left(string_text2, InStr(Trim(string_text2), "/>") - 2)
                            date_string = string_text2
                            date_string = change_date_format(date_string, 1, 0, 2)

                            date_string = Replace(date_string, "'", "")
                            date_string = Replace(date_string, """", "")
                        End If
                    End If

                    string_text2 = string_text3
                    If InStr(Trim(string_text2), "og:description"" content=""") > 0 Then
                        string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), "og:description"" content=""") - 24)

                        If InStr(Trim(string_text2), "/>") > 0 Then
                            string_text2 = Left(string_text2, InStr(Trim(string_text2), "/>") - 2)
                            description_string = string_text2
                            description_string = Replace(description_string, "'", "''")
                            description_string = Replace(description_string, """", "")
                        End If
                    End If



                Catch ex As Exception

                End Try


                If record_exists(date_string, title_string, link_to_go) = False Then


                    description_string = replace_all_chars(description_string)
                    description_string = Left(description_string, 500)

                    If InStr(description_string, "flightglobal.com") Then
                        description_string = ""
                    End If

                    insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
                    insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
                    insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
                    insert_string = insert_string & " ) VALUES ( "
                    insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
                    insert_string = insert_string & link_to_go & "','" & id & "','','',"
                    insert_string = insert_string & "'',0"
                    insert_string = insert_string & " ) "

                    ' Response.Write(insert_string & "<br><br>")




                    'SETUP AND EXECUTE THE SQL INSERT COMMAND
                    Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                    ABI_News_Scraper_Flight_Global_Old = ABI_News_Scraper_Flight_Global_Old + 1

                End If
            Next

        Catch ex As Exception
            Response.Write(insert_string)
        Finally
        End Try

    End Function
    Function ABI_News_Scraper_AIN_Online(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_AIN_Online = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Dim array_split_use() As String

    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text

            spot_to_find = InStr(string_text, "views-row-first")

            string_text = Right(string_text, Len(string_text) - spot_to_find - 15)

            spot_to_find = InStr(string_text, "<h2")
            string_text = Left(string_text, spot_to_find - 1)


            array_split = Split(string_text, "-field-pubdate")

            For i = 1 To array_split.Length - 1

                string_text = array_split(i)

                spot_to_find = InStr(string_text, "xsd:dateTime")
                date_string = Right(string_text, Len(string_text) - spot_to_find - 13)

                spot_to_find = InStr(date_string, "content=")
                date_string = Right(date_string, Len(date_string) - spot_to_find - 8)

                spot_to_find = InStr(date_string, "T")
                date_string = Left(date_string, spot_to_find - 1)

                array_split_use = Split(date_string, "-")
                date_string = array_split_use(1) & "/" & array_split_use(2) & "/" & array_split_use(0)


                spot_to_find = InStr(string_text, "href=")
                string_text = Right(string_text, Len(string_text) - spot_to_find - 5)

                spot_to_find = InStr(string_text, ">")
                link_to_go = Left(string_text, spot_to_find - 2)

                string_text = Right(string_text, Len(string_text) - spot_to_find)

                spot_to_find = InStr(string_text, "</a>")
                title_string = Left(string_text, spot_to_find - 1)

                If InStr(title_string, "typeof=""foaf:Image") > 0 Then
                    spot_to_find = InStr(string_text, "href=")
                    title_string = Right(string_text, Len(string_text) - spot_to_find - 5)

                    spot_to_find = InStr(title_string, "href=")
                    title_string = Right(title_string, Len(title_string) - spot_to_find - 5)

                    spot_to_find = InStr(title_string, ">")

                    title_string = Right(title_string, Len(title_string) - spot_to_find)

                    spot_to_find = InStr(title_string, "</a>")
                    title_string = Left(title_string, spot_to_find - 1)
                End If


                link_to_go = "http://www.ainonline.com" & link_to_go

                If InStr(link_to_go, "http") = 0 Then
                    link_to_go = "http://" & link_to_go
                End If

                title_string = Replace(title_string, "'", "&#39;")

                ' spot_to_find = InStr(string_text, "pubdate") 



                spot_to_find = InStr(string_text, "field-content teaser")
                If spot_to_find = 0 Then  ' then this is one of the side articles 
                    description_string = title_string
                Else
                    description_string = Right(string_text, Len(string_text) - spot_to_find - 21)
                End If

                spot_to_find = InStr(description_string, "</div>")
                description_string = Left(description_string, spot_to_find - 1)


                If record_exists(date_string, title_string, link_to_go) = False Then


                    description_string = replace_all_chars(description_string)
                    description_string = Left(description_string, 500)

                    title_string = replace_all_chars(title_string)

                    insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
                    insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
                    insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
                    insert_string = insert_string & " ) VALUES ( "
                    insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
                    insert_string = insert_string & link_to_go & "','" & id & "','','',"
                    insert_string = insert_string & "'',0"
                    insert_string = insert_string & " ) "

                    ' Response.Write(insert_string & "<br><br>")




                    'SETUP AND EXECUTE THE SQL INSERT COMMAND
                    Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                    ABI_News_Scraper_AIN_Online = ABI_News_Scraper_AIN_Online + 1

                End If
            Next

        Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try

  End Function

    Function ABI_News_Scraper_AIN_Online_Old(ByVal link As String, ByVal id As Integer) As Long
        ABI_News_Scraper_AIN_Online_Old = 0
        Dim Str As System.IO.Stream
        Dim srRead As System.IO.StreamReader
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
        Dim resp As System.Net.WebResponse = req.GetResponse
        Dim string_text As String = ""
        Dim string_text2 As String = ""
        Dim spot_to_find As Integer = 0
        Dim spot_to_find2 As Integer = 0
        Dim array_split() As String
        Dim i As Integer = 0
        Dim date_string As String = ""
        Dim description_string As String = ""
        Dim title_string As String = ""
        Dim insert_string As String = ""
        Dim link_to_go As String = ""
        Dim array_split_use() As String

        Try

            Str = resp.GetResponseStream
            srRead = New System.IO.StreamReader(Str)
            ' read all the text 
            string_text = srRead.ReadToEnd().ToString
            string_text = string_text

            spot_to_find = InStr(string_text, "views-row views-row-")

            string_text = Right(string_text, Len(string_text) - spot_to_find - 20)

            spot_to_find = InStr(string_text, "<h2")
            string_text = Left(string_text, spot_to_find - 1)


            array_split = Split(string_text, "views-row views-row-")

            For i = 0 To array_split.Length - 1

                string_text = array_split(i)

                If InStr(string_text, "Clean Sky") > 0 Then
                    string_text = string_text
                End If

                spot_to_find = InStr(string_text, "xsd:dateTime")
                date_string = Right(string_text, Len(string_text) - spot_to_find - 13)

                spot_to_find = InStr(date_string, "content=")
                date_string = Right(date_string, Len(date_string) - spot_to_find - 8)

                spot_to_find = InStr(date_string, "T")
                date_string = Left(date_string, spot_to_find - 1)

                array_split_use = Split(date_string, "-")
                date_string = array_split_use(1) & "/" & array_split_use(2) & "/" & array_split_use(0)


                spot_to_find = InStr(string_text, "href=")
                string_text = Right(string_text, Len(string_text) - spot_to_find - 5)

                spot_to_find = InStr(string_text, ">")
                link_to_go = Left(string_text, spot_to_find - 2)

                string_text = Right(string_text, Len(string_text) - spot_to_find)

                spot_to_find = InStr(string_text, "</a>")
                title_string = Left(string_text, spot_to_find - 1)

                If InStr(title_string, "typeof=""foaf:Image") > 0 Then
                    spot_to_find = InStr(string_text, "href=")
                    title_string = Right(string_text, Len(string_text) - spot_to_find - 5)

                    spot_to_find = InStr(title_string, "href=")
                    title_string = Right(title_string, Len(title_string) - spot_to_find - 5)

                    spot_to_find = InStr(title_string, ">")

                    title_string = Right(title_string, Len(title_string) - spot_to_find)

                    spot_to_find = InStr(title_string, "</a>")
                    title_string = Left(title_string, spot_to_find - 1)
                End If


                link_to_go = "http://www.ainonline.com" & link_to_go

                If InStr(link_to_go, "http") = 0 Then
                    link_to_go = "http://" & link_to_go
                End If

                title_string = Replace(title_string, "'", "&#39;")

                ' spot_to_find = InStr(string_text, "pubdate") 



                spot_to_find = InStr(string_text, "<div class=""field-content"">")
                If spot_to_find = 0 Then  ' then this is one of the side articles 
                    description_string = title_string
                Else
                    description_string = Right(string_text, Len(string_text) - spot_to_find - 26)
                End If


                spot_to_find = InStr(description_string, "</div>")
                description_string = Left(description_string, spot_to_find - 1)


                If record_exists(date_string, title_string, link_to_go) = False Then


                    description_string = replace_all_chars(description_string)
                    description_string = Left(description_string, 500)

                    insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
                    insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
                    insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
                    insert_string = insert_string & " ) VALUES ( "
                    insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
                    insert_string = insert_string & link_to_go & "','" & id & "','','',"
                    insert_string = insert_string & "'',0"
                    insert_string = insert_string & " ) "

                    ' Response.Write(insert_string & "<br><br>")




                    'SETUP AND EXECUTE THE SQL INSERT COMMAND
                    Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()

                    ABI_News_Scraper_AIN_Online_Old = ABI_News_Scraper_AIN_Online_Old + 1

                End If
            Next

        Catch ex As Exception
            Response.Write(insert_string)
        Finally
        End Try

    End Function
    Function ABI_News_Scraper_AVIATION_WEEK(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_AVIATION_WEEK = 0

    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader


    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse

    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text



      spot_to_find = InStr(string_text, "<div class=""teaser-content"">")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 32)

      spot_to_find = InStr(string_text, "Special Topics")
      string_text = Left(string_text, spot_to_find - 1)



      array_split = Split(string_text, "<div class=""teaser-content"">")

      For i = 1 To array_split.Length - 1
        string_text = array_split(i)




        spot_to_find = InStr(string_text, "created")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 8)
        spot_to_find = InStr(string_text, "</div>")
        date_string = Left(string_text, spot_to_find - 1)

        date_string = change_date_format(date_string, 0, 1, 2)


        spot_to_find = InStr(string_text, "title=")
        title_string = Right(string_text, Len(string_text) - spot_to_find - 6)
        link_to_go = Left(string_text, spot_to_find - 2)


        spot_to_find = InStr(title_string, ">")
        title_string = Left(title_string, spot_to_find - 2)

        spot_to_find = InStr(string_text, "<div class=""teaser-body"">")
        description_string = Right(string_text, Len(string_text) - spot_to_find - 24)
        spot_to_find = InStr(description_string, "<")
        description_string = Left(description_string, spot_to_find - 1)
        description_string = Replace(description_string, "     ", "")
        description_string = Trim(description_string)

        spot_to_find = InStr(link_to_go, "href=")
        link_to_go = Right(link_to_go, Len(link_to_go) - spot_to_find - 6)

        link_to_go = Trim(link) & "/" & Trim(link_to_go)

        If InStr(link_to_go, "http") = 0 Then
          link_to_go = "http://" & link_to_go
        End If


        If record_exists(date_string, title_string, link_to_go) = False Then


          description_string = replace_all_chars(description_string)
          description_string = Left(description_string, 500)



          insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
          insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
          insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
          insert_string = insert_string & " ) VALUES ( "
          insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
          insert_string = insert_string & link_to_go & "','" & id & "','','',"
          insert_string = insert_string & "'',0"
          insert_string = insert_string & " ) "

          '   Response.Write(insert_string & "<br><br>")


          'SETUP AND EXECUTE THE SQL INSERT COMMAND
          Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

          sqlComm.ExecuteNonQuery()
          sqlComm.Dispose()

          ABI_News_Scraper_AVIATION_WEEK = ABI_News_Scraper_AVIATION_WEEK + 1

        End If
      Next

    Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try



  End Function
  Function ABI_News_Scraper_FlyCorperate(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_FlyCorperate = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader


    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse

    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text



      spot_to_find = InStr(string_text, "FlyCorporate Business Aviation News")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 32)


      array_split = Split(string_text, "<item>")

      For i = 1 To array_split.Length - 1
        string_text = array_split(i)


        spot_to_find = InStr(string_text, "<title>")
        title_string = Right(string_text, Len(string_text) - spot_to_find - 6)
        spot_to_find = InStr(title_string, "</title>")
        title_string = Left(title_string, spot_to_find - 1)


        spot_to_find = InStr(string_text, "<link>")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 5)
        spot_to_find = InStr(string_text, "</link>")
        link_to_go = Left(string_text, spot_to_find - 1)


        spot_to_find = InStr(string_text, "<pubDate>")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 8)
        spot_to_find = InStr(string_text, "</pubDate>")
        date_string = Left(string_text, spot_to_find - 1)



        spot_to_find = InStr(string_text, "<description>")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 12)
        spot_to_find = InStr(string_text, "</description>")
        description_string = Left(string_text, spot_to_find - 1)



        If InStr(link_to_go, "http") = 0 Then
          link_to_go = "http://" & link_to_go
        End If

        date_string = change_date_format(date_string, 1, 0, 2)

        If record_exists(date_string, title_string, link_to_go) = False Then


          description_string = replace_all_chars(description_string)
          description_string = Left(description_string, 500)



          insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
          insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
          insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
          insert_string = insert_string & " ) VALUES ( "
          insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
          insert_string = insert_string & link_to_go & "','" & id & "','','',"
          insert_string = insert_string & "'',0"
          insert_string = insert_string & " ) "

          '  Response.Write(insert_string & "<br><br>")




          'SETUP AND EXECUTE THE SQL INSERT COMMAND
          Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

          sqlComm.ExecuteNonQuery()
          sqlComm.Dispose()

          ABI_News_Scraper_FlyCorperate = ABI_News_Scraper_FlyCorperate + 1

        End If
      Next

    Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try

  End Function
  Function ABI_News_Scraper_Rotorpad(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_Rotorpad = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text

      spot_to_find = InStr(string_text, "<div class=" & Chr("34") & "leadingarticles" & Chr("34") & ">")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 29)


      array_split = Split(string_text, "<h1 class=" & Chr("34") & "title" & Chr("34") & ">")

      For i = 1 To array_split.Length - 1
        string_text = array_split(i)
        string_text = Trim(string_text)

        spot_to_find = InStr(string_text, "</h1>")
        title_string = Left(string_text, spot_to_find - 1)
        title_string = Trim(title_string)

        spot_to_find = InStr(string_text, "<span class=" & Chr("34") & "created" & Chr("34") & ">")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 27)
        spot_to_find = InStr(string_text, "</span>")
        date_string = Left(string_text, spot_to_find - 1)

        date_string = change_date_format(date_string, 1, 0, 2)





        spot_to_find = InStr(string_text, "<a class=" & Chr("34") & "readmore-link" & Chr("34"))
        link_to_go = Right(string_text, Len(string_text) - spot_to_find - 22)
        spot_to_find = InStr(link_to_go, "href=")
        link_to_go = Right(link_to_go, Len(link_to_go) - spot_to_find - 5)
        spot_to_find = InStr(link_to_go, "title=")
        link_to_go = Left(link_to_go, spot_to_find - 3)



        If InStr(link_to_go, "http") = 0 Then
          link_to_go = "http://www.rotorpad.com" & link_to_go
        End If

        spot_to_find = InStr(string_text, "<p>")
        description_string = Right(string_text, Len(string_text) - spot_to_find - 2)
        spot_to_find = InStr(description_string, "<div")
        description_string = Left(description_string, spot_to_find - 1)




        If record_exists(date_string, title_string, link_to_go) = False Then


          description_string = replace_all_chars(description_string)
          description_string = Replace(description_string, "This is preliminary information, subject to change, and may contain errors. Any errors in this report will be corrected when the final report has been completed.", "")
          description_string = Left(description_string, 500)



          insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
          insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
          insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
          insert_string = insert_string & " ) VALUES ( "
          insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
          insert_string = insert_string & link_to_go & "','" & id & "','','',"
          insert_string = insert_string & "'',0"
          insert_string = insert_string & " ) "

          Response.Write(insert_string & "<br><br>")




          'SETUP AND EXECUTE THE SQL INSERT COMMAND
          Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

          sqlComm.ExecuteNonQuery()
          sqlComm.Dispose()

          ABI_News_Scraper_Rotorpad = ABI_News_Scraper_Rotorpad + 1

        End If
      Next

    Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try

  End Function
  Function ABI_News_Scraper_Avweb(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_Avweb = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text

      spot_to_find = InStr(string_text, "<h1>More News</h1>")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 18)


      spot_to_find = InStr(string_text, "Complete News")
      string_text = Left(string_text, spot_to_find - 1)


      array_split = Split(string_text, "<article>")

      For i = 1 To array_split.Length - 1
        string_text = array_split(i)

        spot_to_find = InStr(string_text, "href")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 4)


        spot_to_find = InStr(string_text, ">")
        link_to_go = Left(string_text, spot_to_find - 2)
        link_to_go = Replace(link_to_go, """", "")

        If InStr(link_to_go, "http") = 0 Then
          link_to_go = "http://www.avweb.com" & link_to_go
        End If


        string_text = Right(Trim(string_text), Len(string_text) - spot_to_find)

        spot_to_find = InStr(string_text, "</a>")
        title_string = Left(string_text, spot_to_find - 1)
        title_string = replace_all_chars(title_string)

        spot_to_find = InStr(string_text, "|</span>")
        string_text = Right(Trim(string_text), Len(string_text) - spot_to_find - 8)

        spot_to_find = InStr(string_text, "</p>")
        date_string = Left(string_text, spot_to_find - 1)


        date_string = change_date_format(date_string, 0, 1, 2)

        string_text = Right(Trim(string_text), Len(string_text) - spot_to_find - 3)


        spot_to_find = InStr(string_text, "<a class")
        description_string = Left(string_text, spot_to_find - 1)
        description_string = Replace(description_string, "<p>", "")
        description_string = Trim(description_string)



        If record_exists(date_string, title_string, link_to_go) = False Then


          description_string = replace_all_chars(description_string)
          description_string = Left(description_string, 500)


          insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
          insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
          insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
          insert_string = insert_string & " ) VALUES ( "
          insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
          insert_string = insert_string & link_to_go & "','" & id & "','','',"
          insert_string = insert_string & "'',0"
          insert_string = insert_string & " ) "

          ' Response.Write(insert_string & "<br><br>")



          'SETUP AND EXECUTE THE SQL INSERT COMMAND
          Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

          sqlComm.ExecuteNonQuery()
          sqlComm.Dispose()
          ABI_News_Scraper_Avweb = ABI_News_Scraper_Avweb + 1

        End If
      Next

    Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try

  End Function

  Function ABI_News_Scraper_Vertical(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_Vertical = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Dim every_three As Integer = 1
    Dim Str2 As System.IO.Stream
    Dim srRead2 As System.IO.StreamReader
    Dim req2 As System.Net.WebRequest
    Dim resp2 As System.Net.WebResponse
    Dim original_string_text2 As String = ""
    Dim string_text3 As String = ""

    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text

            spot_to_find = InStr(string_text, "press-releases")
            string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

            array_split = Split(string_text, "press-releases ")

            For i = 2 To array_split.Length - 1
                string_text = array_split(i)


                spot_to_find = InStr(string_text, ">")
                string_text = Right(string_text, Len(string_text) - spot_to_find - 1)

                spot_to_find = InStr(string_text, "</span>")
                date_string = Left(string_text, spot_to_find - 1)


                date_string = change_date_format(date_string, 0, 1, 2)


                spot_to_find = InStr(string_text, "href")
                string_text = Right(string_text, Len(string_text) - spot_to_find - 4)


                spot_to_find = InStr(string_text, ">")
                link_to_go = Left(string_text, spot_to_find - 2)
                link_to_go = Replace(link_to_go, """", "")
                link_to_go = Replace(link_to_go, "'", "")

                string_text = Right(Trim(string_text), Len(string_text) - spot_to_find)

                spot_to_find = InStr(string_text, "</h5>")
                title_string = Left(string_text, spot_to_find - 1)
                title_string = Replace(title_string, "</a>", "")

                spot_to_find = InStr(string_text, "tp-excerpt")
                string_text = Right(Trim(string_text), Len(string_text) - spot_to_find - 12)


                spot_to_find = InStr(string_text, "</p>")
                description_string = Left(string_text, spot_to_find - 1)
                description_string = Replace(description_string, "<p>", "")
                description_string = Trim(description_string)


                'Try

                '    req2 = System.Net.WebRequest.Create(link_to_go)
                '    resp2 = req2.GetResponse

                '    Str2 = resp2.GetResponseStream
                '    srRead2 = New System.IO.StreamReader(Str2)
                '    string_text3 = srRead2.ReadToEnd().ToString
                '    string_text2 = string_text3


                '    If InStr(Trim(string_text2), "entry-date published updated") > 0 Then
                '        string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), "entry-date published updated") - 28)

                '        If InStr(Trim(string_text2), ">") > 0 Then
                '            string_text2 = Right(string_text2, Len(string_text2) - InStr(Trim(string_text2), ">") - 1)

                '            If InStr(Trim(string_text2), "</time>") > 0 Then
                '                string_text2 = Left(string_text2, InStr(Trim(string_text2), "</time>") - 1)
                '                date_string = string_text2
                '                date_string = change_date_format(date_string, 0, 1, 2)
                '            End If
                '        End If
                '    End If


                'Catch ex As Exception

                'End Try

                If record_exists(date_string, title_string, link_to_go) = False Then


                    description_string = replace_all_chars(description_string)

                    description_string = Left(description_string, 500)


                    insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
                    insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
                    insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
                    insert_string = insert_string & " ) VALUES ( "
                    insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
                    insert_string = insert_string & link_to_go & "','" & id & "','','',"
                    insert_string = insert_string & "'',0"
                    insert_string = insert_string & " ) "

                    ' Response.Write(insert_string & "<br><br>")



                    ''''  SETUP And EXECUTE THE SQL INSERT COMMAND
                    Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()
                    ABI_News_Scraper_Vertical = ABI_News_Scraper_Vertical + 1

                End If
            Next

        Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try

  End Function
  Function Get_Date_From_Avweb(ByVal link As String, ByVal temp_title As String) As String
    Get_Date_From_Avweb = ""
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim orig_string_text As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim i As Integer = 0

    Try


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text
      orig_string_text = string_text

      spot_to_find = InStr(string_text, "<span class=" & Chr("34") & "headline" & Chr("34") & ">")
      string_text = Left(string_text, spot_to_find - 1)


      spot_to_find = InStr(string_text, "<p class=" & Chr("34") & "copy" & Chr("34") & ">")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 15)

      spot_to_find = InStr(string_text, "</p>")
      string_text = Left(string_text, spot_to_find - 1)

      Get_Date_From_Avweb = string_text



    Catch ex As Exception
      Response.Write(ex)
    Finally

    End Try
  End Function
  Function ABI_News_Scraper_BART(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_BART = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("https://www.bartintl.com/news")   ' switched from the link 
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim array_split_2() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Dim temp_dater_string As String = ""
    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text

      spot_to_find = InStr(string_text, "<div id=""MainContent"">")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 18)



      array_split = Split(string_text, "panel-heading text-capitalize"">")

      For i = 1 To array_split.Length - 1
        string_text = array_split(i)
        string_text = Trim(string_text)

        spot_to_find = InStr(string_text, "</div>")
        title_string = Left(string_text, spot_to_find - 1)



        spot_to_find = InStr(string_text, "<div class=""date"">")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 17)
        spot_to_find = InStr(string_text, "</div>")
        date_string = Left(string_text, spot_to_find - 1)
        date_string = Replace(date_string, "/", "-")
        date_string = change_date_format(date_string, 1, 0, 2)

        spot_to_find = InStr(string_text, "readmore hide")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 13)

        spot_to_find = InStr(string_text, "data-url=""")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 9)

        spot_to_find = InStr(string_text, """></div>")
        link_to_go = Left(string_text, spot_to_find - 1)
        If InStr(link_to_go, "http") = 0 Then
          link_to_go = "https://www.bartintl.com/" & Trim(link_to_go)
        End If
        link_to_go = Replace(link_to_go, "'", "")
        link_to_go = Replace(link_to_go, """", "")

        spot_to_find = InStr(string_text, "<p>")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 2)
        spot_to_find = InStr(string_text, "</p>")
        description_string = Left(string_text, spot_to_find - 1)





        If record_exists(date_string, title_string, link_to_go) = False Then


          description_string = replace_all_chars(description_string)
          description_string = Left(description_string, 500)



          insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
          insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
          insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
          insert_string = insert_string & " ) VALUES ( "
          insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
          insert_string = insert_string & link_to_go & "','" & id & "','','',"
          insert_string = insert_string & "'',0"
          insert_string = insert_string & " ) "

          '    Response.Write(insert_string & "<br><br>")




          'SETUP AND EXECUTE THE SQL INSERT COMMAND
          Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

          sqlComm.ExecuteNonQuery()
          sqlComm.Dispose()

          ABI_News_Scraper_BART = ABI_News_Scraper_BART + 1

        End If
      Next

    Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try

  End Function
  'Function ABI_News_Scraper_BART(ByVal link As String, ByVal id As Integer) As Long
  '  ABI_News_Scraper_BART = 0
  '  Dim Str As System.IO.Stream
  '  Dim srRead As System.IO.StreamReader
  '  Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
  '  Dim resp As System.Net.WebResponse = req.GetResponse
  '  Dim string_text As String = ""
  '  Dim string_text2 As String = ""
  '  Dim spot_to_find As Integer = 0
  '  Dim spot_to_find2 As Integer = 0
  '  Dim array_split() As String
  '  Dim array_split_2() As String
  '  Dim i As Integer = 0
  '  Dim date_string As String = ""
  '  Dim description_string As String = ""
  '  Dim title_string As String = ""
  '  Dim insert_string As String = ""
  '  Dim link_to_go As String = ""
  '  Dim temp_dater_string As String = ""
  '  Try

  '    Str = resp.GetResponseStream
  '    srRead = New System.IO.StreamReader(Str)
  '    ' read all the text 
  '    string_text = srRead.ReadToEnd().ToString
  '    string_text = string_text

  '    spot_to_find = InStr(string_text, "<div class=" & Chr("34") & "inner" & Chr("34") & ">")
  '    string_text = Right(string_text, Len(string_text) - spot_to_find - 18)

  '    spot_to_find = InStr(string_text, "<div class=" & Chr("34") & "inner" & Chr("34") & ">")
  '    string_text = Right(string_text, Len(string_text) - spot_to_find - 18)


  '    array_split = Split(string_text, "<div class=" & Chr("34") & "inner" & Chr("34") & ">")

  '    For i = 0 To array_split.Length - 1
  '      string_text = array_split(i)
  '      string_text = Trim(string_text)

  '      spot_to_find = InStr(string_text, "<span>")
  '      string_text = Right(string_text, Len(string_text) - spot_to_find - 5)

  '      spot_to_find = InStr(string_text, "</span>")
  '      date_string = Left(string_text, spot_to_find - 1)

  '      array_split_2 = Split(date_string, "/")

  '      temp_dater_string = array_split_2(1) & "/" & array_split_2(0) & "/" & array_split_2(2)

  '      date_string = temp_dater_string

  '      spot_to_find = InStr(string_text, "href=")
  '      string_text = Right(string_text, Len(string_text) - spot_to_find - 5)

  '      spot_to_find = InStr(string_text, "title")
  '      link_to_go = Left(string_text, spot_to_find - 3)


  '      spot_to_find = InStr(string_text, ">")
  '      string_text = Right(string_text, Len(string_text) - spot_to_find)

  '      spot_to_find = InStr(string_text, "</a>")
  '      title_string = Left(string_text, spot_to_find - 1)

  '      title_string = replace_all_chars(title_string)


  '      spot_to_find = InStr(string_text, "<p>")
  '      string_text = Right(string_text, Len(string_text) - spot_to_find - 2)

  '      spot_to_find = InStr(string_text, "</p>")
  '      description_string = Left(string_text, spot_to_find - 1)

  '      If InStr(link_to_go, "http") = 0 Then
  '        link_to_go = "http://www.bartintl.com" & link_to_go
  '      End If



  '      If record_exists(date_string, title_string, link_to_go) = False Then


  '        description_string = replace_all_chars(description_string)
  '        description_string = Left(description_string, 500)



  '        insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
  '        insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
  '        insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
  '        insert_string = insert_string & " ) VALUES ( "
  '        insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
  '        insert_string = insert_string & link_to_go & "','" & id & "','','',"
  '        insert_string = insert_string & "'',0"
  '        insert_string = insert_string & " ) "

  '        '    Response.Write(insert_string & "<br><br>")




  '        'SETUP AND EXECUTE THE SQL INSERT COMMAND
  '        Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

  '        sqlComm.ExecuteNonQuery()
  '        sqlComm.Dispose()

  '        ABI_News_Scraper_BART = ABI_News_Scraper_BART + 1

  '      End If
  '    Next

  '  Catch ex As Exception
  '    Response.Write(insert_string)
  '  Finally
  '  End Try

  'End Function
  Function ABI_News_Scraper_Corp_Jet(ByVal link As String, ByVal id As Integer) As Long
    ABI_News_Scraper_Corp_Jet = 0
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader
    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(link)
    Dim resp As System.Net.WebResponse = req.GetResponse
    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim array_split_2() As String
    Dim i As Integer = 0
    Dim date_string As String = ""
    Dim description_string As String = ""
    Dim title_string As String = ""
    Dim insert_string As String = ""
    Dim link_to_go As String = ""
    Dim temp_dater_string As String = ""
    Try

      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString
      string_text = string_text

      spot_to_find = InStr(string_text, "Recent articles")
      string_text = Right(string_text, Len(string_text) - spot_to_find - 14)

      spot_to_find = InStr(string_text, "<a href=" & Chr("34"))
      string_text = Right(string_text, Len(string_text) - spot_to_find - 6)


      array_split = Split(string_text, "<a href=" & Chr("34"))

      For i = 0 To array_split.Length - 1
        string_text = array_split(i)
        string_text = Trim(string_text)

        spot_to_find = InStr(string_text, ">")
        link_to_go = Left(string_text, spot_to_find - 2)

        string_text = Right(string_text, Len(string_text) - spot_to_find)

        spot_to_find = InStr(string_text, "</a>")
        title_string = Left(string_text, spot_to_find - 1)

        spot_to_find = InStr(string_text, "<span")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 5)

        spot_to_find = InStr(string_text, "</span>")
        date_string = Left(string_text, spot_to_find - 1)

        spot_to_find = InStr(date_string, ">")
        date_string = Right(date_string, Len(date_string) - spot_to_find)
        date_string = Trim(date_string)

        date_string = change_date_format(date_string, 0, 1, 2)


        spot_to_find = InStr(string_text, "<span>")
        string_text = Right(string_text, Len(string_text) - spot_to_find - 5)

        spot_to_find = InStr(string_text, "</span>")
        description_string = Left(string_text, spot_to_find - 1)


        If InStr(link_to_go, "http") = 0 Then
          link_to_go = "http://corpjetfin.live.subhub.com" & link_to_go
        End If



        If record_exists(date_string, title_string, link_to_go) = False Then


          description_string = replace_all_chars(description_string)
          description_string = Left(description_string, 500)



          insert_string = "Insert into ABI_News_Links_Temp(tmpnewslnk_date, tmpnewslnk_title, tmpnewslnk_description,"
          insert_string = insert_string & " tmpnewslnk_web_address, tmpnewslnk_source_id, tmpnewslnk_make_name, tmpnewslnk_amod_id, "
          insert_string = insert_string & " tmpnewslnk_abi_make_name, tmpnewslnk_abi_amod_id"
          insert_string = insert_string & " ) VALUES ( "
          insert_string = insert_string & "'" & date_string & "','" & title_string & "','" & description_string & "','"
          insert_string = insert_string & link_to_go & "','" & id & "','','',"
          insert_string = insert_string & "'',0"
          insert_string = insert_string & " ) "

          '  Response.Write(insert_string & "<br><br>")




          'SETUP AND EXECUTE THE SQL INSERT COMMAND
          Dim sqlComm As New SqlClient.SqlCommand(insert_string, MySqlConn_JETNET2) 'MySql.Data.MySqlClient.MySqlCommand(sQuery, MyMTDConn) - Amanda - 7/18/2011

          sqlComm.ExecuteNonQuery()
          sqlComm.Dispose()

          ABI_News_Scraper_Corp_Jet = ABI_News_Scraper_Corp_Jet + 1

        End If
      Next

    Catch ex As Exception
      Response.Write(insert_string)
    Finally
    End Try

  End Function

  Public Function record_exists(ByVal temp_date As String, ByVal temp_title As String, ByVal temp_link As String) As Boolean
    record_exists = False
    Dim temp_query As String = ""

    ' temp_query = " Select * from ABI_News_Links_Temp where tmpnewslnk_date = '" & temp_date & "' and  tmpnewslnk_title = '" & temp_title & "' and  tmpnewslnk_web_address = '" & temp_link & "' "

    temp_query = " Select * from ABI_News_Links_Temp where tmpnewslnk_title = '" & temp_title & "' "

    MySqlCommand_JETNET.CommandText = temp_query
    MyAircraftReader_JETNET = MySqlCommand_JETNET.ExecuteReader()
    If MyAircraftReader_JETNET.HasRows Then
      record_exists = True
    Else
    End If
    MyAircraftReader_JETNET.Close()


    'temp_query = " Select * from ABI_News_Links where abinewslnk_date = '" & temp_date & "' and  abinewslnk_title = '" & temp_title & "' and  abinewslnk_web_address = '" & temp_link & "' "


    'MySqlCommand_JETNET3.CommandText = temp_query
    'MyAircraftReader_JETNET3 = MySqlCommand_JETNET3.ExecuteReader()
    'If MyAircraftReader_JETNET3.HasRows Then
    '    record_exists = True
    'Else
    'End If
    'MyAircraftReader_JETNET3.Close()


  End Function

  Function replace_all_chars(ByVal string_to_replace As String) As String
    Dim i As Integer = 0
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim temp_string As String = ""
    Dim temp_string2 As String = ""

    If InStr(string_to_replace, "Download High-Resolution Image") > 0 Then
      spot_to_find = InStr(string_to_replace, "Download High-Resolution Image")
      temp_string = Left(string_to_replace, spot_to_find - 1)
      spot_to_find2 = InStr(string_to_replace, "MB)")
      string_to_replace = Right(string_to_replace, Len(string_to_replace) - spot_to_find2 - 2)
      string_to_replace = temp_string & string_to_replace
    End If


    ' incase there is a lingering one 
    If Left(string_to_replace, 2) = "" & Chr("34") & ">" Then
      string_to_replace = Right(string_to_replace, Len(string_to_replace) - 2)
    ElseIf Left(string_to_replace, 1) = ">" Then
      string_to_replace = Right(string_to_replace, Len(string_to_replace) - 1)
    End If



    If InStr(string_to_replace, "&lt;p&gt;&lt;a href") > 0 Then
      string_to_replace = Left(string_to_replace, InStr(string_to_replace, "&lt;p&gt;&lt;a href") - 1)
    End If


        string_to_replace = Replace(string_to_replace, Chr("10"), "")  ' line break - replace 

        string_to_replace = Replace(string_to_replace, "<br>", "")
    string_to_replace = Replace(string_to_replace, "<br/>", "")
    string_to_replace = Replace(string_to_replace, "<br />", "")

    string_to_replace = Replace(string_to_replace, "<p>", "")
    string_to_replace = Replace(string_to_replace, "</p>", "")

    string_to_replace = Replace(string_to_replace, "<em>", "")
    string_to_replace = Replace(string_to_replace, "</em>", "")

    string_to_replace = Replace(string_to_replace, "<ul>", "")
    string_to_replace = Replace(string_to_replace, "</ul>", "")
    string_to_replace = Replace(string_to_replace, "<li>", "-")
    string_to_replace = Replace(string_to_replace, "</li>", "")

    string_to_replace = Replace(string_to_replace, "<strong>", "")
    string_to_replace = Replace(string_to_replace, "</strong>", "")
    string_to_replace = Replace(string_to_replace, "<b>", "")
    string_to_replace = Replace(string_to_replace, "</b>", "")

    string_to_replace = Replace(string_to_replace, "<td>", "")
    string_to_replace = Replace(string_to_replace, "</td>", "")
    string_to_replace = Replace(string_to_replace, "<tr>", "")
    string_to_replace = Replace(string_to_replace, "</tr>", "")
    string_to_replace = Replace(string_to_replace, "<table>", "")
    string_to_replace = Replace(string_to_replace, "</table>", "")



    string_to_replace = Replace(string_to_replace, "<div>", "")
    string_to_replace = Replace(string_to_replace, "</div>", "")
    string_to_replace = Replace(string_to_replace, "<div align=" & Chr("34") & "center" & Chr("34") & ">", "")
    string_to_replace = Replace(string_to_replace, "<div align=" & Chr("34") & "right" & Chr("34") & ">", "")
    string_to_replace = Replace(string_to_replace, "<div align=" & Chr("34") & "left" & Chr("34") & ">", "")
    string_to_replace = Replace(string_to_replace, "<div align=" & Chr("34") & "text-align:left" & Chr("34") & ">", "")
    string_to_replace = Replace(string_to_replace, "<div align=" & Chr("34") & "text-align:center" & Chr("34") & ">", "")
    string_to_replace = Replace(string_to_replace, "<div align=" & Chr("34") & "text-align:right" & Chr("34") & ">", "")



    string_to_replace = Replace(string_to_replace, "&#039;", "")
    string_to_replace = Replace(string_to_replace, Chr("34"), "'")
    string_to_replace = Replace(string_to_replace, "&quot;", "")
    string_to_replace = Replace(string_to_replace, "&ldquo;", "")
    string_to_replace = Replace(string_to_replace, "&rsquo;", "")
    string_to_replace = Replace(string_to_replace, "&#8221;", "'") ' right double quote
    string_to_replace = Replace(string_to_replace, "&#8212;", "-") ' em dash , which is a long dash 
    string_to_replace = Replace(string_to_replace, "&lt;p&gt;", " ")
    string_to_replace = Replace(string_to_replace, "—", "-")
        string_to_replace = Replace(string_to_replace, "'", "")



        For i = 0 To 25
      string_to_replace = string_to_replace
      spot_to_find = 0
      spot_to_find2 = 0

      If InStr(string_to_replace, "<") > 0 Then
        spot_to_find = InStr(string_to_replace, "<")
      Else
        i = 26
      End If
      If InStr(string_to_replace, "<") > 0 Then
        spot_to_find2 = InStr(string_to_replace, ">")
      Else
        i = 26
      End If

      If spot_to_find > 0 And spot_to_find2 > 0 And (spot_to_find2 > spot_to_find) And (spot_to_find2 < Len(string_to_replace)) And (spot_to_find < Len(string_to_replace)) Then
        temp_string = Left(string_to_replace, spot_to_find - 1)
        temp_string2 = Right(string_to_replace, Len(string_to_replace) - spot_to_find2)
        string_to_replace = temp_string & temp_string2
      End If

    Next
    string_to_replace = Replace(string_to_replace, "</a>", "")
    string_to_replace = Replace(string_to_replace, "<", "")
    string_to_replace = Replace(string_to_replace, ">", "")


    For i = 0 To 5
      string_to_replace = Replace(string_to_replace, "   ", " ")
    Next

    replace_all_chars = string_to_replace
  End Function

  Function change_date_format(ByVal date_original As String, ByVal place_of_month As Integer, ByVal place_of_day As Integer, ByVal place_of_year As Integer) As String
    Dim i As Integer = 0
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim temp_string As String = ""
    Dim temp_string2 As String = ""
    Dim temp_month As String = ""
    Dim temp_day As String = ""
    Dim temp_year As String = ""

    Dim array_split() As String

    temp_string = date_original

    temp_string = Replace(temp_string, "Sunday", "")
    temp_string = Replace(temp_string, "Monday", "")
    temp_string = Replace(temp_string, "Tuesday", "")
    temp_string = Replace(temp_string, "Wednesday", "")
    temp_string = Replace(temp_string, "Thursday", "")
    temp_string = Replace(temp_string, "Friday", "")
    temp_string = Replace(temp_string, "Saturday", "")


    temp_string = Replace(temp_string, "Sun", "")
    temp_string = Replace(temp_string, "Mon", "")
    temp_string = Replace(temp_string, "Tue", "")
    temp_string = Replace(temp_string, "Wed", "")
    temp_string = Replace(temp_string, "Thu", "")
    temp_string = Replace(temp_string, "Fri", "")
    temp_string = Replace(temp_string, "Sat", "")

    temp_string = Replace(temp_string, ".", "")
    temp_string = Replace(temp_string, ",", "")
    temp_string = Trim(temp_string)


    If Left(temp_string, 1) = " " Then
      temp_string = Right(temp_string, Len(temp_string) - 1)
    End If


    If InStr(temp_string, "-") > 0 Then
      temp_string = Replace(temp_string, "-", " ")
    End If

    array_split = Split(temp_string, " ")

    temp_month = array_split(place_of_month)

    If Trim(temp_month) = "January" Or Trim(temp_month) = "Jan" Then
      temp_month = "01"
    ElseIf Trim(temp_month) = "February" Or Trim(temp_month) = "Feb" Then
      temp_month = "02"
    ElseIf Trim(temp_month) = "March" Or Trim(temp_month) = "Mar" Then
      temp_month = "03"
    ElseIf Trim(temp_month) = "April" Or Trim(temp_month) = "Apr" Then
      temp_month = "04"
    ElseIf Trim(temp_month) = "May" Or Trim(temp_month) = "May" Then
      temp_month = "05"
    ElseIf Trim(temp_month) = "June" Or Trim(temp_month) = "Jun" Then
      temp_month = "06"
    ElseIf Trim(temp_month) = "July" Or Trim(temp_month) = "Jul" Then
      temp_month = "07"
    ElseIf Trim(temp_month) = "August" Or Trim(temp_month) = "Aug" Then
      temp_month = "08"
    ElseIf Trim(temp_month) = "September" Or Trim(temp_month) = "Sep" Then
      temp_month = "09"
    ElseIf Trim(temp_month) = "October" Or Trim(temp_month) = "Oct" Then
      temp_month = "10"
    ElseIf Trim(temp_month) = "November" Or Trim(temp_month) = "Nov" Then
      temp_month = "11"
    ElseIf Trim(temp_month) = "December" Or Trim(temp_month) = "Dec" Then
      temp_month = "12"
    Else
      If Not IsNumeric(temp_month) Then
        temp_month = "01"
      End If
    End If



    temp_day = array_split(place_of_day)
    temp_year = array_split(place_of_year)

    temp_string = temp_month & "/" & temp_day & "/" & temp_year


    change_date_format = temp_string
  End Function

#End Region


#Region "STOLEN_EVO_FUNCTIONS"

  Public Sub generateAdminReport(ByVal nReportID As Integer, ByRef out_ReportString As String, Optional ByVal nSubID As Long = 0, Optional ByVal sRptType As String = "", Optional ByVal bUseClientConn As Boolean = False, Optional ByVal bIsAdminReport As Boolean = False, Optional ByVal is_aerodex_limited As Boolean = False)

    Dim results_table As New DataTable
    Dim htmlOut As New StringBuilder

    Dim sRptQuery As String = ""
    Dim sRptTitle As String = ""
    Dim sRptReportType As String = ""
    Dim temp_spot1 As Integer = 0
    Dim temp_sport2 As Integer = 0
    Dim temp_start_string As String = ""
    Dim temp_end_string As String = ""
    Dim temp_string As String = ""
    Dim temp_add_string As String = ""
    Dim temp_final_string As String = ""

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader = Nothing
    Dim SqlException As SqlClient.SqlException : SqlException = Nothing

    Try

      results_table = getAdminReportListDataTable(nReportID, nSubID, sRptType, bIsAdminReport, is_aerodex_limited)

      If Not IsNothing(results_table) Then

        If results_table.Rows.Count > 0 Then

          sRptQuery = results_table.Rows(0).Item("sqlrep_query").ToString.Trim
          sRptTitle = results_table.Rows(0).Item("sqlrep_title").ToString.Trim
          If Not IsDBNull(results_table.Rows(0).Item("sqlrep_type")) Then
            sRptReportType = results_table.Rows(0).Item("sqlrep_type").ToString.Trim
          End If

        End If

      End If

      If Trim(sRptReportType) = "Aircraft" Or Trim(sRptReportType) = "Company" Or Left(Trim(sRptReportType), 8) = "Aircraft" Then

        temp_string = Trim(sRptQuery)

        temp_final_string = Replace(temp_string, "/* INSERT SUBSCRIPTION */", temp_add_string)
        sRptQuery = temp_final_string
        'If InStr(UCase(Trim(temp_string)), "GROUP BY ") > 0 Then
        '  temp_spot1 = InStr(UCase(Trim(temp_string)), "GROUP BY ")
        'ElseIf InStr(UCase(Trim(temp_string)), "ORDER BY ") > 0 Then
        '  temp_spot1 = InStr(UCase(Trim(temp_string)), "ORDER BY ")
        'End If

        'If temp_spot1 > 0 Then
        '  temp_start_string = Left(temp_string, temp_spot1 - 1)

        '  temp_end_string = Right(temp_string, Len(temp_string) - temp_spot1 + 1)


        '  If Trim(sRptReportType) = "Aircraft" Then
        '    temp_add_string = commonEvo.BuildProductCodeCheckWhereClause(HttpContext.Current.Session.Item("localSubscription").crmHelicopter_Flag, HttpContext.Current.Session.Item("localSubscription").crmBusiness_Flag, HttpContext.Current.Session.Item("localSubscription").crmCommercial_Flag, False, HttpContext.Current.Session.Item("localSubscription").crmYacht_Flag, False, True)
        '  ElseIf Trim(sRptReportType) = "Company" Then
        '    temp_add_string = commonEvo.BuildCompanyProductCodeCheckWhereClause(HttpContext.Current.Session.Item("localSubscription").crmHelicopter_Flag, HttpContext.Current.Session.Item("localSubscription").crmBusiness_Flag, HttpContext.Current.Session.Item("localSubscription").crmCommercial_Flag, False, HttpContext.Current.Session.Item("localSubscription").crmYacht_Flag, False)
        '  End If

        '  temp_final_string = temp_start_string & " " & temp_add_string & " " & temp_end_string

        'End If

        'sRptQuery = temp_final_string
      End If

      'HttpContext.Current.Session.Item("localUser").crmUser_DebugText += "<br /><br /><b>generateAdminReport(ByVal nReportID As Integer, ByRef out_ReportString As String, Optional ByVal nSubID As Long = 0, Optional ByVal bUseClientConn As Boolean = False)</b><br />" + sRptQuery.ToString


      SqlConn.ConnectionString = JETNET_LIVE_SQL_CONN

      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 240

      SqlCommand.CommandText = sRptQuery.Trim

      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        results_table = New DataTable
        results_table.Load(SqlReader)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = results_table.GetErrors()
        HttpContext.Current.Session.Item("localUser").crmUser_DebugText += "<br /><br /><b>Error in generateAdminReport load datatable</b><br /> " + constrExc.Message
      End Try

      If Not IsNothing(results_table) Then

        If results_table.Rows.Count > 0 Then

          htmlOut.Append(include_excel_admin_report_style())
          htmlOut.Append("<table border=""1"" cellpadding=""2"" cellspacing=""0"">")

          ' first add the report title
          htmlOut.Append("<tr><td align=""left"" valign=""middle"" colspan=""" + results_table.Columns.Count.ToString + """><b>" + sRptTitle.Trim + "</b></td></tr>")

          ' second generate the header based off the column names in the datatable
          htmlOut.Append("<tr bgcolor=""#CCCCCC"">")
          For Each c As DataColumn In results_table.Columns
            htmlOut.Append("<td align=""left"">" + c.ColumnName.ToUpper.Replace("CCOUNT", "COUNT").Trim + "</td>")
          Next
          htmlOut.Append("</tr>")

          ' second display the report data based off the column names in the datatable
          For Each r As DataRow In results_table.Rows

            htmlOut.Append("<tr>")

            ' ramble through each "column name" and display data
            For Each c As DataColumn In results_table.Columns
              htmlOut.Append("<td align=""left"" valign=""top"">" + r.Item(c.ColumnName).ToString.Trim + "</td>")
            Next

            htmlOut.Append("</tr>")

          Next

          htmlOut.Append("</table>")

        End If

      End If

    Catch ex As Exception

      HttpContext.Current.Session.Item("localUser").crmUser_DebugText += "Error in generateAdminReport(ByVal nReportID As Integer, ByRef out_ReportString As String, Optional ByVal nSubID As Long = 0, Optional ByVal bUseClientConn As Boolean = False) " + ex.Message

    Finally

    End Try

    'return resulting html string
    out_ReportString = htmlOut.ToString
    htmlOut = Nothing
    results_table = Nothing

  End Sub

  Public Function include_excel_admin_report_style() As String

    Dim htmlOut = New StringBuilder()

    htmlOut.Append("<style type='text/css'>")
    htmlOut.Append("  td.textformat {mso-number-format:'\@'}")
    htmlOut.Append("  td.textdate {mso-number-format:'Short Date'}")
    htmlOut.Append("</style>")

    Return htmlOut.ToString
    htmlOut = Nothing

  End Function


  Public Function getAdminReportListDataTable(ByVal nReportID As Integer, ByVal nSubID As Long, Optional ByVal sRptType As String = "", Optional ByVal bIsAdminReport As Boolean = False, Optional ByVal is_aerodex_limited As Boolean = False) As DataTable

    Dim atemptable As New DataTable
    Dim sQuery = New StringBuilder()

    Dim SqlConn As New SqlClient.SqlConnection
    Dim SqlCommand As New SqlClient.SqlCommand
    Dim SqlReader As SqlClient.SqlDataReader
    Dim SqlException As SqlClient.SqlException : SqlException = Nothing 'sqlrep_level = 'JETNET'

    Try

      sQuery.Append("SELECT * FROM SQL_Report WITH(NOLOCK)")

      sQuery.Append(" WHERE sqlrep_level = 'JETNET' AND sqlrep_sub_id = 0 AND sqlrep_id = " + nReportID.ToString)


      'MSW 11/12/15 - true would mean the person is aerodex, so they can only see ones where aerodex is Y
      If is_aerodex_limited = True Then
        sQuery.Append(" and sqlrep_aerodex_flag = 'Y' ")
      End If


      sQuery.Append(" ORDER BY sqlrep_type, sqlrep_title")

      ' HttpContext.Current.Session.Item("localUser").crmUser_DebugText += "<br /><br /><b>getAdminReportListDataTable(ByVal nReportID As Integer, ByVal nSubID As Long, Optional ByVal sRptType As String = "", Optional ByVal bIsAdminReport As Boolean = False) As DataTable</b><br />" + sQuery.ToString

      SqlConn.ConnectionString = JETNET_LIVE_SQL_CONN

      SqlConn.Open()
      SqlCommand.Connection = SqlConn
      SqlCommand.CommandType = CommandType.Text
      SqlCommand.CommandTimeout = 240

      SqlCommand.CommandText = sQuery.ToString

      SqlReader = SqlCommand.ExecuteReader(CommandBehavior.CloseConnection)

      Try
        atemptable.Load(SqlReader)
      Catch constrExc As System.Data.ConstraintException
        Dim rowsErr As System.Data.DataRow() = atemptable.GetErrors()
        HttpContext.Current.Session.Item("localUser").crmUser_DebugText += "<br /><br /><b>Error in getAdminReportListDataTable load datatable</b><br /> " + constrExc.Message
      End Try

    Catch ex As Exception
      Return Nothing

      HttpContext.Current.Session.Item("localUser").crmUser_DebugText += "<br /><br /><b>Error in getAdminReportListDataTable() As DataTable</b><br />" + ex.Message

    Finally
      SqlReader = Nothing

      SqlConn.Dispose()
      SqlConn.Close()
      SqlConn = Nothing

      SqlCommand.Dispose()
      SqlCommand = Nothing
    End Try

    Return atemptable

  End Function
  Public Shared Function GenerateFileName(ByVal s_filename As String, ByVal s_filetype As String, ByVal b_replace_filetype As Boolean) As String

    Dim s_seperator As String = "_"
    Dim s_tmpFileName As String = ""

    Dim s_day As String = ""
    Dim s_month As String = ""
    Dim s_year As String = ""
    Dim n_hour As Integer = 0
    Dim s_minute As String = ""
    Dim s_second As String = ""
    Dim s_msecond As String = ""
    Dim s_ampm As String = ""

    If Not b_replace_filetype Then

      s_day = Now().Day.ToString
      s_month = Now().Month.ToString
      s_year = Now().Year.ToString
      n_hour = CInt(Now().Hour.ToString)
      s_minute = Now().Minute.ToString
      s_second = Now().Second.ToString
      s_msecond = Now().Millisecond.ToString

      Select Case n_hour

        Case 0
          s_ampm = "AM"
          n_hour = 12
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
          s_ampm = "AM"
        Case 12
          s_ampm = "PM"
        Case 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23
          s_ampm = "PM"
          n_hour = n_hour - 12

      End Select

      s_tmpFileName = s_month + s_seperator + s_day + s_seperator + s_year + _
                    s_seperator + n_hour.ToString + s_seperator + s_minute + s_seperator + s_second + s_seperator + s_msecond + _
                    s_seperator + s_ampm

      If Not String.IsNullOrEmpty(s_filename) Then
        s_tmpFileName = s_filename + s_seperator + s_tmpFileName
      End If

      If Not String.IsNullOrEmpty(s_filetype) Then
        s_tmpFileName += s_filetype
      End If

    Else

      Dim pos As Integer = s_filename.IndexOf(".")

      If pos > 0 Then
        'strip off old extension and put new one on
        s_tmpFileName = s_filename.Remove(pos, (s_filename.Length - pos))
        s_tmpFileName += s_filetype
      Else
        s_tmpFileName = s_filename + s_filetype
      End If

    End If

    Return s_tmpFileName

  End Function

  Private Function write_report_string_to_file(ByVal sOutoutString As String, ByVal sReportname As String) As Boolean

    Try

      Dim f As System.IO.StreamWriter

      f = System.IO.File.CreateText(HttpContext.Current.Server.MapPath("pictures\" + sReportname.Trim))   'HttpContext.Current.Session.Item("MarketSummaryFolderVirtualPath").ToString) + 

      ' write to the file
      f.WriteLine(sOutoutString)

      'close the streamwriter
      f.Close()
      f.Dispose()
      f = Nothing

    Catch ex As Exception
      HttpContext.Current.Session.Item("localUser").crmUser_DebugText += "Error in write_report_string_to_file(ByVal sOutoutString As String, ByVal sReportname As String) As Boolean " + ex.Message
      Return False
    End Try

    Return True

  End Function

#End Region


  Public Function scrape_for_Euro_Control()
    Dim Str As System.IO.Stream
    Dim srRead As System.IO.StreamReader


    Dim string_text As String = ""
    Dim string_text2 As String = ""
    Dim i As Integer = 0
    Dim final_string As String = ""
    Dim original_string_text As String = ""
    Dim article_link As String = ""
    Dim spot_to_find As Integer = 0
    Dim spot_to_find2 As Integer = 0
    Dim array_split() As String
    Dim k As Integer = 0
    Dim skip_this As Boolean = False
    Dim extra_note As String = ""

    Dim temp_ac_name As String = ""
    Dim temp_engine As String = ""
    Dim temp_eng As String = ""
    Dim temp_av As String = ""
    Dim temp_ac_id As Long = 0
    Dim temp_make As String = ""
    Dim temp_temp As String
    Dim temp_model As String = ""
    Dim temp_year As String = ""
    Dim array_split_make() As String
    Dim tcount As Integer = 0




    Try


      System.Threading.Thread.Sleep(10)
      Response.Flush()
      System.Threading.Thread.Sleep(10)

      Dim req As System.Net.WebRequest

      'req = System.Net.WebRequest.Create("http://www.eurocontrol.int/rmalive/operatorList.do?onclickAction=refreshIcaoCode&action=search&d-49520-p=2&operatorIcaoCode=VJT&aircraftIcaoType=&operatorName=VISTA+JET+LTD")

      '  req = System.Net.WebRequest.Create("http://www.eurocontrol.int/rmalive/operatorList.do?onclickAction=refreshIcaoCode&action=search&d-49520-p=2&operatorIcaoCode=VJT&aircraftIcaoType=&operatorName=VISTA+JET+LTD")
      '  req = System.Net.WebRequest.Create("http://www.eurocontrol.int/rmalive/operatorList.do?onclickAction=refreshIcaoCode&action=search&d-49520-p=2&operatorIcaoCode=VJT&aircraftIcaoType=&operatorName=VISTA+JET+LTD")
      '  req = System.Net.WebRequest.Create("http://www.eurocontrol.int/rmalive/operatorList.do?onclickAction=refreshIcaoCode&action=search&d-49520-p=2&operatorIcaoCode=VJT&aircraftIcaoType=&operatorName=VISTA+JET+LTD")
      req = System.Net.WebRequest.Create("http://www.eurocontrol.int/rmalive/regulatorList.do?&operatorState=GERMANY")


      Dim resp As System.Net.WebResponse = req.GetResponse


      Str = resp.GetResponseStream
      srRead = New System.IO.StreamReader(Str)
      ' read all the text 
      string_text = srRead.ReadToEnd().ToString

      resp.Close()
      resp = Nothing
      req = Nothing

      string_text = string_text
      original_string_text = string_text


      spot_to_find = InStr(string_text, "ninetyPctWidth")
      If spot_to_find > 0 Then
        string_text = Right(string_text, Len(string_text) - spot_to_find + 5)

        array_split = Split(string_text, "AdId"":")

        For i = 2 To array_split.Length - 1
          tcount = tcount + 1  ' do every third one 
          string_text = array_split(i)
          original_string_text = string_text

          acpub_count = acpub_count + 1

          temp_ac_name = ""
          temp_engine = ""
          temp_eng = ""
          temp_av = ""

          pub_reg_no = ""
          pub_ser_no = ""
          pub_desc = ""
          pub_price = ""
          pub_aftt = ""
          pub_seller_info = ""
          pub_picture = ""
          pub_status = ""
          pub_url = ""
          has_pics = False
          aftt_different = ""
          temp_year = ""

          spot_to_find = InStr(string_text, ",")
          If spot_to_find > 0 Then
            pub_url = Left(Trim(string_text), spot_to_find - 1)

            pub_url = "https://www.globalair.com/aircraft-for-sale/ListingDetail/TMAKE-TMODEL?AdId=" & pub_url
          End If

          spot_to_find = InStr(string_text, "PlaneTitle")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 10)
          End If


          spot_to_find = InStr(string_text, ",")
          If spot_to_find > 0 Then
            temp_ac_name = Left(Trim(string_text), spot_to_find - 1)

            spot_to_find = InStr(temp_ac_name, ":")
            If spot_to_find > 0 Then
              temp_ac_name = Right(temp_ac_name, Len(temp_ac_name) - spot_to_find - 1)
            End If

            temp_ac_name = Replace(Trim(temp_ac_name), """", "")
          End If

          cutme(temp_ac_name)



          spot_to_find = InStr(string_text, "TotalTime")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 9)

            spot_to_find = InStr(string_text, ",")
            If spot_to_find > 0 Then
              pub_aftt = Left(Trim(string_text), spot_to_find - 1)
            End If
            pub_aftt = Replace(pub_aftt, """", "")
            pub_aftt = Replace(pub_aftt, ":", "")
            cutme(pub_aftt)
          End If


          spot_to_find = InStr(string_text, "Price")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 6)

            spot_to_find = InStr(string_text, ",")
            If spot_to_find > 0 Then
              pub_price = Left(Trim(string_text), spot_to_find - 1)
            End If
            pub_price = Replace(pub_price, """", "")
            pub_price = Replace(pub_price, ":", "")
            cutme(pub_price)
          End If



          spot_to_find = InStr(string_text, "BrokerName")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 10)

            spot_to_find = InStr(string_text, ",")
            If spot_to_find > 0 Then
              pub_seller_info = Left(Trim(string_text), spot_to_find - 1)
            End If
            pub_seller_info = Replace(pub_seller_info, """", "")
            pub_seller_info = Replace(pub_seller_info, ":", "")
            cutme(pub_seller_info)
          End If


          spot_to_find = InStr(string_text, "SerialNum")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 9)

            spot_to_find = InStr(string_text, ",")
            If spot_to_find > 0 Then
              pub_ser_no = Left(Trim(string_text), spot_to_find - 1)
            End If
            pub_ser_no = Replace(pub_ser_no, """", "")
            pub_ser_no = Replace(pub_ser_no, ":", "")
            cutme(pub_ser_no)
          End If


          spot_to_find = InStr(string_text, "RegNum")
          If spot_to_find > 0 Then
            string_text = Right(string_text, Len(string_text) - spot_to_find - 6)

            spot_to_find = InStr(string_text, ",")
            If spot_to_find > 0 Then
              pub_reg_no = Left(Trim(string_text), spot_to_find - 1)
            End If
            pub_reg_no = Replace(pub_reg_no, """", "")
            pub_reg_no = Replace(pub_reg_no, ":", "")
            cutme(pub_reg_no)
          End If





          If IsNumeric(Left(Trim(temp_ac_name), "4")) Then
            temp_year = Left(Trim(temp_ac_name), "4")
          Else
            temp_year = ""
          End If

          acpub_original_name = temp_ac_name & " " & pub_ser_no

          Response.Write("<Br>")
          Response.Write("<Br>" & temp_ac_name)
          Response.Write("<Br>" & pub_price)
          Response.Write("<Br>" & pub_url)
          Response.Write("<Br>" & pub_seller_info)
          'Response.Write("<Br>" & pub_desc)
          Response.Write("<Br>" & temp_year)
          Response.Write("<Br>" & pub_ser_no)
          Response.Write("<Br>" & pub_reg_no)
          Response.Write("<Br>" & pub_aftt)

          If Trim(temp_ac_name) <> "" Then

            array_split_make = Split(Trim(temp_ac_name), " ")

            If array_split_make.Length = 2 Then
              temp_make = array_split_make(0)
              temp_model = array_split_make(1)
            ElseIf array_split_make.Length = 3 Then
              temp_make = array_split_make(1)
              temp_model = array_split_make(2)
            ElseIf array_split_make.Length = 4 Then
              temp_make = array_split_make(2)
              temp_model = array_split_make(3)
            ElseIf array_split_make.Length = 5 Then
              temp_make = array_split_make(3)
              temp_model = array_split_make(4)
            Else
              temp_temp = ""
            End If



            pub_url = Replace(Trim(pub_url), "TMAKE", Trim(temp_make))
            pub_url = Replace(Trim(pub_url), "TMODEL", Trim(temp_model))


            temp_ac_id = find_ac_global_search(pub_ser_no, temp_make, temp_model, "")
            If temp_ac_id = 0 Then
              temp_ac_id = find_ac_global_search(pub_ser_no, "", temp_model, "")
              If temp_ac_id = 0 Then
                temp_ac_id = find_ac_global_search(pub_ser_no, "", "", pub_reg_no)
                If temp_ac_id = 0 Then

                  temp_ac_id = find_ac_global_search(pub_ser_no, "", "", "")
                  If temp_ac_id = 0 Then
                    temp_ac_id = find_ac_global_search("", "", "", pub_reg_no)
                    If temp_ac_id = 0 Then
                      temp_ac_id = temp_ac_id
                    End If
                  End If

                End If
              End If

            End If

            acpub_price_details = ""
                        If On_Naughty_List(temp_ac_name) = True Then
                            ' if its on naughtly list then exclude 
                            temp_ac_id = temp_ac_id
                        Else
                            If temp_ac_id > 0 Then
                Call find_ac_data(temp_ac_id)
              Else
                acpub_process_status = "For Sale Not Found – No AC Match"
                acpub_status = "O"
              End If

              If Trim(aftt_different) <> "" Then
                pub_desc = pub_desc & aftt_different
              End If


              If Trim(acpub_price_details) <> "" Then
                If Trim(aftt_different) <> "" Then
                  pub_desc = pub_desc & ", "
                End If
                pub_desc = pub_desc & " " & acpub_price_details
              End If


              Call check_insert_ac_pub(temp_ac_id, 8)
            End If

            Response.Write("<Br>AC ID:" & temp_ac_id)
          End If


          'If tcount = 2 Then
          '  tcount = 0
          'End If

        Next


      End If




    Catch ex As Exception
    Finally

    End Try

  End Function
End Class