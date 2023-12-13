﻿Option Explicit On
Public PrevLine As Variant, CurrentLine As Variant, NextLine As Variant, NextNextLine As Variant
Public RunDate As Variant, TailNo As Variant, Current_AFHR As Variant, Venue As Variant
Public RSActivity As Variant, RSDeviation As Variant, ArrRSInput As Variant, i As Variant, j As Variant
Public RSCoincident As Variant, RSCFU As Variant, RSInProgress As Variant, RSOverdue As Variant, RSSMR As Variant
Public Tail_SerialNo As Variant, PART As Variant, CAGE As Variant, LCN As Variant, ALC As Variant
Public Position_Name As Variant, WAC As Variant, Life_Since_New As Variant, Life_Since_New_Type As Variant
Public RSEventDriven As Variant, RSDateDriven As Variant, RSHourDriven As Variant, RSDeviationRpt As Variant, ArrActivity As Variant
Public RSCFURpt As Variant, ArrCFU As Variant, RSSMRRpt As Variant, ArrSMR As Variant, RSHeaderInfo As Variant
Public MFTFolderName As String, FoundDate As String, FoundHour As String, InitParent As String, InitType As String
Public i_row As Integer, i_Coincident As Integer, i_InProgress As Integer, i_Overdue As Integer, i_Temp As Integer
Public ArrCoincident As Variant, ArrOverdue As Variant, ArrInProgress As Variant, ArrDeviation As Variant, ArrTemp As Variant
Public strRunDate As String, Duplicate As Boolean
Public steps As Integer, LoadBarStep As Integer
Public RST As ADODB.Recordset

Private Sub Btn_CFU_Click()

    If Not Me.cbTail.Value = "*" Then
        DoCmd.openreport "Rpt_CFU", acViewPreview, , "[Tail No]=" & "'" & Me.cbTail.Value & "'"
    ElseIf Not Me.cbVenue.Value = "*" Then
        DoCmd.openreport "Rpt_CFU", acViewPreview, , "[Venue]=" & "'" & Me.cbVenue.Value & "'"
    Else
        DoCmd.openreport "Rpt_CFU", acViewPreview
    End If

End Sub

Private Sub Btn_Date_Driven_Click()

    If Not Me.cbTail.Value = "*" Then
        DoCmd.openreport "Rpt_Date_Driven", acViewPreview, , "[Tail No]=" & "'" & Me.cbTail.Value & "'"
    ElseIf Not Me.cbVenue.Value = "*" Then
        DoCmd.openreport "Rpt_Date_Driven", acViewPreview, , "[Venue]=" & "'" & Me.cbVenue.Value & "'"
    Else
        DoCmd.openreport "Rpt_Date_Driven", acViewPreview
    End If

End Sub

Private Sub Btn_Deviation_Click()

    If Not Me.cbTail.Value = "*" Then
        DoCmd.openreport "Rpt_Deviation", acViewPreview, , "[Tail No]=" & "'" & Me.cbTail.Value & "'"
    ElseIf Not Me.cbVenue.Value = "*" Then
        DoCmd.openreport "Rpt_Deviation", acViewPreview, , "[Venue]=" & "'" & Me.cbVenue.Value & "'"
    Else
        DoCmd.openreport "Rpt_Deviation", acViewPreview
    End If

End Sub

Private Sub Btn_Event_Driven_Click()

    If Not Me.cbTail.Value = "*" Then
        DoCmd.openreport "Rpt_Event_Driven", acViewPreview, , "[Tail No]=" & "'" & Me.cbTail.Value & "'"
    ElseIf Not Me.cbVenue.Value = "*" Then
        DoCmd.openreport "Rpt_Event_Driven", acViewPreview, , "[Venue]=" & "'" & Me.cbVenue.Value & "'"
    Else
        DoCmd.openreport "Rpt_Event_Driven", acViewPreview
    End If

End Sub

Private Sub Btn_Hour_Driven_Click()

    If Not Me.cbTail.Value = "*" Then
        DoCmd.openreport "Rpt_Hour_Driven", acViewPreview, , "[Tail No]=" & "'" & Me.cbTail.Value & "'"
    ElseIf Not Me.cbVenue.Value = "*" Then
        DoCmd.openreport "Rpt_Hour_Driven", acViewPreview, , "[Venue]=" & "'" & Me.cbVenue.Value & "'"
    Else
        DoCmd.openreport "Rpt_Hour_Driven", acViewPreview
    End If

End Sub

Private Sub Btn_SMR_Click()

    If Not Me.cbTail.Value = "*" Then
        DoCmd.openreport "Rpt_SMR", acViewPreview, , "[Tail No]=" & "'" & Me.cbTail.Value & "'"
    ElseIf Not Me.cbVenue.Value = "*" Then
        DoCmd.openreport "Rpt_SMR", acViewPreview, , "[Venue]=" & "'" & Me.cbVenue.Value & "'"
    Else
        DoCmd.openreport "Rpt_SMR", acViewPreview
    End If

End Sub

Private Sub Btn_Summary_Click()

    If Not Me.cbTail.Value = "*" Then
        DoCmd.openreport "Rpt_Summary", acViewPreview, , "[Tail No]=" & "'" & Me.cbTail.Value & "'"
    ElseIf Not Me.cbVenue.Value = "*" Then
        DoCmd.openreport "Rpt_Summary", acViewPreview, , "[Venue]=" & "'" & Me.cbVenue.Value & "'"
    Else
        DoCmd.openreport "Rpt_Summary", acViewPreview
    End If

End Sub

Private Sub cbTail_GotFocus()

    If Me.cbVenue.Value = "*" Then
        Me.cbTail.RowSource = "SELECT DISTINCT tbl_MFT_Activity.[TAIL NO] FROM tbl_MFT_Activity UNION SELECT tbl_MFT_ComboValues.ComboValues FROM tbl_MFT_ComboValues ORDER BY tbl_MFT_Activity.[Tail No];"
    Else
        Me.cbTail.RowSource = "SELECT DISTINCT tbl_MFT_Activity.[TAIL NO] FROM tbl_MFT_Activity WHERE tbl_MFT_Activity.[VENUE] =  '" & Me.cbVenue.Value & "' UNION SELECT tbl_MFT_ComboValues.ComboValues FROM tbl_MFT_ComboValues;"
    End If

End Sub

Private Sub cbVenue_Change()

    ApplyMFTFilter

End Sub

Private Sub cbTail_Change()

    Dim CurrentAFHRS As Variant
    Dim CAMMDate As Variant

    ApplyMFTFilter

    If Not Me.cbTail = "*" Then
        CurrentAFHRS = DLookup("Current_AFHR", "tbl_MFT_Activity", "[Tail No]='" & Me.cbTail & "'")
        Me.txtAFHRS = CurrentAFHRS
    Else
        Me.txtAFHRS = ""
    End If

    If Not Me.cbTail = "*" Then
        CAMMDate = DLookup("RunDate", "tbl_MFT_Activity", "[Tail No]='" & Me.cbTail & "'")
        CAMMDate = Format(CAMMDate, "dd mmm yyyy")
        Me.txtCAMM = CAMMDate
    Else
        Me.txtCAMM = ""
    End If

End Sub

Private Sub Command131_Click()

    MFTFolderName = GetFolderName("")
    MFTFolderName = MFTFolderName & "\"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "tbl_MFT_Hour_Driven", MFTFolderName & "MFT Exported Data", True, "Hour Driven"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "tbl_MFT_Date_Driven", MFTFolderName & "MFT Exported Data", True, "Date Driven"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "tbl_MFT_Event_Driven", MFTFolderName & "MFT Exported Data", True, "Event Driven"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "tbl_MFT_Deviation_Rpt", MFTFolderName & "MFT Exported Data", True, "Deviations"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "tbl_MFT_CFU_Rpt", MFTFolderName & "MFT Exported Data", True, "CFU"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "tbl_MFT_SMR_Rpt", MFTFolderName & "MFT Exported Data", True, "SMR"
    MsgBox "Export Complete"

End Sub

Private Sub Command146_Click()

    DoCmd.openform("frm_MFT_Stagger_Chart_Hour")

End Sub

Private Sub Command147_Click()

    DoCmd.openform("frm_MFT_Stagger_Chart_Date")

End Sub

Private Sub Command148_Click()

    DoCmd.openform("frm_MFT_Stagger_Chart_Header_Info")

End Sub

Private Sub Form_Load()

    DoCmd.Maximize
    '    DoCmd.GoToControl "SubFrm_MFT"

End Sub

Sub Command7_Click()

    Dim CountTxt As Integer
    Dim FlagFirst As Boolean
    Dim FlagExists As Boolean
    Dim k As Integer
    Dim MFTFileName As Variant
    Dim RSInput As Variant
    Dim ReportVenue As Variant
    Dim MaintenanceType As String

    Me.Refresh

    MFTFolderName = GetFolderName("")

    CountTxt = CountFiles("txt")

    Dim strFile As String
    Dim ArrTailNo()

    strFile = Dir(MFTFolderName & "\*.txt")

    If Not strFile = "" And Not CountTxt = 0 Then

        DoCmd.SetWarnings False
        DoCmd.openform("Vixen Update Msgbox")
        Forms![Vixen Update Msgbox].Repaint
        
        'Loading bar info
        steps = CountTxt + 9
        LoadBarStatus steps, 1

        FlagFirst = False
        LoadBarStep = 2

        DoCmd.RunSQL "DELETE * FROM tbl_CurrentMFTTails"
        DoCmd.OpenQuery "CurrentMFTTails"
                
        Set RST = CurrentProject.Connection.Execute("select * from tbl_CurrentMFTTails")
        If RST.EOF = False Then
            ArrTailNo = RST.GetRows
        Else
            ReDim ArrTailNo(1, 0)
        End If
        If UBound(ArrTailNo, 2) = 0 Then
            k = 0
        Else
            k = UBound(ArrTailNo, 2) + 1
        End If

        Do While Len(strFile) > 0
            MFTFileName = MFTFolderName & "\" & strFile

            If MFTFileName <> "" Then
                DoCmd.RunSQL "DELETE * FROM MFTInput"

                DoCmd.TransferText _
                    TransferType:=acImportFixed,
                    SpecificationName:="MFTInput",
                    TableName:="MFTInput",
                    filename:=MFTFileName
            End If

            LoadBarStatus steps, LoadBarStep
            LoadBarStep = LoadBarStep + 1
            
            Set RSInput = CurrentDb.OpenRecordset("SELECT * FROM MFTInput")
            RSInput.MoveLast
            RSInput.MoveFirst
            ArrRSInput = RSInput.GetRows(RSInput.RecordCount)

            If FlagFirst = False Then
                'Setup Recordsets
                Set RSHeaderInfo = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_Header_Info")
                
                Set RSActivity = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_Activity")
                Set RSDeviation = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_Deviation")
                Set RSCoincident = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_Coincident")
                Set RSCFU = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_CFU")
                Set RSInProgress = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_In_Progress")
                Set RSOverdue = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_Overdue")
                Set RSSMR = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_SMR")
            
                FlagFirst = True
            End If

            For i = 0 To UBound(ArrRSInput, 2)
                'Get previous line data
                If i > 0 Then
                    PrevLine = ArrRSInput(0, i - 1)
                End If
                'Get current line data
                CurrentLine = ArrRSInput(0, i)
                'Get next line data
                If i < UBound(ArrRSInput, 2) Then
                    NextLine = ArrRSInput(0, i + 1)
                End If
                'Get current+2 line data
                If i < UBound(ArrRSInput, 2) - 1 Then
                    NextNextLine = ArrRSInput(0, i + 2)
                End If

                'Run Date
                If Left(CurrentLine, 16) Like "*STARTED DATE*" Then
                    RunDate = Mid(CurrentLine, InStr(CurrentLine, "=") + 2, 8)
                End If
                'Venue
                If Left(CurrentLine, 10) Like "Q19 Venue:" Or Left(CurrentLine, 10) Like "Q22 Venue:" Then
                    ReportVenue = Trim(Mid(CurrentLine, 17, 100))
                    ReportVenue = Mid(ReportVenue, InStr(ReportVenue, ":") + 1, 8)
                    CopyFile MFTFileName, PathBE & "CAMM2\" & ReportVenue & ".txt", False
                End If
                If Left(CurrentLine, 16) Like "*VENUE         :*" Then
                    Venue = Trim(Mid(CurrentLine, 17, 50))
                End If

                'Serial Number/Tail Number
                If InStr(Mid(CurrentLine, 79, 15), "SERIAL NUMBER:") Then
                    FlagExists = False
                    TailNo = Trim(Mid(CurrentLine, 94, 8))
                    For j = 0 To UBound(ArrTailNo, 2)
                        If ArrTailNo(0, j) = TailNo Then
                            If RunDate >= ArrTailNo(1, j) Then
                                DoCmd.RunSQL "DELETE * FROM tbl_MFT_Header_Info WHERE [Tail No]='" & TailNo & "';"
                                DoCmd.RunSQL "DELETE * FROM tbl_MFT_Activity WHERE [Tail No]='" & TailNo & "';"
                                DoCmd.RunSQL "DELETE * FROM tbl_MFT_Deviation WHERE [Tail No]='" & TailNo & "';"
                                DoCmd.RunSQL "DELETE * FROM tbl_MFT_Coincident WHERE [Tail No]='" & TailNo & "';"
                                DoCmd.RunSQL "DELETE * FROM tbl_MFT_CFU WHERE [Tail No]='" & TailNo & "';"
                                DoCmd.RunSQL "DELETE * FROM tbl_MFT_In_Progress WHERE [Tail No]='" & TailNo & "';"
                                DoCmd.RunSQL "DELETE * FROM tbl_MFT_Overdue WHERE [Tail No]='" & TailNo & "';"
                                DoCmd.RunSQL "DELETE * FROM tbl_MFT_SMR WHERE [Tail No]='" & TailNo & "';"
                                FlagExists = True
                                Exit For
                            End If
                        End If
                    Next
                    If FlagExists = False Then
                        ReDim Preserve ArrTailNo(1, k)
                        ArrTailNo(0, k) = TailNo
                        ArrTailNo(1, k) = RunDate
                        k = k + 1
                    End If
                End If

                'Current AFHRS for Serial Number/Tail Number and Header Info
                j = 1
                If InStr(Left(CurrentLine, 26), "LIFE SINCE NEW:    LIFE") Then
                    Do
                        Current_AFHR = Trim(Mid(NextLine, 16, 12))
                        GetHeaderInfo()
                        j = j + 1
                    Loop Until IsNull(Mid(ArrRSInput(0, i + j), 16, 12))
                End If

                'Check current Maintenance Type
                If Left(CurrentLine, 25) = "-- FORECAST ACTIVITIES --" Then
                    MaintenanceType = "Activities"
                ElseIf Left(CurrentLine, 16) = "-- DEVIATIONS --" Then
                    MaintenanceType = "Deviations"
                ElseIf InStr(CurrentLine, "-- COINCIDENT MAINTENANCE ACTIVITIES --") Then
                    MaintenanceType = "Coincident"
                ElseIf InStr(CurrentLine, "-- CFUs --") Then
                    MaintenanceType = "CFU"
                ElseIf InStr(CurrentLine, "-- JOBS IN PROGRESS --") Then
                    MaintenanceType = "In Progress"
                ElseIf InStr(CurrentLine, "-- OVERDUE MAINTENANCE --") Then
                    MaintenanceType = "Overdue"
                ElseIf InStr(CurrentLine, "-- SMRs --") Then
                    MaintenanceType = "SMR"
                End If

                'Get all Maint Activities
                If MaintenanceType = "Activities" Then
                    If Trim(Mid(CurrentLine, 64, 5)) = "MOD" Or Trim(Mid(CurrentLine, 64, 5)) = "STI" Or Trim(Mid(CurrentLine, 64, 5)) = "SM" Or
                    Trim(Mid(CurrentLine, 64, 5)) = "DVI" Then
                        GetActivityData()
                    End If
                End If
                'Get Deviation data
                If MaintenanceType = "Deviations" Then
                    If InStr(Trim(Mid(CurrentLine, 64, 12)), "DEV-") Then
                        GetDeviationData()
                    ElseIf Mid(CurrentLine, 98, 1) = " " And Trim(Mid(CurrentLine, 99, 1)) = "Y" And Mid(CurrentLine, 100, 1) = " " Then
                        GetDeviationData()
                    ElseIf Mid(CurrentLine, 98, 1) = " " And Trim(Mid(CurrentLine, 99, 1)) = "N" And Mid(CurrentLine, 100, 1) = " " Then
                        GetDeviationData()
                    End If
                End If
                'Get Coincident data
                If MaintenanceType = "Coincident" Then
                    If Trim(Mid(CurrentLine, 64, 5)) = "MOD" Or Trim(Mid(CurrentLine, 64, 5)) = "STI" Or Trim(Mid(CurrentLine, 64, 5)) = "SM" Or
                    Trim(Mid(CurrentLine, 64, 5)) = "DVI" Then
                        GetCoincidentData()
                    End If
                End If
                'Get CFU data
                If MaintenanceType = "CFU" Then
                    If Trim(Mid(CurrentLine, 64, 5)) = "CFU" Then
                        GetCFUData()
                    End If
                End If
                'Get SMR data
                If MaintenanceType = "SMR" Then
                    If Trim(Mid(CurrentLine, 64, 5)) = "SMR" Then
                        GetSMRData()
                    End If
                End If
                'Get In Progress data
                If MaintenanceType = "In Progress" Then
                    If Trim(Mid(CurrentLine, 73, 3)) = "Y" Or Trim(Mid(CurrentLine, 73, 3)) = "N" Then
                        GetInProgressData()
                    End If
                End If
                'Get Overdue Data
                If MaintenanceType = "Overdue" Then
                    If Trim(Mid(CurrentLine, 73, 3)) = "Y" Or Trim(Mid(CurrentLine, 73, 3)) = "N" Then
                        GetOverdueData()
                    End If
                End If
            Next

            If ReportVenue = "" Then
                CopyFile MFTFileName, PathBE & "CAMM2\" & "No Venue.txt", False
            End If

            strFile = Dir()
            RSInput.Close
            Set RSInput = Nothing
        Loop

        SortDataForReporting()

        LoadBarStatus steps, LoadBarStep

        ClearMemory()

        Forms![frm_MFT].Form.Requery
        Forms![frm_MFT].Form.Refresh

        DoCmd.Close acForm, "Vixen Update Msgbox", acSaveNo '*
        DoCmd.SetWarnings True '*
        MsgBox "Update Complete.", vbOKOnly, "Complete"
    End If

End Sub

Sub ClearMemory()

    RST.Close
    RSActivity.Close
    RSDeviation.Close
    RSCoincident.Close
    RSCFU.Close
    RSInProgress.Close
    RSOverdue.Close
    RSSMR.Close

    DoCmd.RunSQL "DELETE * FROM MFTInput;"

End Sub

Sub GetHeaderInfo()

    Dim HeaderTailNo As Variant
    Dim HeaderVenue As Variant
    Dim HeaderLife As Variant
    Dim HeaderLifeType As Variant
    Dim HeaderInterval As Variant

    HeaderTailNo = TailNo
    HeaderVenue = Venue
    HeaderLife = Trim(Mid(ArrRSInput(0, i + j), 16, 12))
    HeaderLifeType = Trim(Mid(ArrRSInput(0, i + j), 32, 12))
    HeaderInterval = Trim(Mid(ArrRSInput(0, i + j), 48, 12))
    If HeaderInterval = "" Then : HeaderInterval = 0

        RSHeaderInfo.AddNew
        RSHeaderInfo![Tail No] = HeaderTailNo
        RSHeaderInfo![Venue] = HeaderVenue
        RSHeaderInfo![Life] = HeaderLife
        RSHeaderInfo![Life Type] = HeaderLifeType
        RSHeaderInfo![INTERVAL] = HeaderInterval
        RSHeaderInfo.Update

End Sub

Sub GetActivityData()

    Dim Maint_Activity As Variant
    Dim Date_due As Variant
    Dim Life_Type_Int As Variant
    Dim Int_Rem As Variant
    Dim TSSCode1 As Variant
    Dim TSSCode2 As Variant
    Dim TSSCode3 As Variant
    Dim TSSCode4 As Variant
    Dim Maint_Type As Variant
    Dim LVL As Variant
    Dim STD As Variant
    Dim REC As Variant
    Dim CLM As Variant
    Dim Date_Plan As Variant
    Dim INTERVAL As Variant
    Dim Init_Type As Variant
    Dim INSTANCE As Variant
    Dim Init_LCN As Variant
    Dim Init_ALC As Variant
    Dim DateString As Variant

    'Get Full line of data from CAMM2 if it exists
    If Not Left(CurrentLine, 63) = "                                                               " Then
        Tail_SerialNo = Trim(Left(PrevLine, 23))
        PART = Trim(Mid(PrevLine, 24, 30))
        CAGE = Mid(Trim(Mid(PrevLine, 56, 7)), 2, Len(Trim(Mid(PrevLine, 56, 7))))
        LCN = Trim(Left(CurrentLine, 18))
        ALC = Trim(Mid(CurrentLine, 19, 3))
        Position_Name = Trim(Mid(CurrentLine, 24, 19))
        WAC = Trim(Mid(CurrentLine, 44, 5))
        Life_Since_New = Trim(Mid(NextLine, 44, 9))
        Life_Since_New_Type = Trim(Mid(NextLine, 54, 7))
    End If
    'Link remaining data to the previous full line of data
    Maint_Activity = Trim(Mid(PrevLine, 64, 21))
    Date_due = Trim(Mid(PrevLine, 86, 10))
    Life_Type_Int = Trim(Mid(PrevLine, 96, 10))
    Int_Rem = Trim(Mid(PrevLine, 105, 10))
    TSSCode1 = Trim(Mid(PrevLine, 121, 10))
    TSSCode2 = Trim(Mid(CurrentLine, 121, 10))
    TSSCode3 = Trim(Mid(NextLine, 121, 10))
    TSSCode4 = Trim(Mid(NextLine, 121, 10))
    Maint_Type = Trim(Mid(CurrentLine, 64, 4))
    LVL = Trim(Mid(CurrentLine, 69, 3))
    STD = Trim(Mid(CurrentLine, 74, 2))
    REC = Trim(Mid(CurrentLine, 78, 3))
    CLM = Trim(Mid(CurrentLine, 82, 3))
    Date_Plan = Trim(Mid(CurrentLine, 86, 10))
    INTERVAL = Trim(Mid(CurrentLine, 96, 10))
    Init_Type = Trim(Mid(NextLine, 64, 5))
    INSTANCE = Trim(Mid(NextLine, 74, 10))
    Init_LCN = Trim(Mid(NextLine, 96, 11))
    Init_ALC = Trim(Mid(NextLine, 114, 4))
    'Write Maint Activity data to table
    RSActivity.AddNew
    RSActivity![Tail No] = TailNo
    RSActivity![Current_AFHR] = Current_AFHR
    RSActivity![Venue] = Venue
    RSActivity![TAIL/SERIAL NO] = Tail_SerialNo
    RSActivity!PART = PART
    RSActivity!CAGE = CAGE
    RSActivity![MAINT ACTIVITY] = Maint_Activity
    If Not Date_due = "" Then
        DateString = ConvertToDateSerial(Date_due)
        RSActivity![DATE DUE] = DateString
    End If
    RSActivity![LIFE TYPE INT] = Life_Type_Int
    If Not Int_Rem = "" Then
        RSActivity!REM = Int_Rem
    End If
    RSActivity![TSS CODES 1] = TSSCode1
    RSActivity![TSS CODES 2] = TSSCode2
    RSActivity![TSS CODES 3] = TSSCode3
    RSActivity![TSS CODES 4] = TSSCode4
    RSActivity!LCN = LCN
    RSActivity!ALC = ALC
    RSActivity![POSITION NAME] = Position_Name
    RSActivity!WAC = WAC
    RSActivity!Type = Maint_Type
    RSActivity!LVL = LVL
    RSActivity!STD = STD
    RSActivity!REC = REC
    RSActivity!CLM = CLM
    RSActivity![DATE PLAN] = Date_Plan
    RSActivity!INTERVAL = INTERVAL
    RSActivity![LIFE SINCE New] = Life_Since_New
    RSActivity![LIFE SINCE New TYPE] = Life_Since_New_Type
    RSActivity![INIT TYPE] = Init_Type
    RSActivity!INSTANCE = INSTANCE
    RSActivity![INIT LCN] = Init_LCN
    RSActivity![INIT ALC] = Init_ALC
    RSActivity![RunDate] = RunDate
    RSActivity.Update

End Sub

Sub GetDeviationData()

    Dim Deviation_Order As Variant
    Dim STD As Variant
    Dim EXPIRED As Variant
    Dim TITLE As Variant
    Dim Expiry_Conditions As Variant
    Dim LIMITATION As Variant
    Dim ACTIVITY As Variant
    Dim Dev_Type As Variant
    Dim Init_LCN As Variant
    Dim Init_ALC As Variant
    Dim INTERVAL As Variant
    Dim Int_Rem As Variant
    Dim Date_due As Variant
    Dim DateString As Variant


    If Not Left(CurrentLine, 63) = "                                                               " Then
        Tail_SerialNo = Trim(Left(CurrentLine, 23))
        PART = Trim(Mid(CurrentLine, 24, 30))
        CAGE = Mid(Trim(Mid(CurrentLine, 56, 7)), 2, Len(Trim(Mid(CurrentLine, 56, 7))))
        LCN = Trim(Left(NextLine, 18))
        ALC = Trim(Mid(NextLine, 19, 3))
        Position_Name = Trim(Mid(NextLine, 24, 19))
        WAC = Trim(Mid(NextLine, 44, 5))
        Life_Since_New = Trim(Mid(NextNextLine, 44, 9))
        Life_Since_New_Type = Trim(Mid(NextNextLine, 54, 7))
    End If
    Deviation_Order = Trim(Mid(CurrentLine, 64, 30))
    STD = Trim(Mid(CurrentLine, 98, 5))
    EXPIRED = Trim(Mid(CurrentLine, 117, 5))
    TITLE = Trim(Mid(NextLine, 64, 50))
    If Trim(Mid(ArrRSInput(0, i + 3), 64, 31)) = "*** NO EXPIRY CONDITIONS ***" Then
        Expiry_Conditions = "NO EXPIRY CONDITIONS"
    Else
        Expiry_Conditions = ""
    End If
    'Collect all limitation data including data that crosses onto another page
    LIMITATION = ""
    If InStr(ArrRSInput(0, i + 7), "LIMITATION:") Then
        j = 7
        Do
            If Mid(ArrRSInput(0, i + j), 56, 12) = "For-Official" And Not Trim(Mid(ArrRSInput(0, (i + j) + 10), 64, 7)) = "(cont)" Then
                Exit Do
            End If
            If InStr(Trim(Mid(ArrRSInput(0, (i + j)), 64, 12)), "DEV-") Then
                Exit Do
            End If
            If Trim(Mid(ArrRSInput(0, (i + j) + 11), 64, 7)) = "(cont)" Then
                j = j + 11
            End If
            If Not Trim(Mid(ArrRSInput(0, i + j), 76, 55)) = "" Then
                LIMITATION = LIMITATION & Trim(Mid(ArrRSInput(0, i + j), 76, 55)) & vbNewLine
            End If
            If Left(ArrRSInput(0, i + j), 3) = "-- " Then
                Exit Do
            End If
            j = j + 1
        Loop
    End If
    'Get all the deviation order details; can include multiple for one deviation
    j = 5
    Do
        If Not Trim(Mid(ArrRSInput(0, i + 3), 64, 31)) = "*** NO EXPIRY CONDITIONS ***" Then
            Dev_Type = Trim(Mid(ArrRSInput(0, i + j), 64, 7))
            ACTIVITY = Trim(Mid(ArrRSInput(0, i + j), 72, 15))
            Init_LCN = Trim(Mid(ArrRSInput(0, i + j), 98, 2))
            If Len(Trim(Mid(ArrRSInput(0, i + j), 111, 9))) < 3 Then
                Init_ALC = Trim(Mid(ArrRSInput(0, i + j), 118, 2))
            Else
                Init_ALC = ""
            End If
            INTERVAL = Trim(Mid(ArrRSInput(0, i + j), 89, 8))
            Int_Rem = Trim(Mid(ArrRSInput(0, i + j), 103, 7))
            If Len(Trim(Mid(ArrRSInput(0, i + j), 111, 9))) > 2 Then
                Date_due = Trim(Mid(ArrRSInput(0, i + j), 111, 9))
            Else
                Date_due = ""
            End If
        Else
            Dev_Type = ""
            ACTIVITY = ""
            Init_LCN = ""
            Init_ALC = ""
            INTERVAL = ""
            Int_Rem = ""
            Date_due = ""
        End If
        'Write all deviation data to table
        RSDeviation.AddNew
        RSDeviation![Tail No] = TailNo
        RSDeviation![Current_AFHR] = Current_AFHR
        RSDeviation![Venue] = Venue
        RSDeviation![TAIL/SERIAL NO] = Tail_SerialNo
        RSDeviation!PART = PART
        RSDeviation!CAGE = CAGE
        RSDeviation![DEVIATION ORDER] = Deviation_Order
        RSDeviation!STD = STD
        RSDeviation!EXPIRED = EXPIRED
        RSDeviation!LCN = LCN
        RSDeviation!ALC = ALC
        RSDeviation![POSITION NAME] = Position_Name
        RSDeviation!WAC = WAC
        RSDeviation!TITLE = TITLE
        RSDeviation![LIFE SINCE New] = Life_Since_New
        RSDeviation![LIFE SINCE New TYPE] = Life_Since_New_Type
        RSDeviation![EXPIRY CONDITIONS] = Expiry_Conditions
        RSDeviation!LIMITATION = LIMITATION
        RSDeviation![Type] = Dev_Type
        RSDeviation![ACTIVITY] = ACTIVITY
        RSDeviation![INIT LCN] = Init_LCN
        RSDeviation![INIT ALC] = Init_ALC
        RSDeviation![INTERVAL] = INTERVAL
        If Not Int_Rem = "" And IsNumeric(Int_Rem) Then
            RSDeviation![REM] = Int_Rem
        End If
        If Not Date_due = "" Then
            DateString = ConvertToDateSerial(Date_due)
            RSDeviation![DATE DUE] = DateString
        End If
        RSDeviation![RunDate] = RunDate
        RSDeviation.Update

        'Exit loop if the next line is a blank space
        j = j + 1
        If IsNull(Trim(Mid(ArrRSInput(0, i + j), 64, 7))) Or Trim(Mid(ArrRSInput(0, i + 3), 64, 31)) = "*** NO EXPIRY CONDITIONS ***" Then
            Exit Do
        End If
    Loop

End Sub

Sub GetCoincidentData()

    Dim Maint_Activity As Variant
    Dim Date_due As Variant
    Dim Life_Type_Int As Variant
    Dim Int_Rem As Variant
    Dim TSSCode1 As Variant
    Dim TSSCode2 As Variant
    Dim TSSCode3 As Variant
    Dim TSSCode4 As Variant
    Dim Maint_Type As Variant
    Dim LVL As Variant
    Dim STD As Variant
    Dim REC As Variant
    Dim CLM As Variant
    Dim Date_Plan As Variant
    Dim INTERVAL As Variant
    Dim Init_Type As Variant
    Dim INSTANCE As Variant
    Dim Init_LCN As Variant
    Dim Init_ALC As Variant
    Dim DateString As Variant

    'Get Full line of data from CAMM2 if it exists
    If Not Left(CurrentLine, 63) = "                                                               " Then
        Tail_SerialNo = Trim(Left(PrevLine, 23))
        PART = Trim(Mid(PrevLine, 24, 30))
        CAGE = Mid(Trim(Mid(PrevLine, 56, 7)), 2, Len(Trim(Mid(PrevLine, 56, 7))))
        LCN = Trim(Left(CurrentLine, 18))
        ALC = Trim(Mid(CurrentLine, 19, 3))
        Position_Name = Trim(Mid(CurrentLine, 24, 19))
        WAC = Trim(Mid(CurrentLine, 44, 5))
        Life_Since_New = Trim(Mid(NextLine, 44, 9))
        Life_Since_New_Type = Trim(Mid(NextLine, 54, 7))
    End If
    'Link remaining data to the previous full line of data
    Maint_Activity = Trim(Mid(PrevLine, 64, 21))
    Date_due = Trim(Mid(PrevLine, 86, 10))
    Life_Type_Int = Trim(Mid(PrevLine, 96, 10))
    Int_Rem = Trim(Mid(PrevLine, 105, 10))
    TSSCode1 = Trim(Mid(PrevLine, 121, 10))
    TSSCode2 = Trim(Mid(CurrentLine, 121, 10))
    TSSCode3 = Trim(Mid(NextLine, 121, 10))
    TSSCode4 = Trim(Mid(NextLine, 121, 10))
    Maint_Type = Trim(Mid(CurrentLine, 64, 4))
    LVL = Trim(Mid(CurrentLine, 69, 3))
    STD = Trim(Mid(CurrentLine, 74, 2))
    REC = Trim(Mid(CurrentLine, 78, 3))
    CLM = Trim(Mid(CurrentLine, 82, 3))
    Date_Plan = Trim(Mid(CurrentLine, 86, 10))
    INTERVAL = Trim(Mid(CurrentLine, 96, 10))
    Init_Type = Trim(Mid(NextLine, 64, 5))
    INSTANCE = Trim(Mid(NextLine, 74, 17))
    Init_LCN = Trim(Mid(NextLine, 96, 11))
    Init_ALC = Trim(Mid(NextLine, 114, 4))
    'Write Maint Activity data to table
    RSCoincident.AddNew
    RSCoincident![Tail No] = TailNo
    RSCoincident![Current_AFHR] = Current_AFHR
    RSCoincident![Venue] = Venue
    RSCoincident![TAIL/SERIAL NO] = Tail_SerialNo
    RSCoincident!PART = PART
    RSCoincident!CAGE = CAGE
    RSCoincident![MAINT ACTIVITY] = Maint_Activity
    If Not Date_due = "" Then
        DateString = ConvertToDateSerial(Date_due)
        RSCoincident![DATE DUE] = DateString
    End If
    RSCoincident![LIFE TYPE INT] = Life_Type_Int
    If Not Int_Rem = "" Then
        RSCoincident!REM = Int_Rem
    End If
    RSCoincident![TSS CODES 1] = TSSCode1
    RSCoincident![TSS CODES 2] = TSSCode2
    RSCoincident![TSS CODES 3] = TSSCode3
    RSCoincident![TSS CODES 4] = TSSCode4
    RSCoincident!LCN = LCN
    RSCoincident!ALC = ALC
    RSCoincident![POSITION NAME] = Position_Name
    RSCoincident!WAC = WAC
    RSCoincident!Type = Maint_Type
    RSCoincident!LVL = LVL
    RSCoincident!STD = STD
    RSCoincident!REC = REC
    RSCoincident!CLM = CLM
    RSCoincident![DATE PLAN] = Date_Plan
    RSCoincident!INTERVAL = INTERVAL
    RSCoincident![LIFE SINCE New] = Life_Since_New
    RSCoincident![LIFE SINCE New TYPE] = Life_Since_New_Type
    RSCoincident![INIT TYPE] = Init_Type
    RSCoincident!INSTANCE = INSTANCE
    RSCoincident![INIT LCN] = Init_LCN
    RSCoincident![INIT ALC] = Init_ALC
    RSCoincident![RunDate] = RunDate
    RSCoincident.Update

End Sub

Sub GetCFUData()

    Dim Maint_Activity As Variant
    Dim Date_due As Variant
    Dim Life_Type_Int As Variant
    Dim Int_Rem As Variant
    Dim TSSCode1 As Variant
    Dim TSSCode2 As Variant
    Dim TSSCode3 As Variant
    Dim TSSCode4 As Variant
    Dim Maint_Type As Variant
    Dim LVL As Variant
    Dim STD As Variant
    Dim REC As Variant
    Dim CLM As Variant
    Dim Date_Plan As Variant
    Dim INTERVAL As Variant
    Dim Init_Type As Variant
    Dim INSTANCE As Variant
    Dim Init_LCN As Variant
    Dim Init_ALC As Variant
    Dim DateString As Variant
    Dim SYMPTOMS As Variant

    'Get Full line of data from CAMM2 if it exists
    If Not Left(CurrentLine, 63) = "                                                               " Then
        Tail_SerialNo = Trim(Left(PrevLine, 23))
        PART = Trim(Mid(PrevLine, 24, 30))
        CAGE = Mid(Trim(Mid(PrevLine, 56, 7)), 2, Len(Trim(Mid(PrevLine, 56, 7))))
        LCN = Trim(Left(CurrentLine, 18))
        ALC = Trim(Mid(CurrentLine, 19, 3))
        Position_Name = Trim(Mid(CurrentLine, 24, 19))
        WAC = Trim(Mid(CurrentLine, 44, 5))
        Life_Since_New = Trim(Mid(NextLine, 44, 9))
        Life_Since_New_Type = Trim(Mid(NextLine, 54, 7))
    End If
    'Link remaining data to the previous full line of data
    Maint_Activity = Trim(Mid(PrevLine, 64, 21))
    Date_due = Trim(Mid(PrevLine, 86, 10))
    Life_Type_Int = Trim(Mid(PrevLine, 96, 10))
    Int_Rem = Trim(Mid(PrevLine, 105, 10))
    Maint_Type = Trim(Mid(CurrentLine, 64, 4))
    LVL = Trim(Mid(CurrentLine, 69, 3))
    STD = Trim(Mid(CurrentLine, 74, 2))
    REC = Trim(Mid(CurrentLine, 78, 3))
    CLM = Trim(Mid(CurrentLine, 82, 3))
    Date_Plan = Trim(Mid(CurrentLine, 86, 10))
    INTERVAL = Trim(Mid(CurrentLine, 96, 10))
    Init_Type = Trim(Mid(NextLine, 64, 5))
    INSTANCE = Trim(Mid(NextLine, 74, 10))
    Init_LCN = Trim(Mid(NextLine, 96, 11))
    Init_ALC = Trim(Mid(NextLine, 114, 4))

    SYMPTOMS = ""
    If InStr(ArrRSInput(0, i + 2), "SYMPTOMS:") Or InStr(ArrRSInput(0, i + 3), "SYMPTOMS:") Then
        j = 2
        Do
            If Mid(ArrRSInput(0, i + j), 56, 12) = "For-Official" And Not Trim(Mid(ArrRSInput(0, (i + j) + 10), 64, 7)) = "(cont)" Then
                Exit Do
            End If
            If InStr(Trim(Mid(ArrRSInput(0, ((i + j) + 1)), 64, 9)), "CFU") Then
                Exit Do
            End If
            If ((i + j) + 11) <= UBound(ArrRSInput, 2) Then
                If Trim(Mid(ArrRSInput(0, (i + j) + 11), 64, 7)) = "(cont)" Then
                    j = j + 11
                End If
            Else
                Exit Do
            End If
            If Not Trim(Mid(ArrRSInput(0, i + j), 74, 57)) = "" Then
                SYMPTOMS = SYMPTOMS & Trim(Mid(ArrRSInput(0, i + j), 74, 57)) & vbNewLine
            End If
            If Left(ArrRSInput(0, i + j), 3) = "-- " Then
                Exit Do
            End If
            j = j + 1
        Loop
    End If

    'Write Maint Activity data to table
    RSCFU.AddNew
    RSCFU![Tail No] = TailNo
    RSCFU![Current_AFHR] = Current_AFHR
    RSCFU![Venue] = Venue
    RSCFU![TAIL/SERIAL NO] = Tail_SerialNo
    RSCFU!PART = PART
    RSCFU!CAGE = CAGE
    RSCFU![MAINT ACTIVITY] = Maint_Activity
    If Not Date_due = "" Then
        DateString = ConvertToDateSerial(Date_due)
        RSCFU![DATE DUE] = DateString
    End If
    RSCFU![LIFE TYPE INT] = Life_Type_Int
    If Not Int_Rem = "" And IsNumeric(Int_Rem) Then
        RSCFU!REM = Int_Rem
    End If
    RSCFU!LCN = LCN
    RSCFU!ALC = ALC
    RSCFU![POSITION NAME] = Position_Name
    RSCFU!WAC = WAC
    RSCFU!Type = Maint_Type
    RSCFU!LVL = LVL
    RSCFU!STD = STD
    RSCFU!REC = REC
    RSCFU!CLM = CLM
    RSCFU![DATE PLAN] = Date_Plan
    RSCFU!INTERVAL = INTERVAL
    RSCFU![LIFE SINCE New] = Life_Since_New
    RSCFU![LIFE SINCE New TYPE] = Life_Since_New_Type
    RSCFU![INIT TYPE] = Init_Type
    RSCFU!INSTANCE = INSTANCE
    RSCFU![INIT LCN] = Init_LCN
    RSCFU![INIT ALC] = Init_ALC
    RSCFU![SYMPTOMS] = SYMPTOMS
    RSCFU![RunDate] = RunDate
    RSCFU.Update

End Sub

Sub GetSMRData()

    Dim Maint_Activity As Variant
    Dim Date_due As Variant
    Dim Life_Type_Int As Variant
    Dim Int_Rem As Variant
    Dim TSSCode1 As Variant
    Dim TSSCode2 As Variant
    Dim TSSCode3 As Variant
    Dim TSSCode4 As Variant
    Dim Maint_Type As Variant
    Dim LVL As Variant
    Dim STD As Variant
    Dim REC As Variant
    Dim CLM As Variant
    Dim Date_Plan As Variant
    Dim INTERVAL As Variant
    Dim Init_Type As Variant
    Dim INSTANCE As Variant
    Dim Init_LCN As Variant
    Dim Init_ALC As Variant
    Dim DateString As Variant
    Dim SYMPTOMS As Variant

    'Get Full line of data from CAMM2 if it exists
    If Not Left(CurrentLine, 63) = "                                                               " Then
        Tail_SerialNo = Trim(Left(PrevLine, 23))
        PART = Trim(Mid(PrevLine, 24, 30))
        CAGE = Mid(Trim(Mid(PrevLine, 56, 7)), 2, Len(Trim(Mid(PrevLine, 56, 7))))
        LCN = Trim(Left(CurrentLine, 18))
        ALC = Trim(Mid(CurrentLine, 19, 3))
        Position_Name = Trim(Mid(CurrentLine, 24, 19))
        WAC = Trim(Mid(CurrentLine, 44, 5))
        Life_Since_New = Trim(Mid(NextLine, 44, 9))
        Life_Since_New_Type = Trim(Mid(NextLine, 54, 7))
    End If
    'Link remaining data to the previous full line of data
    Maint_Activity = Trim(Mid(PrevLine, 64, 21))
    Date_due = Trim(Mid(PrevLine, 86, 10))
    Life_Type_Int = Trim(Mid(PrevLine, 96, 10))
    Int_Rem = Trim(Mid(PrevLine, 105, 10))
    Maint_Type = Trim(Mid(CurrentLine, 64, 4))
    LVL = Trim(Mid(CurrentLine, 69, 3))
    STD = Trim(Mid(CurrentLine, 74, 2))
    REC = Trim(Mid(CurrentLine, 78, 3))
    CLM = Trim(Mid(CurrentLine, 82, 3))
    Date_Plan = Trim(Mid(CurrentLine, 86, 10))
    INTERVAL = Trim(Mid(CurrentLine, 96, 10))
    Init_Type = Trim(Mid(NextLine, 64, 5))
    INSTANCE = Trim(Mid(NextLine, 74, 10))
    Init_LCN = Trim(Mid(NextLine, 96, 11))
    Init_ALC = Trim(Mid(NextLine, 114, 4))

    SYMPTOMS = ""
    If InStr(ArrRSInput(0, i + 2), "SYMPTOMS:") Or InStr(ArrRSInput(0, i + 3), "SYMPTOMS:") Then
        j = 2
        Do
            If Mid(ArrRSInput(0, i + j), 56, 12) = "For-Official" And Not Trim(Mid(ArrRSInput(0, (i + j) + 10), 64, 7)) = "(cont)" Then
                Exit Do
            End If
            If InStr(Trim(Mid(ArrRSInput(0, ((i + j) + 1)), 64, 9)), "SMR") Then
                Exit Do
            End If
            If ((i + j) + 11) <= UBound(ArrRSInput, 2) Then
                If Trim(Mid(ArrRSInput(0, (i + j) + 11), 64, 7)) = "(cont)" Then
                    j = j + 11
                End If
            Else
                Exit Do
            End If
            If Not Trim(Mid(ArrRSInput(0, i + j), 74, 57)) = "" Then
                SYMPTOMS = SYMPTOMS & Trim(Mid(ArrRSInput(0, i + j), 74, 57)) & vbNewLine
            End If
            If Left(ArrRSInput(0, i + j), 3) = "-- " Then
                Exit Do
            End If
            j = j + 1
        Loop
    End If

    'Write Maint Activity data to table
    RSSMR.AddNew
    RSSMR![Tail No] = TailNo
    RSSMR![Current_AFHR] = Current_AFHR
    RSSMR![Venue] = Venue
    RSSMR![TAIL/SERIAL NO] = Tail_SerialNo
    RSSMR!PART = PART
    RSSMR!CAGE = CAGE
    RSSMR![MAINT ACTIVITY] = Maint_Activity
    If Not Date_due = "" Then
        DateString = ConvertToDateSerial(Date_due)
        RSSMR![DATE DUE] = DateString
    End If
    RSSMR![LIFE TYPE INT] = Life_Type_Int
    If Not Int_Rem = "" And IsNumeric(Int_Rem) Then
        RSSMR!REM = Int_Rem
    End If
    RSSMR!LCN = LCN
    RSSMR!ALC = ALC
    RSSMR![POSITION NAME] = Position_Name
    RSSMR!WAC = WAC
    RSSMR!Type = Maint_Type
    RSSMR!LVL = LVL
    RSSMR!STD = STD
    RSSMR!REC = REC
    RSSMR!CLM = CLM
    RSSMR![DATE PLAN] = Date_Plan
    RSSMR!INTERVAL = INTERVAL
    RSSMR![LIFE SINCE New] = Life_Since_New
    RSSMR![LIFE SINCE New TYPE] = Life_Since_New_Type
    RSSMR![INIT TYPE] = Init_Type
    RSSMR!INSTANCE = INSTANCE
    RSSMR![INIT LCN] = Init_LCN
    RSSMR![INIT ALC] = Init_ALC
    RSSMR![SYMPTOMS] = SYMPTOMS
    RSSMR![RunDate] = RunDate
    RSSMR.Update

End Sub

Sub GetInProgressData()

    Dim Maint_Activity As Variant
    Dim Date_due As Variant
    Dim Life_Type_Int As Variant
    Dim Int_Rem As Variant
    Dim TSSCode1 As Variant
    Dim TSSCode2 As Variant
    Dim TSSCode3 As Variant
    Dim TSSCode4 As Variant
    Dim Maint_Type As Variant
    Dim LVL As Variant
    Dim STD As Variant
    Dim REC As Variant
    Dim CLM As Variant
    Dim Date_Plan As Variant
    Dim INTERVAL As Variant
    Dim Init_Type As Variant
    Dim INSTANCE As Variant
    Dim Init_LCN As Variant
    Dim Init_ALC As Variant
    Dim DateString As Variant
    Dim SYMPTOMS As Variant

    'Get Full line of data from CAMM2 if it exists
    If Not Left(CurrentLine, 63) = "                                                               " Then
        Tail_SerialNo = Trim(Left(PrevLine, 23))
        PART = Trim(Mid(PrevLine, 24, 30))
        CAGE = Mid(Trim(Mid(PrevLine, 56, 7)), 2, Len(Trim(Mid(PrevLine, 56, 7))))
        LCN = Trim(Left(CurrentLine, 18))
        ALC = Trim(Mid(CurrentLine, 19, 3))
        Position_Name = Trim(Mid(CurrentLine, 24, 19))
        WAC = Trim(Mid(CurrentLine, 44, 5))
        Life_Since_New = Trim(Mid(NextLine, 44, 9))
        Life_Since_New_Type = Trim(Mid(NextLine, 54, 7))
    End If
    'Link remaining data to the previous full line of data
    Maint_Activity = Trim(Mid(PrevLine, 64, 21))
    Date_due = Trim(Mid(PrevLine, 86, 10))
    Life_Type_Int = Trim(Mid(PrevLine, 96, 10))
    Int_Rem = Trim(Mid(PrevLine, 105, 10))
    TSSCode1 = Trim(Mid(PrevLine, 121, 10))
    TSSCode2 = Trim(Mid(CurrentLine, 121, 10))
    TSSCode3 = Trim(Mid(NextLine, 121, 10))
    TSSCode4 = Trim(Mid(NextLine, 121, 10))
    Maint_Type = Trim(Mid(CurrentLine, 64, 4))
    LVL = Trim(Mid(CurrentLine, 69, 3))
    STD = Trim(Mid(CurrentLine, 74, 2))
    REC = Trim(Mid(CurrentLine, 78, 3))
    CLM = Trim(Mid(CurrentLine, 82, 3))
    Date_Plan = Trim(Mid(CurrentLine, 86, 10))
    INTERVAL = Trim(Mid(CurrentLine, 96, 10))
    Init_Type = Trim(Mid(NextLine, 64, 5))
    INSTANCE = Trim(Mid(NextLine, 74, 10))
    Init_LCN = Trim(Mid(NextLine, 96, 11))
    Init_ALC = Trim(Mid(NextLine, 114, 4))
    'Write Maint Activity data to table
    RSInProgress.AddNew
    RSInProgress![Tail No] = TailNo
    RSInProgress![Current_AFHR] = Current_AFHR
    RSInProgress![Venue] = Venue
    RSInProgress![TAIL/SERIAL NO] = Tail_SerialNo
    RSInProgress!PART = PART
    RSInProgress!CAGE = CAGE
    RSInProgress![MAINT ACTIVITY] = Maint_Activity
    If Not Date_due = "" Then
        DateString = ConvertToDateSerial(Date_due)
        RSInProgress![DATE DUE] = DateString
    End If
    RSInProgress![LIFE TYPE INT] = Life_Type_Int
    If Not Int_Rem = "" And IsNumeric(Int_Rem) Then
        RSInProgress!REM = Int_Rem
    End If
    RSInProgress![TSS CODES 1] = TSSCode1
    RSInProgress![TSS CODES 2] = TSSCode2
    RSInProgress![TSS CODES 3] = TSSCode3
    RSInProgress![TSS CODES 4] = TSSCode4
    RSInProgress!LCN = LCN
    RSInProgress!ALC = ALC
    RSInProgress![POSITION NAME] = Position_Name
    RSInProgress!WAC = WAC
    RSInProgress!Type = Maint_Type
    RSInProgress!LVL = LVL
    RSInProgress!STD = STD
    RSInProgress!REC = REC
    RSInProgress!CLM = CLM
    RSInProgress![DATE PLAN] = Date_Plan
    RSInProgress!INTERVAL = INTERVAL
    RSInProgress![LIFE SINCE New] = Life_Since_New
    RSInProgress![LIFE SINCE New TYPE] = Life_Since_New_Type
    RSInProgress![INIT TYPE] = Init_Type
    RSInProgress!INSTANCE = INSTANCE
    RSInProgress![INIT LCN] = Init_LCN
    RSInProgress![INIT ALC] = Init_ALC
    RSInProgress![RunDate] = RunDate
    RSInProgress.Update

End Sub

Sub GetOverdueData()

    Dim Maint_Activity As Variant
    Dim Date_due As Variant
    Dim Life_Type_Int As Variant
    Dim Int_Rem As Variant
    Dim TSSCode1 As Variant
    Dim TSSCode2 As Variant
    Dim TSSCode3 As Variant
    Dim TSSCode4 As Variant
    Dim Maint_Type As Variant
    Dim LVL As Variant
    Dim STD As Variant
    Dim REC As Variant
    Dim CLM As Variant
    Dim Date_Plan As Variant
    Dim INTERVAL As Variant
    Dim Init_Type As Variant
    Dim INSTANCE As Variant
    Dim Init_LCN As Variant
    Dim Init_ALC As Variant
    Dim DateString As Variant
    Dim SYMPTOMS As Variant

    'Get Full line of data from CAMM2 if it exists
    If Not Left(CurrentLine, 63) = "                                                               " Then
        Tail_SerialNo = Trim(Left(PrevLine, 23))
        PART = Trim(Mid(PrevLine, 24, 30))
        CAGE = Mid(Trim(Mid(PrevLine, 56, 7)), 2, Len(Trim(Mid(PrevLine, 56, 7))))
        LCN = Trim(Left(CurrentLine, 18))
        ALC = Trim(Mid(CurrentLine, 19, 3))
        Position_Name = Trim(Mid(CurrentLine, 24, 19))
        WAC = Trim(Mid(CurrentLine, 44, 5))
        Life_Since_New = Trim(Mid(NextLine, 44, 9))
        Life_Since_New_Type = Trim(Mid(NextLine, 54, 7))
    End If
    'Link remaining data to the previous full line of data
    Maint_Activity = Trim(Mid(PrevLine, 64, 21))
    Date_due = Trim(Mid(PrevLine, 86, 10))
    Life_Type_Int = Trim(Mid(PrevLine, 96, 10))
    Int_Rem = Trim(Mid(PrevLine, 105, 10))
    TSSCode1 = Trim(Mid(PrevLine, 121, 10))
    TSSCode2 = Trim(Mid(CurrentLine, 121, 10))
    TSSCode3 = Trim(Mid(NextLine, 121, 10))
    TSSCode4 = Trim(Mid(NextLine, 121, 10))
    Maint_Type = Trim(Mid(CurrentLine, 64, 4))
    LVL = Trim(Mid(CurrentLine, 69, 3))
    STD = Trim(Mid(CurrentLine, 74, 2))
    REC = Trim(Mid(CurrentLine, 78, 3))
    CLM = Trim(Mid(CurrentLine, 82, 3))
    Date_Plan = Trim(Mid(CurrentLine, 86, 10))
    INTERVAL = Trim(Mid(CurrentLine, 96, 10))
    Init_Type = Trim(Mid(NextLine, 64, 5))
    INSTANCE = Trim(Mid(NextLine, 74, 18))
    Init_LCN = Trim(Mid(NextLine, 96, 11))
    Init_ALC = Trim(Mid(NextLine, 114, 4))
    'Write Maint Activity data to table
    RSOverdue.AddNew
    RSOverdue![Tail No] = TailNo
    RSOverdue![Current_AFHR] = Current_AFHR
    RSOverdue![Venue] = Venue
    RSOverdue![TAIL/SERIAL NO] = Tail_SerialNo
    RSOverdue!PART = PART
    RSOverdue!CAGE = CAGE
    RSOverdue![MAINT ACTIVITY] = Maint_Activity
    If Not Date_due = "" Then
        DateString = ConvertToDateSerial(Date_due)
        RSOverdue![DATE DUE] = DateString
    End If
    RSOverdue![LIFE TYPE INT] = Life_Type_Int
    If Not Int_Rem = "" And IsNumeric(Int_Rem) Then
        RSOverdue!REM = Int_Rem
    End If
    RSOverdue![TSS CODES 1] = TSSCode1
    RSOverdue![TSS CODES 2] = TSSCode2
    RSOverdue![TSS CODES 3] = TSSCode3
    RSOverdue![TSS CODES 4] = TSSCode4
    RSOverdue!LCN = LCN
    RSOverdue!ALC = ALC
    RSOverdue![POSITION NAME] = Position_Name
    RSOverdue!WAC = WAC
    RSOverdue!Type = Maint_Type
    RSOverdue!LVL = LVL
    RSOverdue!STD = STD
    RSOverdue!REC = REC
    RSOverdue!CLM = CLM
    RSOverdue![DATE PLAN] = Date_Plan
    RSOverdue!INTERVAL = INTERVAL
    RSOverdue![LIFE SINCE New] = Life_Since_New
    RSOverdue![LIFE SINCE New TYPE] = Life_Since_New_Type
    RSOverdue![INIT TYPE] = Init_Type
    RSOverdue!INSTANCE = INSTANCE
    RSOverdue![INIT LCN] = Init_LCN
    RSOverdue![INIT ALC] = Init_ALC
    RSOverdue![RunDate] = RunDate
    RSOverdue.Update

End Sub

Private Sub Command77_Click()

    DoCmd.SetWarnings False

    If MsgBox("Are you sure you want to delete all MFT data?", vbYesNo + vbCritical) = vbNo Then Exit Sub

    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Activity"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Deviation"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Deviation_Rpt"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Coincident"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_CFU"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_CFU_Rpt"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_In_Progress"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Overdue"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_SMR"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_SMR_Rpt"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Date_Driven"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Event_Driven"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Hour_Driven"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Header_Info"
    Me.Requery
    Me.Refresh

    DoCmd.SetWarnings True
End Sub



Function GetFileName(Optional OpenAt As String) As String

    Dim lCount As Long

    GetFileName = vbNullString

    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = OpenAt
        .Show
        For lCount = 1 To .SelectedItems.Count
            GetFileName = .SelectedItems(lCount)
        Next lCount
    End With

End Function

Function GetFolderName(Optional OpenAt As String) As String

    Dim lCount As Long

    GetFolderName = vbNullString

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = OpenAt
        .Show
        For lCount = 1 To .SelectedItems.Count
            GetFolderName = .SelectedItems(lCount)
        Next lCount
    End With

End Function

Function CountFiles(Optional strExt As String) As Double
    'Count files in a directory.  If a file extension is provided,
    'then count only files of that type, otherwise return a count of all files.
    Dim objFso As Object
    Dim objFiles As Object
    Dim objFile As Object

    'Set Error Handling
    On Error GoTo EarlyExit
    'Create objects to get a count of files in the directory
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFiles = objFso.GetFolder(MFTFolderName).Files
    'Count files (that match the extension if provided)
    If strExt = "*.*" Then ' All files
        CountFiles = objFiles.Count
    Else
        For Each objFile In objFiles
            If UCase(Right(objFile.Path, (Len(objFile.Path) - InStrRev(objFile.Path, ".")))) = UCase(strExt) Then
                CountFiles = CountFiles + 1
            End If
        Next objFile
    End If

EarlyExit:
    'Clean up
    On Error Resume Next
    Set objFile = Nothing
    Set objFiles = Nothing
    Set objFso = Nothing
    On Error GoTo 0

End Function

Sub SortDataForReporting()

    Dim RST As ADODB.Recordset
    Dim RSTemp As ADODB.Recordset

    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Event_Driven"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Date_Driven"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Hour_Driven"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_Deviation_Rpt"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_CFU_Rpt"
    DoCmd.RunSQL "DELETE * FROM tbl_MFT_SMR_Rpt"
    
    Set RSEventDriven = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_Event_Driven")
    Set RSDateDriven = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_Date_Driven")
    Set RSHourDriven = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_Hour_Driven")
    Set RSDeviationRpt = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_Deviation_Rpt")
    Set RSCFURpt = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_CFU_Rpt")
    Set RSSMRRpt = CurrentDb.OpenRecordset("SELECT * FROM tbl_MFT_SMR_Rpt")
    
    Set RST = CurrentProject.Connection.Execute("SELECT * FROM tbl_MFT_Activity")
    If RST.EOF = False Then
        ArrActivity = RST.GetRows
    End If
    Set RST = CurrentProject.Connection.Execute("SELECT * FROM tbl_MFT_Coincident")
    If RST.EOF = False Then
        ArrCoincident = RST.GetRows
    End If
    Set RST = CurrentProject.Connection.Execute("SELECT * FROM tbl_MFT_In_Progress")
    If RST.EOF = False Then
        ArrInProgress = RST.GetRows
    End If
    Set RST = CurrentProject.Connection.Execute("SELECT * FROM tbl_MFT_Overdue")
    If RST.EOF = False Then
        ArrOverdue = RST.GetRows
    End If
    Set RST = CurrentProject.Connection.Execute("SELECT * FROM tbl_MFT_Deviation")
    If RST.EOF = False Then
        ArrDeviation = RST.GetRows
    End If
    Set RST = CurrentProject.Connection.Execute("SELECT * FROM tbl_MFT_CFU")
    If RST.EOF = False Then
        ArrCFU = RST.GetRows
    End If
    Set RST = CurrentProject.Connection.Execute("SELECT * FROM tbl_MFT_SMR")
    If RST.EOF = False Then
        ArrSMR = RST.GetRows
    End If

    LoadBarStatus steps, LoadBarStep
    LoadBarStep = LoadBarStep + 1

    ArrTemp = ArrActivity
    'tbl_MFT_Activity
    If Not IsEmpty(ArrActivity) = True Then
        For i_row = 0 To UBound(ArrActivity, 2)
            ResetFlags()
            'Date Driven
            If Not ArrActivity(7, i_row) = "" Then
                WriteDateDriven i_row, "Activity Activity - Date Driven"
            'Hour Driven
            ElseIf Right(ArrActivity(8, i_row), 2) = "HR" Then
                WriteHourDriven i_row, "Activity Activity - Hour Driven"
            'Event Driven
            Else
                WriteEventDriven i_row, "Activity Activity - Event Driven"
            End If
        Next
    End If

    LoadBarStatus steps, LoadBarStep
    LoadBarStep = LoadBarStep + 1

    ArrTemp = ArrCoincident
    'tbl_MFT_Coincident
    If Not IsEmpty(ArrCoincident) = True Then
        For i_row = 0 To UBound(ArrCoincident, 2)
            ResetFlags()
            'Date Driven
            If Not ArrCoincident(7, i_row) = "" Then
                'WriteDateDriven i_row, "Coincident Coincident - Date Driven"
                'Coincident Date Driven
            ElseIf ArrCoincident(27, i_row) = "STRT" Or ArrCoincident(27, i_row) = "CMPL" Then
                SortData "Coincident"
            'Hour Driven
            ElseIf Right(ArrCoincident(8, i_row), 2) = "HR" Then
                'WriteHourDriven i_row, "Coincident Coincident - Hour Driven"
            End If
        Next
    End If

    LoadBarStatus steps, LoadBarStep
    LoadBarStep = LoadBarStep + 1

    ArrTemp = ArrInProgress
    'tbl_MFT_In_Progress
    If Not IsEmpty(ArrInProgress) = True Then
        For i_row = 0 To UBound(ArrInProgress, 2)
            ResetFlags()
            'Date Driven
            If Not ArrInProgress(7, i_row) = "" Then
                WriteDateDriven i_row, "In Progress In Progress - Date Driven", , , "In Progress"
            'Hour Driven
            ElseIf Right(ArrInProgress(8, i_row), 2) = "HR" Then
                WriteHourDriven i_row, "In Progress In Progress - Hour Driven", , , "In Progress"
            'In Progress Date Driven
            ElseIf Not ArrInProgress(27, i_row) = "" Then
                SortData "In Progress"
            End If
        Next
    End If

    LoadBarStatus steps, LoadBarStep
    LoadBarStep = LoadBarStep + 1

    ArrTemp = ArrOverdue
    'tbl_MFT_Overdue
    If Not IsEmpty(ArrOverdue) = True Then
        For i_row = 0 To UBound(ArrOverdue, 2)
            ResetFlags()
            strRunDate = ArrOverdue(32, i_row)
            'Date Driven
            If Not ArrOverdue(7, i_row) = "" Then
                WriteDateDriven i_row, "Overdue Overdue - Date Driven", strRunDate
            'Overdue Date Driven
            ElseIf Not ArrOverdue(27, i_row) = "" And Not ArrOverdue(27, i_row) = "" And Not Right(ArrOverdue(8, i_row), 2) = "HR" Then
                WriteDateDriven i_row, "Overdue Nil Date - Date Driven", strRunDate
            'Hour Driven
            ElseIf Right(ArrOverdue(8, i_row), 2) = "HR" Then
                WriteHourDriven i_row, "Overdue Overdue - Hour Driven"
            End If
        Next
    End If

    LoadBarStatus steps, LoadBarStep
    LoadBarStep = LoadBarStep + 1

    ArrTemp = ArrDeviation
    'tbl_MFT_Deviation
    If Not IsEmpty(ArrDeviation) = True Then
        For i_row = 0 To UBound(ArrDeviation, 2)
            ResetFlags()

            If Right(ArrDeviation(17, i_row), 2) = "HR" Then
                WriteDeviationRpt i_row, "Deviation Deviation - Hour Driven"
            ElseIf ArrDeviation(17, i_row) = "DATE" Then
                WriteDeviationRpt i_row, "Deviation Deviation - Date Driven"
            ElseIf ArrDeviation(17, i_row) = "STRT" Or ArrDeviation(17, i_row) = "CMPL" Then
                SortDataDeviation "Deviation"
            ElseIf ArrDeviation(17, i_row) = "" And ArrDeviation(16, i_row) = "NO EXPIRY CONDITIONS" Then
                WriteDeviationRpt i_row, "Deviation Deviation - No Expiry"
            End If
        Next
    End If

    LoadBarStatus steps, LoadBarStep
    LoadBarStep = LoadBarStep + 1

    ArrTemp = ArrCFU
    'tbl_MFT_CFU
    If Not IsEmpty(ArrCFU) = True Then
        For i_row = 0 To UBound(ArrCFU, 2)
            ResetFlags()

            If Right(ArrCFU(8, i_row), 2) = "HR" Then
                WriteCFURpt i_row, "CFU CFU - Hour Driven"
            ElseIf Not ArrCFU(7, i_row) = "" Then
                WriteCFURpt i_row, "CFU CFU - Date Driven"
            ElseIf ArrCFU(23, i_row) = "STRT" Or ArrCFU(23, i_row) = "CMPL" Then
                SortDataCFU "CFU"
            End If
        Next
    End If

    LoadBarStatus steps, LoadBarStep
    LoadBarStep = LoadBarStep + 1

    ArrTemp = ArrSMR
    'tbl_MFT_SMR
    If Not IsEmpty(ArrSMR) = True Then
        For i_row = 0 To UBound(ArrSMR, 2)
            ResetFlags()

            If Right(ArrSMR(8, i_row), 2) = "HR" Then
                WriteSMRRpt i_row, "SMR SMR - Hour Driven"
            ElseIf Not ArrSMR(7, i_row) = "" Then
                WriteSMRRpt i_row, "SMR SMR - Date Driven"
            ElseIf ArrSMR(23, i_row) = "STRT" Or ArrSMR(23, i_row) = "CMPL" Then
                SortDataSMR "SMR"
            End If
        Next
    End If

End Sub

Sub ResetFlags()

    Dim FlagFound As Boolean
    Dim FlagDontAssign As Boolean
    Dim Duplicate As Boolean
    Dim FoundDate As String
    Dim FoundHour As String
    Dim InitParent As String
    Dim InitBy As String
    Dim InitType As String
    Dim NoDuplicateServ As Integer

    FlagFound = False
    FlagDontAssign = False
    Duplicate = False
    FoundDate = ""
    FoundHour = ""
    InitParent = ""
    InitBy = ""
    InitType = ""
    NoDuplicateServ = 0

End Sub

Sub SortData(ActivityType As String)

    i_Temp = i_row
    i = i_row
    ResetFlags()

    Dim InitBy As Variant
    Dim FlagFound As Boolean

    If Not ArrTemp(28, i) = "OH" And Not ArrTemp(28, i) = "BS" Then
        'Check Activity
        For j = 0 To UBound(ArrActivity, 2)
            If ArrTemp(0, i) = ArrActivity(0, j) And ArrTemp(28, i) = ArrActivity(6, j) Then
                If ArrTemp(29, i) = ArrActivity(14, j) Or ArrTemp(29, i) = "" Then
                    If Not ArrActivity(7, j) = "" Then
                        FoundDate = ArrActivity(7, j)
                        InitParent = ArrActivity(16, j)
                        InitBy = ArrTemp(28, i)
                        InitType = ArrTemp(27, i)
                        WriteTypeDriven i_Temp, ActivityType & " Activity - Date Driven", FoundDate, FoundHour, InitBy, ActivityType
                        FlagFound = True
                        'Exit For
                    ElseIf Right(ArrActivity(8, j), 2) = "HR" Then
                        FoundHour = ArrActivity(9, j)
                        InitParent = ArrActivity(16, j)
                        InitBy = ArrTemp(28, i)
                        InitType = ArrTemp(27, i)
                        WriteTypeDriven i_Temp, ActivityType & " Activity - Hour Driven", FoundDate, FoundHour, InitBy, ActivityType
                        FlagFound = True
                        'Exit For
                    End If
                End If
            End If
        Next
        'Check Coincident
        If FlagFound = False Then
            For j = 0 To UBound(ArrCoincident, 2)
                If ArrTemp(0, i) = ArrCoincident(0, j) And ArrTemp(28, i) = ArrCoincident(6, j) Then
                    If ArrTemp(29, i) = ArrCoincident(14, j) Or ArrTemp(29, i) = "" Then
                        If Not ArrCoincident(7, j) = "" Then
                            FoundDate = ArrCoincident(7, j)
                            InitParent = ArrCoincident(16, j)
                            InitBy = ArrTemp(28, i)
                            InitType = ArrTemp(27, i)
                            WriteTypeDriven i_Temp, ActivityType & " Coincident - Date Driven", FoundDate, FoundHour, InitBy, ActivityType
                            FlagFound = True
                            Exit For
                        ElseIf Right(ArrCoincident(8, j), 2) = "HR" Then
                            FoundHour = ArrCoincident(9, j)
                            InitParent = ArrCoincident(16, j)
                            InitBy = ArrTemp(28, i)
                            InitType = ArrTemp(27, i)
                            WriteTypeDriven i_Temp, ActivityType & " Coincident - Hour Driven", FoundDate, FoundHour, InitBy, ActivityType
                            FlagFound = True
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        'Check In Progress
        If FlagFound = False Then
            If Not IsEmpty(ArrInProgress) = True Then
                For j = 0 To UBound(ArrInProgress, 2)
                    If ArrTemp(0, i) = ArrInProgress(0, j) And ArrTemp(28, i) = ArrInProgress(6, j) Then
                        If ArrTemp(29, i) = ArrInProgress(14, j) Or ArrTemp(29, i) = "" Then
                            If Not ArrInProgress(7, j) = "" Then
                                FoundDate = ArrInProgress(7, j)
                                If ActivityType = "In Progress" Then
                                    InitParent = "In Progress"
                                Else
                                    InitParent = ArrInProgress(16, j)
                                End If
                                InitBy = ArrTemp(28, i)
                                InitType = ArrTemp(27, i)
                                WriteTypeDriven i_Temp, ActivityType & " In Progress - Date Driven", FoundDate, FoundHour, InitBy, ActivityType
                                FlagFound = True
                                Exit For
                            ElseIf Right(ArrInProgress(8, j), 2) = "HR" Then
                                FoundHour = ArrInProgress(9, j)
                                If ActivityType = "In Progress" Then
                                    InitParent = "In Progress"
                                Else
                                    InitParent = ArrInProgress(16, j)
                                End If
                                InitBy = ArrTemp(28, i)
                                InitType = ArrTemp(27, i)
                                WriteTypeDriven i_Temp, ActivityType & " In Progress - Hour Driven", FoundDate, FoundHour, InitBy, ActivityType
                                FlagFound = True
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If
        End If
        'Check Overdue
        If FlagFound = False Then
            If Not IsEmpty(ArrOverdue) = True Then
                For j = 0 To UBound(ArrOverdue, 2)
                    If ArrTemp(0, i) = ArrOverdue(0, j) And ArrTemp(28, i) = ArrOverdue(6, j) Then
                        If ArrTemp(29, i) = ArrOverdue(14, j) Or ArrTemp(29, i) = "" Then
                            If Not ArrOverdue(7, j) = "" Then
                                FoundDate = ArrOverdue(7, j)
                                InitParent = ArrOverdue(16, j)
                                InitBy = ArrTemp(28, i)
                                InitType = ArrTemp(27, i)
                                WriteTypeDriven i_Temp, ActivityType & " Overdue - Date Driven", FoundDate, FoundHour, InitBy, ActivityType
                                FlagFound = True
                                Exit For
                            ElseIf Right(ArrOverdue(8, j), 2) = "HR" Then
                                FoundHour = ArrOverdue(9, j)
                                InitParent = ArrOverdue(16, j)
                                InitBy = ArrTemp(28, i)
                                InitType = ArrTemp(27, i)
                                WriteTypeDriven i_Temp, ActivityType & " Overdue - Hour Driven", FoundDate, FoundHour, InitBy, ActivityType
                                FlagFound = True
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End If

End Sub

Sub WriteTypeDriven(i_row As Integer, SortType As String, Optional FoundDateVal As String, Optional FoundHourVal As String, Optional InitByVal As Variant, Optional ActivityTypeVal As String)

    Dim RSTemp As Variant
    Dim DateString As Variant

    If Not FoundDateVal = "" Then
        Set RSTemp = RSDateDriven
    ElseIf Not FoundHourVal = "" Then
        Set RSTemp = RSHourDriven
    End If

    RSTemp.AddNew
    RSTemp![Tail No] = ArrTemp(0, i_row)
    RSTemp![Current_AFHR] = ArrTemp(1, i_row)
    RSTemp![Venue] = ArrTemp(2, i_row)
    RSTemp![TAIL/SERIAL NO] = ArrTemp(3, i_row)
    RSTemp!PART = ArrTemp(4, i_row)
    RSTemp!CAGE = ArrTemp(5, i_row)
    RSTemp![MAINT ACTIVITY] = ArrTemp(6, i_row)
    If Not FoundDateVal = "" Then
        DateString = ConvertToDateSerial(FoundDateVal)
        RSTemp![DATE DUE] = DateString
    Else
        If Not ArrTemp(7, i_row) = "" Then
            DateString = ConvertToDateSerial(ArrTemp(7, i_row))
            RSTemp![DATE DUE] = DateString
        End If
    End If
    If Not InitType = "" And Not InitType = "DATE" Then
        RSTemp![LIFE TYPE INT] = InitType
    Else
        RSTemp![LIFE TYPE INT] = ArrTemp(8, i_row)
    End If
    If Not FoundHourVal = "" Then
        RSTemp!REM = FoundHourVal
    Else
        RSTemp!REM = ArrTemp(9, i_row)
    End If
    RSTemp![TSS CODES 1] = ArrTemp(10, i_row)
    RSTemp![TSS CODES 2] = ArrTemp(11, i_row)
    RSTemp![TSS CODES 3] = ArrTemp(12, i_row)
    RSTemp![TSS CODES 4] = ArrTemp(13, i_row)
    RSTemp!LCN = ArrTemp(14, i_row)
    RSTemp!ALC = ArrTemp(15, i_row)
    RSTemp![POSITION NAME] = ArrTemp(16, i_row)
    RSTemp!WAC = ArrTemp(17, i_row)
    RSTemp!Type = ArrTemp(18, i_row)
    RSTemp!LVL = ArrTemp(19, i_row)
    RSTemp!STD = ArrTemp(20, i_row)
    RSTemp!REC = ArrTemp(21, i_row)
    RSTemp!CLM = ArrTemp(22, i_row)
    RSTemp![DATE PLAN] = ArrTemp(23, i_row)
    RSTemp!INTERVAL = ArrTemp(24, i_row)
    RSTemp![LIFE SINCE New] = ArrTemp(25, i_row)
    RSTemp![LIFE SINCE New TYPE] = ArrTemp(26, i_row)
    RSTemp![INIT TYPE] = ArrTemp(27, i_row)
    RSTemp!INSTANCE = ArrTemp(28, i_row)
    RSTemp![INIT LCN] = ArrTemp(29, i_row)
    RSTemp![INIT ALC] = ArrTemp(30, i_row)
    If Not IsMissing(InitByVal) Then
        RSTemp![INIT BY] = InitByVal
    End If
    RSTemp![INIT PARENT] = InitParent
    RSTemp![Displayed] = SortType
    RSTemp![RunDate] = ArrTemp(32, i_row)
    RSTemp.Update

End Sub

Sub WriteDateDriven(i_row As Integer, SortType As String, Optional FoundDateVal As String, Optional InitByVal As Variant, Optional InitByParentVal As Variant)

    Dim DateString As Variant

    RSDateDriven.AddNew
    RSDateDriven![Tail No] = ArrTemp(0, i_row)
    RSDateDriven![Current_AFHR] = ArrTemp(1, i_row)
    RSDateDriven![Venue] = ArrTemp(2, i_row)
    RSDateDriven![TAIL/SERIAL NO] = ArrTemp(3, i_row)
    RSDateDriven!PART = ArrTemp(4, i_row)
    RSDateDriven!CAGE = ArrTemp(5, i_row)
    RSDateDriven![MAINT ACTIVITY] = ArrTemp(6, i_row)
    If Not FoundDateVal = "" Then
        If IsDate(FoundDateVal) Then
            RSDateDriven![DATE DUE] = FoundDateVal
        Else
            DateString = ConvertToDateSerial(FoundDateVal)
            RSDateDriven![DATE DUE] = DateString
        End If
    Else
        If Not ArrTemp(7, i_row) = "" Then
            DateString = ConvertToDateSerial(ArrTemp(7, i_row))
            RSDateDriven![DATE DUE] = DateString
        End If
    End If
    RSDateDriven![LIFE TYPE INT] = ArrTemp(8, i_row)
    RSDateDriven!REM = ArrTemp(9, i_row)
    RSDateDriven![TSS CODES 1] = ArrTemp(10, i_row)
    RSDateDriven![TSS CODES 2] = ArrTemp(11, i_row)
    RSDateDriven![TSS CODES 3] = ArrTemp(12, i_row)
    RSDateDriven![TSS CODES 4] = ArrTemp(13, i_row)
    RSDateDriven!LCN = ArrTemp(14, i_row)
    RSDateDriven!ALC = ArrTemp(15, i_row)
    RSDateDriven![POSITION NAME] = ArrTemp(16, i_row)
    RSDateDriven!WAC = ArrTemp(17, i_row)
    RSDateDriven!Type = ArrTemp(18, i_row)
    RSDateDriven!LVL = ArrTemp(19, i_row)
    RSDateDriven!STD = ArrTemp(20, i_row)
    RSDateDriven!REC = ArrTemp(21, i_row)
    RSDateDriven!CLM = ArrTemp(22, i_row)
    RSDateDriven![DATE PLAN] = ArrTemp(23, i_row)
    RSDateDriven!INTERVAL = ArrTemp(24, i_row)
    RSDateDriven![LIFE SINCE New] = ArrTemp(25, i_row)
    RSDateDriven![LIFE SINCE New TYPE] = ArrTemp(26, i_row)
    RSDateDriven![INIT TYPE] = ArrTemp(27, i_row)
    RSDateDriven!INSTANCE = ArrTemp(28, i_row)
    RSDateDriven![INIT LCN] = ArrTemp(29, i_row)
    RSDateDriven![INIT ALC] = ArrTemp(30, i_row)
    If Not IsMissing(InitByVal) Then
        RSDateDriven![INIT BY] = InitByVal
    End If
    If Not IsMissing(InitByParentVal) Then
        RSDateDriven![INIT PARENT] = InitByParentVal
    End If
    RSDateDriven![Displayed] = SortType
    RSDateDriven![RunDate] = ArrTemp(32, i_row)
    RSDateDriven.Update

End Sub

Sub WriteHourDriven(i_row As Integer, SortType As String, Optional FoundHourVal As String, Optional InitByVal As Variant, Optional InitByParentVal As Variant)

    RSHourDriven.AddNew
    RSHourDriven![Tail No] = ArrTemp(0, i_row)
    RSHourDriven![Current_AFHR] = ArrTemp(1, i_row)
    RSHourDriven![Venue] = ArrTemp(2, i_row)
    RSHourDriven![TAIL/SERIAL NO] = ArrTemp(3, i_row)
    RSHourDriven!PART = ArrTemp(4, i_row)
    RSHourDriven!CAGE = ArrTemp(5, i_row)
    RSHourDriven![MAINT ACTIVITY] = ArrTemp(6, i_row)
    RSHourDriven![DATE DUE] = ArrTemp(7, i_row)
    RSHourDriven![LIFE TYPE INT] = ArrTemp(8, i_row)
    If Not FoundHourVal = "" Then
        RSHourDriven!REM = FoundHourVal
    Else
        RSHourDriven!REM = ArrTemp(9, i_row)
    End If
    RSHourDriven![TSS CODES 1] = ArrTemp(10, i_row)
    RSHourDriven![TSS CODES 2] = ArrTemp(11, i_row)
    RSHourDriven![TSS CODES 3] = ArrTemp(12, i_row)
    RSHourDriven![TSS CODES 4] = ArrTemp(13, i_row)
    RSHourDriven!LCN = ArrTemp(14, i_row)
    RSHourDriven!ALC = ArrTemp(15, i_row)
    RSHourDriven![POSITION NAME] = ArrTemp(16, i_row)
    RSHourDriven!WAC = ArrTemp(17, i_row)
    RSHourDriven!Type = ArrTemp(18, i_row)
    RSHourDriven!LVL = ArrTemp(19, i_row)
    RSHourDriven!STD = ArrTemp(20, i_row)
    RSHourDriven!REC = ArrTemp(21, i_row)
    RSHourDriven!CLM = ArrTemp(22, i_row)
    RSHourDriven![DATE PLAN] = ArrTemp(23, i_row)
    RSHourDriven!INTERVAL = ArrTemp(24, i_row)
    RSHourDriven![LIFE SINCE New] = ArrTemp(25, i_row)
    RSHourDriven![LIFE SINCE New TYPE] = ArrTemp(26, i_row)
    RSHourDriven![INIT TYPE] = ArrTemp(27, i_row)
    RSHourDriven!INSTANCE = ArrTemp(28, i_row)
    RSHourDriven![INIT LCN] = ArrTemp(29, i_row)
    RSHourDriven![INIT ALC] = ArrTemp(30, i_row)
    If Not IsMissing(InitByVal) Then
        RSHourDriven![INIT BY] = InitByVal
    End If
    If Not IsMissing(InitByParentVal) Then
        RSHourDriven![INIT PARENT] = InitByParentVal
    End If
    RSHourDriven![Displayed] = SortType
    RSHourDriven![RunDate] = ArrTemp(32, i_row)
    RSHourDriven.Update

End Sub

Sub WriteEventDriven(i_row As Integer, SortType As String, Optional FoundHourVal As String, Optional InitByVal As Variant, Optional InitByParentVal As Variant)

    RSEventDriven.AddNew
    RSEventDriven![Tail No] = ArrTemp(0, i_row)
    RSEventDriven![Current_AFHR] = ArrTemp(1, i_row)
    RSEventDriven![Venue] = ArrTemp(2, i_row)
    RSEventDriven![TAIL/SERIAL NO] = ArrTemp(3, i_row)
    RSEventDriven!PART = ArrTemp(4, i_row)
    RSEventDriven!CAGE = ArrTemp(5, i_row)
    RSEventDriven![MAINT ACTIVITY] = ArrTemp(6, i_row)
    RSEventDriven![DATE DUE] = ArrTemp(7, i_row)
    RSEventDriven![LIFE TYPE INT] = ArrTemp(8, i_row)
    RSEventDriven!REM = ArrTemp(9, i_row)
    RSEventDriven![TSS CODES 1] = ArrTemp(10, i_row)
    RSEventDriven![TSS CODES 2] = ArrTemp(11, i_row)
    RSEventDriven![TSS CODES 3] = ArrTemp(12, i_row)
    RSEventDriven![TSS CODES 4] = ArrTemp(13, i_row)
    RSEventDriven!LCN = ArrTemp(14, i_row)
    RSEventDriven!ALC = ArrTemp(15, i_row)
    RSEventDriven![POSITION NAME] = ArrTemp(16, i_row)
    RSEventDriven!WAC = ArrTemp(17, i_row)
    RSEventDriven!Type = ArrTemp(18, i_row)
    RSEventDriven!LVL = ArrTemp(19, i_row)
    RSEventDriven!STD = ArrTemp(20, i_row)
    RSEventDriven!REC = ArrTemp(21, i_row)
    RSEventDriven!CLM = ArrTemp(22, i_row)
    RSEventDriven![DATE PLAN] = ArrTemp(23, i_row)
    RSEventDriven!INTERVAL = ArrTemp(24, i_row)
    RSEventDriven![LIFE SINCE New] = ArrTemp(25, i_row)
    RSEventDriven![LIFE SINCE New TYPE] = ArrTemp(26, i_row)
    RSEventDriven![INIT TYPE] = ArrTemp(27, i_row)
    RSEventDriven!INSTANCE = ArrTemp(28, i_row)
    RSEventDriven![INIT LCN] = ArrTemp(29, i_row)
    RSEventDriven![INIT ALC] = ArrTemp(30, i_row)
    If Not IsMissing(InitByVal) Then
        RSEventDriven![INIT BY] = InitByVal
    End If
    If Not IsMissing(InitByParentVal) Then
        RSEventDriven![INIT PARENT] = InitByParentVal
    End If
    RSEventDriven![Displayed] = SortType
    RSEventDriven![RunDate] = ArrTemp(32, i_row)
    RSEventDriven.Update

End Sub

Sub SortDataDeviation(ActivityType As String)

    i_Temp = i_row
    i = i_row
    ResetFlags()

    Dim InitBy As Variant
    Dim FlagFound As Variant

    If Not ArrTemp(18, i) = "OH" And Not ArrTemp(18, i) = "BS" Then
        'Check Activity
        For j = 0 To UBound(ArrActivity, 2)
            If ArrTemp(0, i) = ArrActivity(0, j) And ArrTemp(18, i) = ArrActivity(6, j) Then
                If (ArrTemp(19, i) = ArrActivity(14, j) And ArrTemp(20, i) = ArrActivity(15, j)) Or ArrTemp(19, i) = "" Then
                    If Not ArrActivity(7, j) = "" Then
                        FoundDate = ArrActivity(7, j)
                        InitParent = ArrActivity(16, j)
                        InitBy = ArrTemp(18, i)
                        InitType = ArrTemp(17, i)
                    ElseIf Right(ArrActivity(8, j), 2) = "HR" Then
                        FoundHour = ArrActivity(9, j)
                        InitParent = ArrActivity(16, j)
                        InitBy = ArrTemp(18, i)
                        InitType = ArrTemp(17, i)
                    End If
                    If Not FoundDate = "" And Not FoundHour = "" Then
                        FlagFound = True
                        Exit For
                    End If
                End If
            End If
        Next
        If FlagFound = True Or Not FoundDate = "" Or Not FoundHour = "" Then
            WriteDeviationRpt i_Temp, ActivityType & " Activity", FoundDate, FoundHour, InitBy, ActivityType
        End If
    End If

End Sub

Sub WriteDeviationRpt(i_row As Integer, SortType As String, Optional FoundDateVal As String, Optional FoundHourVal As String, Optional InitByVal As Variant, Optional ActivityTypeVal As String)

    Dim DateString As Variant
    Dim InitByParentVal As Variant

    RSDeviationRpt.AddNew
    RSDeviationRpt![Tail No] = ArrTemp(0, i_row)
    RSDeviationRpt![Current_AFHR] = ArrTemp(1, i_row)
    RSDeviationRpt![Venue] = ArrTemp(2, i_row)
    RSDeviationRpt![TAIL/SERIAL NO] = ArrTemp(3, i_row)
    RSDeviationRpt!PART = ArrTemp(4, i_row)
    RSDeviationRpt!CAGE = ArrTemp(5, i_row)
    RSDeviationRpt![DEVIATION ORDER] = ArrTemp(6, i_row)
    RSDeviationRpt![STD] = ArrTemp(7, i_row)
    RSDeviationRpt![EXPIRED] = ArrTemp(8, i_row)
    RSDeviationRpt![LCN] = ArrTemp(9, i_row)
    RSDeviationRpt![ALC] = ArrTemp(10, i_row)
    RSDeviationRpt![POSITION NAME] = ArrTemp(11, i_row)
    RSDeviationRpt![WAC] = ArrTemp(12, i_row)
    RSDeviationRpt![TITLE] = ArrTemp(13, i_row)
    RSDeviationRpt![LIFE SINCE New] = ArrTemp(14, i_row)
    RSDeviationRpt![LIFE SINCE New TYPE] = ArrTemp(15, i_row)
    RSDeviationRpt![EXPIRY CONDITIONS] = ArrTemp(16, i_row)
    RSDeviationRpt![Type] = ArrTemp(17, i_row)
    RSDeviationRpt![ACTIVITY] = ArrTemp(18, i_row)
    RSDeviationRpt![INIT LCN] = ArrTemp(19, i_row)
    RSDeviationRpt![INIT ALC] = ArrTemp(20, i_row)
    RSDeviationRpt![INTERVAL] = ArrTemp(21, i_row)
    If Not FoundHourVal = "" Then
        RSDeviationRpt![REM] = FoundHourVal
    Else
        RSDeviationRpt![REM] = ArrTemp(22, i_row)
    End If
    If Not FoundDateVal = "" Then
        If IsDate(FoundDateVal) Then
            RSDeviationRpt![DATE DUE] = FoundDateVal
        Else
            DateString = ConvertToDateSerial(FoundDateVal)
            RSDeviationRpt![DATE DUE] = DateString
        End If
    Else
        If Not ArrTemp(23, i_row) = "" Then
            DateString = ConvertToDateSerial(ArrTemp(23, i_row))
            RSDeviationRpt![DATE DUE] = DateString
        End If
    End If
    RSDeviationRpt![LIMITATION] = ArrTemp(24, i_row)
    If Not IsMissing(InitByVal) Then
        RSDeviationRpt![INIT BY] = InitByVal
    End If
    If Not IsMissing(InitByParentVal) Then
        RSDeviationRpt![INIT PARENT] = InitByParentVal
    End If
    RSDeviationRpt![Displayed] = SortType
    RSDeviationRpt![RunDate] = ArrTemp(25, i_row)
    RSDeviationRpt.Update

End Sub

Sub SortDataCFU(ActivityType As String)

    i_Temp = i_row
    i = i_row
    ResetFlags()

    Dim InitBy As Variant
    Dim FlagFound As Boolean

    If Not ArrTemp(24, i) = "OH" And Not ArrTemp(24, i) = "BS" Then
        'Check Activity
        For j = 0 To UBound(ArrActivity, 2)
            If ArrTemp(0, i) = ArrActivity(0, j) And ArrTemp(24, i) = ArrActivity(6, j) Then
                If (ArrTemp(25, i) = ArrActivity(14, j) And ArrTemp(26, i) = ArrActivity(15, j)) Or ArrTemp(25, i) = "" Then
                    If Not ArrActivity(7, j) = "" Then
                        FoundDate = ArrActivity(7, j)
                        InitParent = ArrActivity(16, j)
                        InitBy = ArrTemp(24, i)
                        InitType = ArrTemp(23, i)
                        'WriteCFURpt i_Temp, ActivityType & " Activity - Date Driven", FoundDate, FoundHour, InitBy, ActivityType
                        'FlagFound = True
                        'Exit For
                    ElseIf Right(ArrActivity(8, j), 2) = "HR" Then
                        FoundHour = ArrActivity(9, j)
                        InitParent = ArrActivity(16, j)
                        InitBy = ArrTemp(24, i)
                        InitType = ArrTemp(23, i)
                        'WriteCFURpt i_Temp, ActivityType & " Activity - Hour Driven", FoundDate, FoundHour, InitBy, ActivityType
                        'FlagFound = True
                        'Exit For
                    End If
                    If Not FoundDate = "" And Not FoundHour = "" Then
                        FlagFound = True
                        Exit For
                    End If
                End If
            End If
        Next
        If FlagFound = True Or Not FoundDate = "" Or Not FoundHour = "" Then
            WriteCFURpt i_Temp, ActivityType & " Activity - Hour Driven", FoundDate, FoundHour, InitBy, ActivityType
        Else
            WriteCFURpt i_Temp, ActivityType & " CFU - Manual Driven", "", "", "", ActivityType
        End If
    End If

End Sub

Sub WriteCFURpt(i_row As Integer, SortType As String, Optional FoundDateVal As String, Optional FoundHourVal As String, Optional InitByVal As Variant, Optional ActivityTypeVal As String)

    Dim DateString As Variant
    Dim InitByParentVal As Variant

    RSCFURpt.AddNew
    RSCFURpt![Tail No] = ArrTemp(0, i_row)
    RSCFURpt![Current_AFHR] = ArrTemp(1, i_row)
    RSCFURpt![Venue] = ArrTemp(2, i_row)
    RSCFURpt![TAIL/SERIAL NO] = ArrTemp(3, i_row)
    RSCFURpt!PART = ArrTemp(4, i_row)
    RSCFURpt!CAGE = ArrTemp(5, i_row)
    RSCFURpt![MAINT ACTIVITY] = ArrTemp(6, i_row)
    If Not FoundDateVal = "" Then
        If IsDate(FoundDateVal) Then
            RSCFURpt![DATE DUE] = FoundDateVal
        Else
            DateString = ConvertToDateSerial(FoundDateVal)
            RSCFURpt![DATE DUE] = DateString
        End If
    Else
        If Not ArrTemp(7, i_row) = "" Then
            DateString = ConvertToDateSerial(ArrTemp(7, i_row))
            RSCFURpt![DATE DUE] = DateString
        End If
    End If
    RSCFURpt![LIFE TYPE INT] = ArrTemp(8, i_row)
    If Not FoundHourVal = "" Then
        RSCFURpt![REM] = FoundHourVal
    Else
        RSCFURpt![REM] = ArrTemp(9, i_row)
    End If
    RSCFURpt![LCN] = ArrTemp(10, i_row)
    RSCFURpt![ALC] = ArrTemp(11, i_row)
    RSCFURpt![POSITION NAME] = ArrTemp(12, i_row)
    RSCFURpt![WAC] = ArrTemp(13, i_row)
    RSCFURpt![Type] = ArrTemp(14, i_row)
    RSCFURpt![LVL] = ArrTemp(15, i_row)
    RSCFURpt![STD] = ArrTemp(16, i_row)
    RSCFURpt![REC] = ArrTemp(17, i_row)
    RSCFURpt![CLM] = ArrTemp(18, i_row)
    RSCFURpt![DATE PLAN] = ArrTemp(19, i_row)
    RSCFURpt![INTERVAL] = ArrTemp(20, i_row)
    RSCFURpt![LIFE SINCE NEW] = ArrTemp(21, i_row)
    RSCFURpt![LIFE SINCE NEW TYPE] = ArrTemp(22, i_row)
    RSCFURpt![INIT TYPE] = ArrTemp(23, i_row)
    RSCFURpt![INSTANCE] = ArrTemp(24, i_row)
    RSCFURpt![INIT LCN] = ArrTemp(25, i_row)
    RSCFURpt![INIT ALC] = ArrTemp(26, i_row)
    If Not IsMissing(InitByVal) Then
        RSCFURpt![INIT BY] = InitByVal
    Else
        RSCFURpt![INIT BY] = ArrTemp(24, i_row)
    End If
    If Not IsMissing(InitByParentVal) Then
        RSCFURpt![INIT PARENT] = InitByParentVal
    End If
    RSCFURpt![SYMPTOMS] = ArrTemp(27, i_row)
    RSCFURpt![Displayed] = SortType
    RSCFURpt![RunDate] = ArrTemp(29, i_row)
    RSCFURpt.Update

End Sub

Sub SortDataSMR(ActivityType As String)
    
    i_Temp = i_row
    i = i_row
    ResetFlags
    
    Dim InitBy As Variant
    Dim FlagFound As Boolean
    
    If Not ArrTemp(24, i) = "OH" And Not ArrTemp(24, i) = "BS" Then
        'Check Activity
        For j = 0 To UBound(ArrActivity, 2)
            If ArrTemp(0, i) = ArrActivity(0, j) And ArrTemp(24, i) = ArrActivity(6, j) Then
                If (ArrTemp(25, i) = ArrActivity(14, j) And ArrTemp(26, i) = ArrActivity(15, j)) Or ArrTemp(25, i) = "" Then
                    If Not ArrActivity(7, j) = "" Then
                        FoundDate = ArrActivity(7, j)
                        InitParent = ArrActivity(16, j)
                        InitBy = ArrTemp(24, i)
                        InitType = ArrTemp(23, i)
                    ElseIf Right(ArrActivity(8, j), 2) = "HR" Then
                        FoundHour = ArrActivity(9, j)
                        InitParent = ArrActivity(16, j)
                        InitBy = ArrTemp(24, i)
                        InitType = ArrTemp(23, i)
                    End If
                    If Not FoundDate = "" And Not FoundHour = "" Then
                        FlagFound = True
                        Exit For
                    End If
                End If
            End If
        Next
        If FlagFound = True Or Not FoundDate = "" Or Not FoundHour = "" Then
            WriteSMRRpt i_Temp, ActivityType & " Activity - Hour Driven", FoundDate, FoundHour, InitBy, ActivityType
        Else
            WriteSMRRpt i_Temp, ActivityType & " CFU - Manual Driven", "", "", "", ActivityType
        End If
    End If

End Sub

Sub WriteSMRRpt(i_row As Integer, SortType As String, Optional FoundDateVal As String, Optional FoundHourVal As String, Optional InitByVal As Variant, Optional ActivityTypeVal As String)

    Dim DateString As Variant
    Dim InitByParentVal As Variant

    RSSMRRpt.AddNew
    RSSMRRpt![Tail No] = ArrTemp(0, i_row)
    RSSMRRpt![Current_AFHR] = ArrTemp(1, i_row)
    RSSMRRpt![Venue] = ArrTemp(2, i_row)
    RSSMRRpt![TAIL/SERIAL NO] = ArrTemp(3, i_row)
    RSSMRRpt!PART = ArrTemp(4, i_row)
    RSSMRRpt!CAGE = ArrTemp(5, i_row)
    RSSMRRpt![MAINT ACTIVITY] = ArrTemp(6, i_row)
    If Not FoundDateVal = "" Then
        If IsDate(FoundDateVal) Then
            RSSMRRpt![DATE DUE] = FoundDateVal
        Else
            DateString = ConvertToDateSerial(FoundDateVal)
            RSSMRRpt![DATE DUE] = DateString
        End If
    Else
        If Not ArrTemp(7, i_row) = "" Then
            DateString = ConvertToDateSerial(ArrTemp(7, i_row))
            RSSMRRpt![DATE DUE] = DateString
        End If
    End If
    RSSMRRpt![LIFE TYPE INT] = ArrTemp(8, i_row)
    If Not FoundHourVal = "" Then
        RSSMRRpt![REM] = FoundHourVal
    ElseIf Not Right(RSSMRRpt![LIFE TYPE INT], 2) = "HR" Then
        'RSSMRRpt![REM] = ""
    Else
        RSSMRRpt![REM] = ArrTemp(9, i_row)
    End If
    RSSMRRpt![LCN] = ArrTemp(10, i_row)
    RSSMRRpt![ALC] = ArrTemp(11, i_row)
    RSSMRRpt![POSITION NAME] = ArrTemp(12, i_row)
    RSSMRRpt![WAC] = ArrTemp(13, i_row)
    RSSMRRpt![Type] = ArrTemp(14, i_row)
    RSSMRRpt![LVL] = ArrTemp(15, i_row)
    RSSMRRpt![STD] = ArrTemp(16, i_row)
    RSSMRRpt![REC] = ArrTemp(17, i_row)
    RSSMRRpt![CLM] = ArrTemp(18, i_row)
    RSSMRRpt![DATE PLAN] = ArrTemp(19, i_row)
    RSSMRRpt![INTERVAL] = ArrTemp(20, i_row)
    RSSMRRpt![LIFE SINCE NEW] = ArrTemp(21, i_row)
    RSSMRRpt![LIFE SINCE NEW TYPE] = ArrTemp(22, i_row)
    RSSMRRpt![INIT TYPE] = ArrTemp(23, i_row)
    RSSMRRpt![INSTANCE] = ArrTemp(24, i_row)
    RSSMRRpt![INIT LCN] = ArrTemp(25, i_row)
    RSSMRRpt![INIT ALC] = ArrTemp(26, i_row)
    If Not IsMissing(InitByVal) Then
        RSSMRRpt![INIT BY] = InitByVal
    Else
        RSSMRRpt![INIT BY] = ArrTemp(24, i_row)
    End If
    If Not IsMissing(InitByParentVal) Then
        RSSMRRpt![INIT PARENT] = InitByParentVal
    End If
    RSSMRRpt![SYMPTOMS] = ArrTemp(27, i_row)
    RSSMRRpt![Displayed] = SortType
    RSSMRRpt![RunDate] = ArrTemp(29, i_row)
    RSSMRRpt.Update

End Sub

Function ConvertToDateSerial(DateToConvert As Variant)

    Dim YearVal As Variant
    Dim MonVal As Variant
    Dim DayVal As Variant

    YearVal = Right(DateToConvert, 4)
    MonVal = Month("01 " & Mid(DateToConvert, 3, 3) & " " & YearVal)
    DayVal = Left(DateToConvert, 2)
    ConvertToDateSerial = DateSerial(YearVal, MonVal, DayVal)
    ConvertToDateSerial = Format(ConvertToDateSerial, "dd-mmm-yyyy")

End Function

Private Sub txtDayLimit_AfterUpdate()

    Forms![frm_MFT].Form.Requery
    
    ApplyMFTFilter

End Sub

Private Sub txtEventLimit_AfterUpdate()

    Forms![frm_MFT].Form.Requery
    
    ApplyMFTFilter

End Sub

Private Sub txtHourLimit_AfterUpdate()

    Forms![frm_MFT].Form.Requery

    ApplyMFTFilter

End Sub

