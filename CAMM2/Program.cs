// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");

using Microsoft.VisualBasic;
using System.IO;
using System;
using CAMM2;
using System.Diagnostics;

new Test().Print();

public class Test
{
    string PrevLine = string.Empty;
    string CurrentLine = string.Empty;
    string NextLine = string.Empty;
    string NextNextLine = string.Empty;
    string RunDate = string.Empty;
    string ReportVenue = string.Empty;
    string Venue = string.Empty;
    string TailNo = string.Empty;
    List<CAMM2AssetDto> assets = new List<CAMM2AssetDto>();
    string[] ArrRSInput = new string[] { };
    decimal CurrentAFHR = 0;
    string MaintenanceType = string.Empty;
    int i = 0; int j = 0;
    string Tail_SerialNo = string.Empty;
    string PART = string.Empty;
    string CAGE = string.Empty;
    string ALC = string.Empty;
    string LCN = string.Empty;
    string Position_Name = string.Empty;
    string WAC = string.Empty;
    string Life_Since_New = string.Empty;
    string Life_Since_New_Type = string.Empty;
    public void Print()
    {
        string textFile = @"Files/20230123 79SQN.txt";
        ArrRSInput = File.ReadAllLines(textFile);

        for (i = 0; i < ArrRSInput.Length; i++)
        {
            if (i > 0)
            {
                PrevLine = ArrRSInput[i - 1];
            }

            // Get current line data
            CurrentLine = ArrRSInput[i];

            // Get next line data
            if (i < ArrRSInput.Length - 1)
            {
                NextLine = ArrRSInput[i + 1];
            }

            // Get current + 2 line data
            if (i < ArrRSInput.Length - 2)
            {
                NextNextLine = ArrRSInput[i + 2];
            }
            // Run Date
            if (CurrentLine.StartsWith("** STARTED DATE"))
            {
                RunDate = CurrentLine.SafeSubstring(CurrentLine.IndexOf("=") + 2, 8);
            }

            // Venue
            if (CurrentLine.StartsWith("Q19 Venue:") || CurrentLine.StartsWith("Q22 Venue:"))
            {
                ReportVenue = CurrentLine.SafeSubstring(16, 100).Trim();
                ReportVenue = ReportVenue.Substring(ReportVenue.IndexOf(":") + 1);
            }

            if (CurrentLine.StartsWith(" VENUE         :"))
            {
                Venue = CurrentLine.SafeSubstring(16, 50).Trim();
            }

            if (CurrentLine.SafeSubstring(78, 15).Contains("SERIAL NUMBER:"))
            {
                TailNo = CurrentLine.SafeSubstring(93, 8).Trim();
                if (!assets.Any(a => a.TailNo == TailNo))
                    assets.Add(new CAMM2AssetDto { Date = RunDate, TailNo = TailNo, Venue = Venue });
            }

            //Current AFHRS for Serial Number/Tail Number and Header Info
            j = 1;
            if (CurrentLine.SafeSubstring(0, 26).Contains("LIFE SINCE NEW:    LIFE"))
            {
                while (!string.IsNullOrEmpty(ArrRSInput[i + j].SafeSubstring(15, 12)))
                {
                    CurrentAFHR = Convert.ToDecimal(NextLine.SafeSubstring(15, 12).Trim());
                    GetHeaderInfo();
                    j++;
                }
            }
            // Check current Maintenance Type
            if (CurrentLine.StartsWith("-- FORECAST ACTIVITIES --"))
            {
                MaintenanceType = "Activities";
            }
            else if (CurrentLine.StartsWith("-- DEVIATIONS --"))
            {
                MaintenanceType = "Deviations";
            }
            else if (CurrentLine.Contains("-- COINCIDENT MAINTENANCE ACTIVITIES --"))
            {
                MaintenanceType = "Coincident";
            }
            else if (CurrentLine.Contains("-- CFUs --"))
            {
                MaintenanceType = "CFU";
            }
            else if (CurrentLine.Contains("-- JOBS IN PROGRESS --"))
            {
                MaintenanceType = "In Progress";
            }
            else if (CurrentLine.Contains("-- OVERDUE MAINTENANCE --"))
            {
                MaintenanceType = "Overdue";
            }
            else if (CurrentLine.Contains("-- SMRs --"))
            {
                MaintenanceType = "SMR";
            }

            // Get all Maint Activities
            if (MaintenanceType == "Activities")
            {
                string maintTypeSubstring = CurrentLine.SafeSubstring(63, 5).Trim();
                if (maintTypeSubstring == "MOD" || maintTypeSubstring == "STI" || maintTypeSubstring == "SM" || maintTypeSubstring == "DVI")
                {
                    GetActivityData();
                }
            }
            //Get Deviation data
            if (MaintenanceType == "Deviations")
            {
                if (CurrentLine.SafeSubstring(63, 12).Trim().Contains("DEV-"))
                {
                    GetDeviationData();
                }
                else if (CurrentLine.SafeSubstring(97, 3) == " Y ")
                {
                    GetDeviationData();
                }
                else if (CurrentLine.SafeSubstring(97, 3) == " N ")
                {
                    GetDeviationData();
                }
            }
            // Get Coincident data
            if (MaintenanceType == "Coincident")
            {
                string trimValue = CurrentLine.SafeSubstring(63, 5).Trim();
                if (trimValue == "MOD" || trimValue == "STI" || trimValue == "SM" || trimValue == "DVI")
                {
                    GetCoincidentData();
                }
            }
            // Get CFU data
            if (MaintenanceType == "CFU")
            {
                if (CurrentLine.SafeSubstring(63, 5).Trim() == "CFU")
                {
                    GetCFUData();
                }
            }

            // Get SMR data
            if (MaintenanceType == "SMR")
            {
                if (CurrentLine.SafeSubstring(63, 5).Trim() == "SMR")
                {
                    GetSMRData();
                }
            }

            // Get In Progress data
            if (MaintenanceType == "In Progress")
            {
                string trimValue = CurrentLine.SafeSubstring(72, 3).Trim();
                if (trimValue == "Y" || trimValue == "N")
                {
                    GetInProgressData();
                }
            }

            // Get Overdue Data
            if (MaintenanceType == "Overdue")
            {
                string trimValue = CurrentLine.SafeSubstring(72, 3).Trim();
                if (trimValue == "Y" || trimValue == "N")
                {
                    GetOverdueData();
                }
            }
        }
    }
    public void GetHeaderInfo()
    {
        var asset = assets.Find(a => a.TailNo == TailNo);

        var interval = ArrRSInput[i + j].SafeSubstring(47, 12);
        var life = ArrRSInput[i + j].SafeSubstring(15, 12);

        var assetHeader = new CAMM2AssetHeaderInfoDto();
        assetHeader.Life = string.IsNullOrEmpty(life) ? 0 : Convert.ToDecimal(ArrRSInput[i + j].SafeSubstring(15, 12));
        assetHeader.LifeType = ArrRSInput[i + j].SafeSubstring(31, 12).Trim();
        assetHeader.Interval = string.IsNullOrEmpty(interval) ? 0 : Convert.ToDecimal(ArrRSInput[i + j].SafeSubstring(47, 12));
        asset.HeaderInfo.Add(assetHeader);
    }
    void GetActivityData()
    {
        var asset = assets.Find(a => a.TailNo == TailNo);
        // Get Full line of data from CAMM2 if it exists
        if (CurrentLine.SafeSubstring(0, 63) != "                                                               ")
        {
            Tail_SerialNo = PrevLine.SafeSubstring(0, 23).Trim();
            PART = PrevLine.SafeSubstring(23, 30).Trim();
            CAGE = PrevLine.SafeSubstring(55, 7).SafeSubstring(1, PrevLine.SafeSubstring(55, 7).Trim().Length);
            LCN = CurrentLine.SafeSubstring(0, 18).Trim();
            ALC = CurrentLine.SafeSubstring(18, 3).Trim();
            Position_Name = CurrentLine.SafeSubstring(23, 19).Trim();
            WAC = CurrentLine.SafeSubstring(43, 5).Trim();
            Life_Since_New = NextLine.SafeSubstring(43, 9).Trim();
            Life_Since_New_Type = NextLine.SafeSubstring(53, 7).Trim();
        }

        // Link remaining data to the previous full line of data
        var Maint_Activity = PrevLine.SafeSubstring(63, 21).Trim();
        var Date_due = PrevLine.SafeSubstring(85, 10).Trim();
        var Life_Type_Int = PrevLine.SafeSubstring(95, 10).Trim();
        var Int_Rem = PrevLine.SafeSubstring(104, 10).Trim();
        var TSSCode1 = PrevLine.SafeSubstring(120, 10).Trim();
        var TSSCode2 = CurrentLine.SafeSubstring(120, 10).Trim();
        var TSSCode3 = NextLine.SafeSubstring(120, 10).Trim();
        var TSSCode4 = NextLine.SafeSubstring(120, 10).Trim();
        var Maint_Type = CurrentLine.SafeSubstring(63, 4).Trim();
        var LVL = CurrentLine.SafeSubstring(68, 3).Trim();
        var STD = CurrentLine.SafeSubstring(73, 2).Trim();
        var REC = CurrentLine.SafeSubstring(77, 3).Trim();
        var CLM = CurrentLine.SafeSubstring(81, 3).Trim();
        var Date_Plan = CurrentLine.SafeSubstring(85, 10).Trim();
        var INTERVAL = CurrentLine.SafeSubstring(95, 10).Trim();
        var Init_Type = NextLine.SafeSubstring(63, 5).Trim();
        var INSTANCE = NextLine.SafeSubstring(73, 10).Trim();
        var Init_LCN = NextLine.SafeSubstring(95, 11).Trim();
        var Init_ALC = NextLine.SafeSubstring(113, 4).Trim();

        // Write Maint Activity data to table
        var activity = new CAMM2ActivityDto();
        activity.CurrentAFHR = CurrentAFHR;
        activity.TailSerialNo = Tail_SerialNo;
        activity.Part = PART;
        activity.Cage = CAGE;
        activity.MaintActivity = Maint_Activity;
        activity.DueDate = Date_due;
        activity.LifeTypeInt = Life_Type_Int;
        activity.REM = Int_Rem;
        activity.TSSCode1 = TSSCode1;
        activity.TSSCode2 = TSSCode2;
        activity.TSSCode3 = TSSCode3;
        activity.TSSCode4 = TSSCode4;
        activity.LCN = LCN;
        activity.ALC = ALC;
        activity.PositionName = Position_Name;
        activity.WAC = WAC;
        activity.Type = Maint_Type;
        activity.LVL = LVL;
        activity.STD = STD;
        activity.REC = REC;
        activity.CLM = CLM;
        activity.DatePlan = Date_Plan;
        activity.Interval = INTERVAL;
        activity.LifeSinceNew = Life_Since_New;
        activity.LifeSinceNewType = Life_Since_New_Type;
        activity.InitType = Init_Type;
        activity.Instance = INSTANCE;
        activity.InitLCN = Init_LCN;
        activity.InitALC = Init_ALC;
        asset.Activities.Add(activity);
    }
    void GetDeviationData()
    {
        var Deviation_Order = string.Empty;
        var STD = string.Empty;
        var EXPIRED = string.Empty;
        var TITLE = string.Empty;
        var Expiry_Conditions = string.Empty;
        var LIMITATION = string.Empty;
        var ACTIVITY = string.Empty;
        var Dev_Type = string.Empty;
        var Init_LCN = string.Empty;
        var Init_ALC = string.Empty;
        var INTERVAL = string.Empty;
        var Int_Rem = string.Empty;
        var Date_due = string.Empty;
        var DateString = string.Empty;

        // Get Full line of data from CAMM2 if it exists
        if (CurrentLine.SafeSubstring(0, 63) != "                                                               ")
        {
            Tail_SerialNo = CurrentLine.SafeSubstring(0, 23).Trim();
            PART = CurrentLine.SafeSubstring(23, 30);
            CAGE = CurrentLine.SafeSubstring(55, 7).SafeSubstring(1, CurrentLine.SafeSubstring(55, 7).Trim().Length);
            LCN = NextLine.SafeSubstring(0, 18).Trim();
            ALC = NextLine.SafeSubstring(18, 3).Trim();
            Position_Name = NextLine.SafeSubstring(23, 19).Trim();
            WAC = NextLine.SafeSubstring(43, 5).Trim();
            Life_Since_New = NextNextLine.SafeSubstring(43, 9).Trim();
            Life_Since_New_Type = NextNextLine.SafeSubstring(53, 7).Trim();
        }

        Deviation_Order = CurrentLine.SafeSubstring(63, 30).Trim();
        STD = CurrentLine.SafeSubstring(97, 5).Trim();
        EXPIRED = CurrentLine.SafeSubstring(116, 5).Trim();
        TITLE = NextLine.SafeSubstring(63, 50).Trim();
        Expiry_Conditions = string.Empty;
        if (ArrRSInput[i + 3].SafeSubstring(63, 31) == "*** NO EXPIRY CONDITIONS ***")
        {
            Expiry_Conditions = "NO EXPIRY CONDITIONS";
        }
        else
        {
            Expiry_Conditions = "";
        }

        // Collect all limitation data including data that crosses onto another page
        LIMITATION = "";
        if (ArrRSInput[i + 7].Contains("LIMITATION:"))
        {
            int j = 7;
            do
            {
                if (ArrRSInput[i + j].SafeSubstring(55, 12) == "For-Official" && ArrRSInput[(i + j) + 10].SafeSubstring(63, 7) != "(cont)")
                {
                    break;
                }

                if (ArrRSInput[(i + j)].SafeSubstring(63, 12).Trim().Contains("DEV-"))
                {
                    break;
                }

                if (ArrRSInput[(i + j) + 11].SafeSubstring(63, 7).Trim() == "(cont)")
                {
                    j = j + 11;
                }

                if (ArrRSInput[i + j].SafeSubstring(75, 55).Trim() != "")
                {
                    LIMITATION = LIMITATION + ArrRSInput[i + j].SafeSubstring(75, 55).Trim() + "\n";
                }

                if (ArrRSInput[i + j].StartsWith("-- "))
                {
                    break;
                }

                j = j + 1;
            } while (true);
        }

        // Get all the deviation order details; can include multiple for one deviation
        j = 5;
        do
        {
            if (ArrRSInput[i + 3].SafeSubstring(63, 31).Trim() != "*** NO EXPIRY CONDITIONS ***")
            {
                Dev_Type = ArrRSInput[i + j].SafeSubstring(63, 7).Trim();
                ACTIVITY = ArrRSInput[i + j].SafeSubstring(71, 15).Trim();
                Init_LCN = ArrRSInput[i + j].SafeSubstring(97, 2).Trim();

                if (ArrRSInput[i + j].SafeSubstring(110, 9).Trim().Length < 3)
                {
                    Init_ALC = ArrRSInput[i + j].SafeSubstring(117, 2).Trim();
                }
                else
                {
                    Init_ALC = "";
                }

                INTERVAL = ArrRSInput[i + j].SafeSubstring(88, 8).Trim();
                Int_Rem = ArrRSInput[i + j].SafeSubstring(102, 7).Trim();

                if (ArrRSInput[i + j].SafeSubstring(110, 9).Trim().Length > 2)
                {
                    Date_due = ArrRSInput[i + j].SafeSubstring(110, 9).Trim();
                }
                else
                {
                    Date_due = "";
                }
            }
            else
            {
                Dev_Type = "";
                ACTIVITY = "";
                Init_LCN = "";
                Init_ALC = "";
                INTERVAL = "";
                Int_Rem = "";
                Date_due = "";
            }
            var asset = assets.Find(a => a.TailNo == TailNo);
            var deviation = new CAMM2DeviationDto();

            // Write all deviation data to table
            deviation.TailNo = TailNo;
            deviation.CurrentAFHR = CurrentAFHR;
            deviation.Venue = Venue;
            deviation.TailSerialNo = Tail_SerialNo;
            deviation.Part = PART;
            deviation.Cage = CAGE;
            deviation.DeviationOrder = Deviation_Order;
            deviation.STD = STD;
            deviation.Expired = EXPIRED;
            deviation.LCN = LCN;
            deviation.ALC = ALC;
            deviation.PositionName = Position_Name;
            deviation.WAC = WAC;
            deviation.Title = TITLE;
            deviation.LifeSinceNew = Life_Since_New;
            deviation.LifeSinceNewType = Life_Since_New_Type;
            deviation.ExpiryConditions = Expiry_Conditions;
            deviation.Limitaion = LIMITATION;
            deviation.Type = Dev_Type;
            deviation.Activity = ACTIVITY;
            deviation.InitLCN = Init_LCN;
            deviation.InitALC = Init_ALC;
            deviation.Interval = INTERVAL;
            deviation.IntRem = Int_Rem;
            deviation.DueDate = Date_due;
            //if (!string.IsNullOrEmpty(Int_Rem.ToString()) && IsNumeric(Int_Rem))
            //{
            //    deviation.REM = Int_Rem;
            //}

            //if (!string.IsNullOrEmpty(Date_due.ToString()))
            //{
            //    DateString = ConvertToDateSerial(Date_due.ToString());
            //    deviation.DATE DUE = DateString;
            //}

            deviation.RunDate = RunDate;
            asset.Deviations.Add(deviation);

            // Exit loop if the next line is a blank space
            j = j + 1;
            if (string.IsNullOrEmpty(ArrRSInput[i + j].SafeSubstring(63, 7).Trim()) || ArrRSInput[i + 3].SafeSubstring(63, 31).Trim() == "*** NO EXPIRY CONDITIONS ***")
            {
                break;
            }
        } while (true);
    }

    void GetCoincidentData()
    {
        var Maint_Activity = string.Empty;
        var Date_due = string.Empty;
        var Life_Type_Int = string.Empty;
        var Int_Rem = string.Empty;
        var TSSCode1 = string.Empty;
        var TSSCode2 = string.Empty;
        var TSSCode3 = string.Empty;
        var TSSCode4 = string.Empty;
        var Maint_Type = string.Empty;
        var LVL = string.Empty;
        var STD = string.Empty;
        var REC = string.Empty;
        var CLM = string.Empty;
        var Date_Plan = string.Empty;
        var INTERVAL = string.Empty;
        var Init_Type = string.Empty;
        var INSTANCE = string.Empty;
        var Init_LCN = string.Empty;
        var Init_ALC = string.Empty;
        var DateString = string.Empty;

        // Get Full line of data from CAMM2 if it exists
        if (CurrentLine.SafeSubstring(0, 63) != "                                                               ")
        {
            Tail_SerialNo = PrevLine.SafeSubstring(0, 23).Trim();
            PART = PrevLine.SafeSubstring(23, 30).Trim();
            CAGE = PrevLine.SafeSubstring(55, 7).Trim().SafeSubstring(1, PrevLine.SafeSubstring(55, 7).Trim().Length);
            LCN = CurrentLine.SafeSubstring(0, 18).Trim();
            ALC = CurrentLine.SafeSubstring(19, 3).Trim();
            Position_Name = CurrentLine.SafeSubstring(24, 19).Trim();
            WAC = CurrentLine.SafeSubstring(43, 5).Trim();
            Life_Since_New = NextLine.SafeSubstring(43, 9).Trim();
            Life_Since_New_Type = NextLine.SafeSubstring(53, 7).Trim();
        }

        // Link remaining data to the previous full line of data
        Maint_Activity = PrevLine.SafeSubstring(63, 21).Trim();
        Date_due = PrevLine.SafeSubstring(85, 10).Trim();
        Life_Type_Int = PrevLine.SafeSubstring(95, 10).Trim();
        Int_Rem = PrevLine.SafeSubstring(104, 10).Trim();
        TSSCode1 = PrevLine.SafeSubstring(120, 10).Trim();
        TSSCode2 = CurrentLine.SafeSubstring(120, 10).Trim();
        TSSCode3 = NextLine.SafeSubstring(120, 10).Trim();
        TSSCode4 = NextLine.SafeSubstring(120, 10).Trim();
        Maint_Type = CurrentLine.SafeSubstring(63, 4).Trim();
        LVL = CurrentLine.SafeSubstring(68, 3).Trim();
        STD = CurrentLine.SafeSubstring(73, 2).Trim();
        REC = CurrentLine.SafeSubstring(77, 3).Trim();
        CLM = CurrentLine.SafeSubstring(81, 3).Trim();
        Date_Plan = CurrentLine.SafeSubstring(85, 10).Trim();
        INTERVAL = CurrentLine.SafeSubstring(95, 10).Trim();
        Init_Type = NextLine.SafeSubstring(63, 5).Trim();
        INSTANCE = NextLine.SafeSubstring(73, 17).Trim();
        Init_LCN = NextLine.SafeSubstring(95, 11).Trim();
        Init_ALC = NextLine.SafeSubstring(113, 4).Trim();

        // Write Maint Activity data to table
        var asset = assets.Find(a => a.TailNo == TailNo);
        var coincident = new CAMM2CoincidentDto();

        coincident.TailNo = TailNo;
        coincident.CurrentAFHR = CurrentAFHR;
        coincident.Venue = Venue;
        coincident.TailSerialNo = Tail_SerialNo;
        coincident.Part = PART;
        coincident.Cage = CAGE;
        coincident.MaintActivity = Maint_Activity;

        //if (!string.IsNullOrEmpty(Date_due.ToString()))
        //{
        //    DateString = ConvertToDateSerial(Date_due.ToString());
        //    coincident.DATE DUE = DateString;
        //}

        coincident.LifeTypeInt = Life_Type_Int;

        //if (!string.IsNullOrEmpty(Int_Rem.ToString()))
        //{
        //    coincident.REM = Int_Rem;
        //}
        coincident.IntRem = Int_Rem;
        coincident.DueDate = Date_due;
        coincident.TSSCode1 = TSSCode1;
        coincident.TSSCode2 = TSSCode2;
        coincident.TSSCode3 = TSSCode3;
        coincident.TSSCode4 = TSSCode4;
        coincident.LCN = LCN;
        coincident.ALC = ALC;
        coincident.PositionName = Position_Name;
        coincident.WAC = WAC;
        coincident.Type = Maint_Type;
        coincident.LVL = LVL;
        coincident.STD = STD;
        coincident.REC = REC;
        coincident.CLM = CLM;
        coincident.DatePlan = Date_Plan;
        coincident.Interval = INTERVAL;
        coincident.LifeSinceNew = Life_Since_New;
        coincident.LifeSinceNewType = Life_Since_New_Type;
        coincident.InitType = Init_Type;
        coincident.Instance = INSTANCE;
        coincident.InitLCN = Init_LCN;
        coincident.InitALC = Init_ALC;
        coincident.RunDate = RunDate;
        asset.Coincidents.Add(coincident);
    }

    void GetCFUData()
    {
        var Maint_Activity = string.Empty;
        var Date_due = string.Empty;
        var Life_Type_Int = string.Empty;
        var Int_Rem = string.Empty;
        var TSSCode1 = string.Empty;
        var TSSCode2 = string.Empty;
        var TSSCode3 = string.Empty;
        var TSSCode4 = string.Empty;
        var Maint_Type = string.Empty;
        var LVL = string.Empty;
        var STD = string.Empty;
        var REC = string.Empty;
        var CLM = string.Empty;
        var Date_Plan = string.Empty;
        var INTERVAL = string.Empty;
        var Init_Type = string.Empty;
        var INSTANCE = string.Empty;
        var Init_LCN = string.Empty;
        var Init_ALC = string.Empty;
        var DateString = string.Empty;
        var SYMPTOMS = string.Empty;

        // Get Full line of data from CAMM2 if it exists
        if (CurrentLine.SafeSubstring(0, 63) != "                                                               ")
        {
            Tail_SerialNo = PrevLine.SafeSubstring(0, 23);
            PART = PrevLine.SafeSubstring(23, 30);
            CAGE = PrevLine.SafeSubstring(55, 7).Trim().SafeSubstring(1, PrevLine.SafeSubstring(55, 7).Trim().Length);
            LCN = CurrentLine.SafeSubstring(0, 18);
            ALC = CurrentLine.SafeSubstring(18, 3);
            Position_Name = CurrentLine.SafeSubstring(23, 19);
            WAC = CurrentLine.SafeSubstring(43, 5);
            Life_Since_New = NextLine.SafeSubstring(43, 9);
            Life_Since_New_Type = NextLine.SafeSubstring(53, 7);
        }

        // Link remaining data to the previous full line of data
        Maint_Activity = PrevLine.SafeSubstring(63, 21);
        Date_due = PrevLine.SafeSubstring(85, 10);
        Life_Type_Int = PrevLine.SafeSubstring(95, 10);
        Int_Rem = PrevLine.SafeSubstring(104, 10);
        Maint_Type = CurrentLine.SafeSubstring(63, 4);
        LVL = CurrentLine.SafeSubstring(68, 3);
        STD = CurrentLine.SafeSubstring(73, 2);
        REC = CurrentLine.SafeSubstring(77, 3);
        CLM = CurrentLine.SafeSubstring(81, 3);
        Date_Plan = CurrentLine.SafeSubstring(85, 10);
        INTERVAL = CurrentLine.SafeSubstring(95, 10);
        Init_Type = NextLine.SafeSubstring(63, 5);
        INSTANCE = NextLine.SafeSubstring(73, 10);
        Init_LCN = NextLine.SafeSubstring(95, 11);
        Init_ALC = NextLine.SafeSubstring(113, 4);

        SYMPTOMS = "";
        if (ArrRSInput[i + 2].Contains("SYMPTOMS:") || ArrRSInput[i + 3].Contains("SYMPTOMS:"))
        {
            j = 2;
            do
            {
                if (ArrRSInput[i + j].SafeSubstring(55, 12) == "For-Official" && ArrRSInput[(i + j) + 10].SafeSubstring(63, 7).Trim() != "(cont)")
                {
                    break;
                }
                if (ArrRSInput[((i + j) + 1)].SafeSubstring(63, 9).Trim().Contains("CFU"))
                {
                    break;
                }
                if (((i + j) + 11) <= ArrRSInput.GetUpperBound(1))
                {
                    if (ArrRSInput[(i + j) + 11].SafeSubstring(63, 7).Trim() == "(cont)")
                    {
                        j = j + 11;
                    }
                }
                else
                {
                    break;
                }
                if (ArrRSInput[i + j].SafeSubstring(73, 57).Trim() != "")
                {
                    SYMPTOMS = SYMPTOMS + ArrRSInput[i + j].SafeSubstring(73, 57).Trim() + "\n";
                }
                if (ArrRSInput[i + j].SafeSubstring(0, 3) == "-- ")
                {
                    break;
                }
                j++;
            } while (true);
        }

        // Write Maint Activity data to table
        var asset = assets.Find(a => a.TailNo == TailNo);
        var cfu = new CAMM2CFUDto();
        cfu.TailNo = TailNo;
        cfu.CurrentAFHR = CurrentAFHR;
        cfu.Venue = Venue;
        cfu.TailSerialNo = Tail_SerialNo;
        cfu.Part = PART;
        cfu.Cage = CAGE;
        cfu.MaintActivity = Maint_Activity;
        //if (!Date_due.Equals("") && Date_due != null)
        //{
        //    DateString = ConvertToDateSerial(Date_due.ToString());
        //    cfu.DATE DUE = DateString;
        //}
        cfu.LifeTypeInt = Life_Type_Int;
        //if (!Int_Rem.Equals("") && IsNumeric(Int_Rem))
        //{
        //    cfu.REM = Int_Rem;
        //}
        cfu.IntRem = Int_Rem;
        cfu.DueDate = Date_due;
        cfu.LCN = LCN;
        cfu.ALC = ALC;
        cfu.PositionName = Position_Name;
        cfu.WAC = WAC;
        cfu.Type = Maint_Type;
        cfu.LVL = LVL;
        cfu.STD = STD;
        cfu.REC = REC;
        cfu.CLM = CLM;
        cfu.DatePlan = Date_Plan;
        cfu.Interval = INTERVAL;
        cfu.LifeSinceNew = Life_Since_New;
        cfu.LifeSinceNewType = Life_Since_New_Type;
        cfu.InitType = Init_Type;
        cfu.Instance = INSTANCE;
        cfu.InitLCN = Init_LCN;
        cfu.InitALC = Init_ALC;
        cfu.Symptoms = SYMPTOMS;
        cfu.RunDate = RunDate;
        asset.CFU.Add(cfu);
    }

    void GetSMRData()
    {
        var Maint_Activity = string.Empty;
        var Date_due = string.Empty;
        var Life_Type_Int = string.Empty;
        var Int_Rem = string.Empty;
        var TSSCode1 = string.Empty;
        var TSSCode2 = string.Empty;
        var TSSCode3 = string.Empty;
        var TSSCode4 = string.Empty;
        var Maint_Type = string.Empty;
        var LVL = string.Empty;
        var STD = string.Empty;
        var REC = string.Empty;
        var CLM = string.Empty;
        var Date_Plan = string.Empty;
        var INTERVAL = string.Empty;
        var Init_Type = string.Empty;
        var INSTANCE = string.Empty;
        var Init_LCN = string.Empty;
        var Init_ALC = string.Empty;
        var DateString = string.Empty;
        var SYMPTOMS = string.Empty;

        // Get Full line of data from CAMM2 if it exists
        if (CurrentLine.SafeSubstring(0, 63) != "                                                               ")
        {
            Tail_SerialNo = PrevLine.SafeSubstring(0, 23).Trim();
            PART = PrevLine.SafeSubstring(23, 30).Trim();
            CAGE = PrevLine.SafeSubstring(55, 7).Trim().SafeSubstring(1, PrevLine.SafeSubstring(55, 7).Trim().Length);
            LCN = CurrentLine.SafeSubstring(0, 18).Trim();
            ALC = CurrentLine.SafeSubstring(18, 3).Trim();
            Position_Name = CurrentLine.SafeSubstring(23, 19).Trim();
            WAC = CurrentLine.SafeSubstring(43, 5).Trim();
            Life_Since_New = NextLine.SafeSubstring(43, 9).Trim();
            Life_Since_New_Type = NextLine.SafeSubstring(53, 7).Trim();
        }

        // Link remaining data to the previous full line of data
        Maint_Activity = PrevLine.SafeSubstring(63, 21).Trim();
        Date_due = PrevLine.SafeSubstring(85, 10).Trim();
        Life_Type_Int = PrevLine.SafeSubstring(95, 10).Trim();
        Int_Rem = PrevLine.SafeSubstring(104, 10).Trim();
        Maint_Type = CurrentLine.SafeSubstring(63, 4).Trim();
        LVL = CurrentLine.SafeSubstring(68, 3).Trim();
        STD = CurrentLine.SafeSubstring(73, 2).Trim();
        REC = CurrentLine.SafeSubstring(77, 3).Trim();
        CLM = CurrentLine.SafeSubstring(81, 3).Trim();
        Date_Plan = CurrentLine.SafeSubstring(85, 10).Trim();
        INTERVAL = CurrentLine.SafeSubstring(95, 10).Trim();
        Init_Type = NextLine.SafeSubstring(63, 5).Trim();
        INSTANCE = NextLine.SafeSubstring(73, 10).Trim();
        Init_LCN = NextLine.SafeSubstring(95, 11).Trim();
        Init_ALC = NextLine.SafeSubstring(113, 4).Trim();

        SYMPTOMS = "";
        if (ArrRSInput[i + 2].Contains("SYMPTOMS:") || ArrRSInput[i + 3].Contains("SYMPTOMS:"))
        {
            j = 2;
            do
            {
                if (ArrRSInput[i + j].SafeSubstring(55, 12).Trim() == "For-Official" && ArrRSInput[i + j + 10].SafeSubstring(63, 7).Trim() != "(cont)")
                {
                    break;
                }

                if (ArrRSInput[i + j + 1].SafeSubstring(63, 9).Trim().Contains("SMR"))
                {
                    break;
                }

                if (i + j + 11 <= ArrRSInput.GetUpperBound(1))
                {
                    if (ArrRSInput[i + j + 11].SafeSubstring(63, 7).Trim() == "(cont)")
                    {
                        j = j + 11;
                    }
                }
                else
                {
                    break;
                }

                if (ArrRSInput[i + j].SafeSubstring(73, 57).Trim() != "")
                {
                    SYMPTOMS += ArrRSInput[i + j].SafeSubstring(73, 57).Trim() + "\n";
                }

                if (ArrRSInput[i + j].SafeSubstring(0, 3) == "-- ")
                {
                    break;
                }

                j++;
            } while (true);
        }

        // Write Maint Activity data to table
        //RSSMR.AddNew();
        var asset = assets.Find(a => a.TailNo == TailNo);
        var smr = new CAMM2SMRDto();
        smr.TailNo = TailNo;
        smr.CurrentAFHR = CurrentAFHR;
        smr.Venue = Venue;
        smr.TailSerialNo = Tail_SerialNo;
        smr.Part = PART;
        smr.Cage = CAGE;
        smr.MaintActivity = Maint_Activity;
        //if (!string.IsNullOrEmpty(Date_due))
        //{
        //    DateString = ConvertToDateSerial(Date_due);
        //    smr.DATE DUE = DateString;
        //}
        smr.LifeTypeInt = Life_Type_Int;
        //if (!string.IsNullOrEmpty(Int_Rem) && double.TryParse(Int_Rem, out double IntRemValue))
        //{
        //    smr.REM = IntRemValue;
        //}
        smr.IntRem = Int_Rem;
        smr.DueDate = Date_due;
        smr.LCN = LCN;
        smr.ALC = ALC;
        smr.PositionName = Position_Name;
        smr.WAC = WAC;
        smr.Type = Maint_Type;
        smr.LVL = LVL;
        smr.STD = STD;
        smr.REC = REC;
        smr.CLM = CLM;
        smr.DatePlan = Date_Plan;
        smr.Interval = INTERVAL;
        smr.LifeSinceNew = Life_Since_New;
        smr.LifeSinceNewType = Life_Since_New_Type;
        smr.InitType = Init_Type;
        smr.Instance = INSTANCE;
        smr.InitLCN = Init_LCN;
        smr.InitALC = Init_ALC;
        smr.Symptoms = SYMPTOMS;
        smr.RunDate = RunDate;
        asset.SMR.Add(smr);
    }

    void GetInProgressData()
    {
        var Maint_Activity = string.Empty;
        var Date_due = string.Empty;
        var Life_Type_Int = string.Empty;
        var Int_Rem = string.Empty;
        var TSSCode1 = string.Empty;
        var TSSCode2 = string.Empty;
        var TSSCode3 = string.Empty;
        var TSSCode4 = string.Empty;
        var Maint_Type = string.Empty;
        var LVL = string.Empty;
        var STD = string.Empty;
        var REC = string.Empty;
        var CLM = string.Empty;
        var Date_Plan = string.Empty;
        var INTERVAL = string.Empty;
        var Init_Type = string.Empty;
        var INSTANCE = string.Empty;
        var Init_LCN = string.Empty;
        var Init_ALC = string.Empty;
        var DateString = string.Empty;

        // Get Full line of data from CAMM2 if it exists
        if (CurrentLine.SafeSubstring(0, 63) != "                                                               ")
        {
            Tail_SerialNo = PrevLine.SafeSubstring(0, 23).Trim();
            PART = PrevLine.SafeSubstring(23, 30).Trim();
            CAGE = PrevLine.SafeSubstring(55, 7).Trim().SafeSubstring(1, PrevLine.SafeSubstring(55, 7).Trim().Length);
            LCN = CurrentLine.SafeSubstring(0, 18).Trim();
            ALC = CurrentLine.SafeSubstring(18, 3).Trim();
            Position_Name = CurrentLine.SafeSubstring(23, 19).Trim();
            WAC = CurrentLine.SafeSubstring(43, 5).Trim();
            Life_Since_New = NextLine.SafeSubstring(43, 9).Trim();
            Life_Since_New_Type = NextLine.SafeSubstring(53, 7).Trim();
        }

        // Link remaining data to the previous full line of data
        Maint_Activity = PrevLine.SafeSubstring(63, 21).Trim();
        Date_due = PrevLine.SafeSubstring(85, 10).Trim();
        Life_Type_Int = PrevLine.SafeSubstring(95, 10).Trim();
        Int_Rem = PrevLine.SafeSubstring(104, 10).Trim();
        TSSCode1 = PrevLine.SafeSubstring(120, 10).Trim();
        TSSCode2 = CurrentLine.SafeSubstring(120, 10).Trim();
        TSSCode3 = NextLine.SafeSubstring(120, 10).Trim();
        TSSCode4 = NextLine.SafeSubstring(120, 10).Trim();
        Maint_Type = CurrentLine.SafeSubstring(63, 4).Trim();
        LVL = CurrentLine.SafeSubstring(68, 3).Trim();
        STD = CurrentLine.SafeSubstring(73, 2).Trim();
        REC = CurrentLine.SafeSubstring(77, 3).Trim();
        CLM = CurrentLine.SafeSubstring(81, 3).Trim();
        Date_Plan = CurrentLine.SafeSubstring(85, 10).Trim();
        INTERVAL = CurrentLine.SafeSubstring(95, 10).Trim();
        Init_Type = NextLine.SafeSubstring(63, 5).Trim();
        INSTANCE = NextLine.SafeSubstring(73, 10).Trim();
        Init_LCN = NextLine.SafeSubstring(95, 11).Trim();
        Init_ALC = NextLine.SafeSubstring(113, 4).Trim();

        var asset = assets.Find(a => a.TailNo == TailNo);
        var inProgress = new CAMM2InProgressDto();
        // Write Maint Activity data to table
        inProgress.TailNo = TailNo;
        inProgress.CurrentAFHR = CurrentAFHR;
        inProgress.Venue = Venue;
        inProgress.TailSerialNo = Tail_SerialNo;
        inProgress.Part = PART;
        inProgress.Cage = CAGE;
        inProgress.MaintActivity = Maint_Activity;
        //if (!string.IsNullOrEmpty(Date_due))
        //{
        //    DateString = ConvertToDateSerial(Date_due);
        //    inProgress.DATE DUE = DateString;
        //}
        inProgress.LifeTypeInt = Life_Type_Int;
        //if (!string.IsNullOrEmpty(Int_Rem) && double.TryParse(Int_Rem, out double IntRemValue))
        //{
        //    inProgress.REM = IntRemValue;
        //}
        inProgress.IntRem = Int_Rem;
        inProgress.DueDate = Date_due;
        inProgress.TSSCode1 = TSSCode1;
        inProgress.TSSCode2 = TSSCode2;
        inProgress.TSSCode3 = TSSCode3;
        inProgress.TSSCode4 = TSSCode4;
        inProgress.LCN = LCN;
        inProgress.ALC = ALC;
        inProgress.PositionName = Position_Name;
        inProgress.WAC = WAC;
        inProgress.Type = Maint_Type;
        inProgress.LVL = LVL;
        inProgress.STD = STD;
        inProgress.REC = REC;
        inProgress.CLM = CLM;
        inProgress.DatePlan = Date_Plan;
        inProgress.Interval = INTERVAL;
        inProgress.LifeSinceNew = Life_Since_New;
        inProgress.LifeSinceNewType = Life_Since_New_Type;
        inProgress.InitType = Init_Type;
        inProgress.Instance = INSTANCE;
        inProgress.InitLCN = Init_LCN;
        inProgress.InitALC = Init_ALC;
        inProgress.RunDate = RunDate;
        asset.InProgress.Add(inProgress);
    }
    void GetOverdueData()
    {
        string Maint_Activity = "";
        string Date_due = "";
        string Life_Type_Int = "";
        string Int_Rem = "";
        string TSSCode1 = "";
        string TSSCode2 = "";
        string TSSCode3 = "";
        string TSSCode4 = "";
        string Maint_Type = "";
        string LVL = "";
        string STD = "";
        string REC = "";
        string CLM = "";
        string Date_Plan = "";
        string INTERVAL = "";
        string Init_Type = "";
        string INSTANCE = "";
        string Init_LCN = "";
        string Init_ALC = "";
        string DateString = "";

        // Get Full line of data from CAMM2 if it exists
        if (CurrentLine.SafeSubstring(0, 63) != "                                                               ")
        {
            Tail_SerialNo = PrevLine.SafeSubstring(0, 23).Trim();
            PART = PrevLine.SafeSubstring(23, 30).Trim();
            CAGE = PrevLine.SafeSubstring(55, 7).Trim().SafeSubstring(1, PrevLine.SafeSubstring(55, 7).Trim().Length);
            LCN = CurrentLine.SafeSubstring(0, 18).Trim();
            ALC = CurrentLine.SafeSubstring(18, 3).Trim();
            Position_Name = CurrentLine.SafeSubstring(23, 19).Trim();
            WAC = CurrentLine.SafeSubstring(43, 5).Trim();
            Life_Since_New = NextLine.SafeSubstring(43, 9).Trim();
            Life_Since_New_Type = NextLine.SafeSubstring(53, 7).Trim();
        }

        // Link remaining data to the previous full line of data
        Maint_Activity = PrevLine.SafeSubstring(63, 21).Trim();
        Date_due = PrevLine.SafeSubstring(85, 10).Trim();
        Life_Type_Int = PrevLine.SafeSubstring(95, 10).Trim();
        Int_Rem = PrevLine.SafeSubstring(104, 10).Trim();
        TSSCode1 = PrevLine.SafeSubstring(120, 10).Trim();
        TSSCode2 = CurrentLine.SafeSubstring(120, 10).Trim();
        TSSCode3 = NextLine.SafeSubstring(120, 10).Trim();
        TSSCode4 = NextLine.SafeSubstring(120, 10).Trim();
        Maint_Type = CurrentLine.SafeSubstring(63, 4).Trim();
        LVL = CurrentLine.SafeSubstring(68, 3).Trim();
        STD = CurrentLine.SafeSubstring(73, 2).Trim();
        REC = CurrentLine.SafeSubstring(77, 3).Trim();
        CLM = CurrentLine.SafeSubstring(81, 3).Trim();
        Date_Plan = CurrentLine.SafeSubstring(85, 10).Trim();
        INTERVAL = CurrentLine.SafeSubstring(95, 10).Trim();
        Init_Type = NextLine.SafeSubstring(63, 5).Trim();
        INSTANCE = NextLine.SafeSubstring(73, 18).Trim();
        Init_LCN = NextLine.SafeSubstring(95, 11).Trim();
        Init_ALC = NextLine.SafeSubstring(113, 4).Trim();

        // Write Maint Activity data to table
        var asset = assets.Find(a => a.TailNo == TailNo);
        var overdue = new CAMM2OverdueDto();

        overdue.TailNo = TailNo;
        overdue.CurrentAFHR = CurrentAFHR;
        overdue.Venue = Venue;
        overdue.TailSerialNo = Tail_SerialNo;
        overdue.Part = PART;
        overdue.Cage = CAGE;
        overdue.MaintActivity = Maint_Activity;
        //if (!string.IsNullOrEmpty(Date_due))
        //{
        //    DateString = ConvertToDateSerial(Date_due);
        //    overdue.DATE DUE = DateString;
        //}
        overdue.LifeTypeInt = Life_Type_Int;
        //if (!string.IsNullOrEmpty(Int_Rem) && double.TryParse(Int_Rem, out double IntRemValue))
        //{
        //    overdue.REM = IntRemValue;
        //}
        overdue.IntRem = Int_Rem;
        overdue.DueDate = Date_due;
        overdue.TSSCode1 = TSSCode1;
        overdue.TSSCode2 = TSSCode2;
        overdue.TSSCode3 = TSSCode3;
        overdue.TSSCode4 = TSSCode4;
        overdue.LCN = LCN;
        overdue.ALC = ALC;
        overdue.PositionName = Position_Name;
        overdue.WAC = WAC;
        overdue.Type = Maint_Type;
        overdue.LVL = LVL;
        overdue.STD = STD;
        overdue.REC = REC;
        overdue.CLM = CLM;
        overdue.DatePlan = Date_Plan;
        overdue.Interval = INTERVAL;
        overdue.LifeSinceNew = Life_Since_New;
        overdue.LifeSinceNewType = Life_Since_New_Type;
        overdue.InitType = Init_Type;
        overdue.Instance = INSTANCE;
        overdue.LCN = Init_LCN;
        overdue.ALC = Init_ALC;
        overdue.RunDate = RunDate;
        asset.Overdue.Add(overdue);
    }
} 

public static class StringExtensions
{
    public static string SafeSubstring(this string text, int start, int length)
    {
        return text.Length <= start ? ""
            : text.Length - start <= length ? text.Substring(start)
            : text.Substring(start, length);
    }
}