using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CAMM2
{
    public class CAMM2AssetDto
    {
        public string Date { get; set; }
        public string TailNo { get; set; }
        public string Venue { get; set; }
        public List<CAMM2AssetHeaderInfoDto> HeaderInfo { get; set; } = new List<CAMM2AssetHeaderInfoDto>();
        public List<CAMM2ActivityDto> Activities { get; set; } = new List<CAMM2ActivityDto>();
        public List<CAMM2DeviationDto> Deviations { get; set; } = new List<CAMM2DeviationDto>();
        public List<CAMM2CoincidentDto> Coincidents { get; set; } = new List<CAMM2CoincidentDto>();
        public List<CAMM2InProgressDto> InProgress { get; set; } = new List<CAMM2InProgressDto>();
        public List<CAMM2OverdueDto> Overdue { get; set; } = new List<CAMM2OverdueDto>(); 
        public List<CAMM2SMRDto> SMR { get; set; } = new List<CAMM2SMRDto>();
        public List<CAMM2CFUDto> CFU { get; set; } = new List<CAMM2CFUDto>();
    }
    public class CAMM2AssetHeaderInfoDto
    {
        public decimal Life { get; set; }
        public string LifeType { get; set; }
        public decimal Interval { get; set; }
    }
    public class CAMM2ActivityDto
    {
        public string TailNo { get; set; }
        public string Venue { get; set; }
        public decimal CurrentAFHR { get; set; }
        public string TailSerialNo { get; set; }
        public string Part { get; set; }
        public string Cage { get; set; }
        public string MaintActivity { get; set; }
        public string DueDate { get; set; }
        public string LifeTypeInt { get; set; }
        public string TSSCode1 { get; set; }
        public string TSSCode2 { get; set; }
        public string TSSCode3 { get; set; }
        public string TSSCode4 { get; set; }
        public string LCN { get; set; }
        public string ALC { get; set; }
        public string PositionName { get; set; }
        public string WAC { get; set; }
        public string Type { get; set; }
        public string LVL { get; set; }
        public string STD { get; set; }
        public string REM { get; set; }
        public string REC { get; set; }
        public string CLM { get; set; }
        public string DatePlan { get; set; }
        public string Interval { get; set; }
        public string LifeSinceNew { get; set; }
        public string LifeSinceNewType { get; set; }
        public string InitType { get; set; }
        public string Instance { get; set; }
        public string InitLCN { get; set; }
        public string InitALC { get; set; }
    }
    public class CAMM2DeviationDto
    {
        public string TailNo { get; set; }
        public string Venue { get; set; }
        public decimal CurrentAFHR { get; set; }
        public string TailSerialNo { get; set; }
        public string Part { get; set; }
        public string Cage { get; set; }
        public string DeviationOrder { get; set; }
        public string STD { get; set; }
        public string Expired { get; set; }
        public string LCN { get; set; }
        public string ALC { get; set; }
        public string PositionName { get; set; }
        public string WAC { get; set; }
        public string Title { get; set; }
        public string LifeSinceNew { get; set; }
        public string LifeSinceNewType { get; set; }
        public string ExpiryConditions { get; set; }
        public string Limitaion { get; set; }
        public string Type { get; set; }
        public string Activity { get; set; }
        public string InitLCN { get; set; }
        public string InitALC { get; set; }
        public string Interval { get; set; }
        public string IntRem { get; set; }
        public string DueDate { get; set; }
        public string RunDate { get; set; }
    }
    public class CAMM2CoincidentDto
    {
        public string TailNo { get; set; }
        public string Venue { get; set; }
        public decimal CurrentAFHR { get; set; }
        public string TailSerialNo { get; set; }
        public string Part { get; set; }
        public string Cage { get; set; }
        public string MaintActivity { get; set; }
        public string DueDate { get; set; }
        public string IntRem { get; set; }
        public string LifeTypeInt { get; set; }
        public string TSSCode1 { get; set; }
        public string TSSCode2 { get; set; }
        public string TSSCode3 { get; set; }
        public string TSSCode4 { get; set; }
        public string LCN { get; set; }
        public string ALC { get; set; }
        public string PositionName { get; set; }
        public string WAC { get; set; }
        public string Type { get; set; }
        public string LVL { get; set; }
        public string STD { get; set; }
        public string REM { get; set; }
        public string REC { get; set; }
        public string CLM { get; set; }
        public string DatePlan { get; set; }
        public string Interval { get; set; }
        public string LifeSinceNew { get; set; }
        public string LifeSinceNewType { get; set; }
        public string InitType { get; set; }
        public string Instance { get; set; }
        public string InitLCN { get; set; }
        public string InitALC { get; set; }
        public string RunDate { get; set; }
    }

    public class CAMM2InProgressDto
    {
        public string TailNo { get; set; }
        public string Venue { get; set; }
        public decimal CurrentAFHR { get; set; }
        public string TailSerialNo { get; set; }
        public string Part { get; set; }
        public string Cage { get; set; }
        public string MaintActivity { get; set; }
        public string DueDate { get; set; }
        public string IntRem { get; set; }
        public string LifeTypeInt { get; set; }
        public string TSSCode1 { get; set; }
        public string TSSCode2 { get; set; }
        public string TSSCode3 { get; set; }
        public string TSSCode4 { get; set; }
        public string LCN { get; set; }
        public string ALC { get; set; }
        public string PositionName { get; set; }
        public string WAC { get; set; }
        public string Type { get; set; }
        public string LVL { get; set; }
        public string STD { get; set; }
        public string REM { get; set; }
        public string REC { get; set; }
        public string CLM { get; set; }
        public string DatePlan { get; set; }
        public string Interval { get; set; }
        public string LifeSinceNew { get; set; }
        public string LifeSinceNewType { get; set; }
        public string InitType { get; set; }
        public string Instance { get; set; }
        public string InitLCN { get; set; }
        public string InitALC { get; set; }
        public string RunDate { get; set; }
    }
    public class CAMM2OverdueDto
    {
        public string TailNo { get; set; }
        public string Venue { get; set; }
        public decimal CurrentAFHR { get; set; }
        public string TailSerialNo { get; set; }
        public string Part { get; set; }
        public string Cage { get; set; }
        public string MaintActivity { get; set; }
        public string DueDate { get; set; }
        public string IntRem { get; set; }
        public string LifeTypeInt { get; set; }
        public string TSSCode1 { get; set; }
        public string TSSCode2 { get; set; }
        public string TSSCode3 { get; set; }
        public string TSSCode4 { get; set; }
        public string LCN { get; set; }
        public string ALC { get; set; }
        public string PositionName { get; set; }
        public string WAC { get; set; }
        public string Type { get; set; }
        public string LVL { get; set; }
        public string STD { get; set; }
        public string REM { get; set; }
        public string REC { get; set; }
        public string CLM { get; set; }
        public string DatePlan { get; set; }
        public string Interval { get; set; }
        public string LifeSinceNew { get; set; }
        public string LifeSinceNewType { get; set; }
        public string InitType { get; set; }
        public string Instance { get; set; }
        public string InitLCN { get; set; }
        public string InitALC { get; set; }
        public string RunDate { get; set; }
    }
    public class CAMM2SMRDto
    {
        public string TailNo { get; set; }
        public string Venue { get; set; }
        public decimal CurrentAFHR { get; set; }
        public string TailSerialNo { get; set; }
        public string Part { get; set; }
        public string Cage { get; set; }
        public string MaintActivity { get; set; }
        public string LifeTypeInt { get; set; }
        public string DueDate { get; set; }
        public string IntRem { get; set; }
        public string LCN { get; set; }
        public string ALC { get; set; }
        public string PositionName { get; set; }
        public string WAC { get; set; }
        public string Type { get; set; }
        public string LVL { get; set; }
        public string STD { get; set; }
        public string REM { get; set; }
        public string REC { get; set; }
        public string CLM { get; set; }
        public string DatePlan { get; set; }
        public string Interval { get; set; }
        public string LifeSinceNew { get; set; }
        public string LifeSinceNewType { get; set; }
        public string InitType { get; set; }
        public string Instance { get; set; }
        public string InitLCN { get; set; }
        public string InitALC { get; set; }
        public string RunDate { get; set; }
        public string Symptoms { get; set; }
    }
    public class CAMM2CFUDto
    {
        public string TailNo { get; set; }
        public string Venue { get; set; }
        public decimal CurrentAFHR { get; set; }
        public string TailSerialNo { get; set; }
        public string Part { get; set; }
        public string Cage { get; set; }
        public string MaintActivity { get; set; }
        public string LifeTypeInt { get; set; }
        public string DueDate { get; set; }
        public string IntRem { get; set; }
        public string LCN { get; set; }
        public string ALC { get; set; }
        public string PositionName { get; set; }
        public string WAC { get; set; }
        public string Type { get; set; }
        public string LVL { get; set; }
        public string STD { get; set; }
        public string REM { get; set; }
        public string REC { get; set; }
        public string CLM { get; set; }
        public string DatePlan { get; set; }
        public string Interval { get; set; }
        public string LifeSinceNew { get; set; }
        public string LifeSinceNewType { get; set; }
        public string InitType { get; set; }
        public string Instance { get; set; }
        public string InitLCN { get; set; }
        public string InitALC { get; set; }
        public string RunDate { get; set; }
        public string Symptoms { get; set; }
    }
}