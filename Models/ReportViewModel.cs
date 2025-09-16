using System;
using System.ComponentModel.DataAnnotations;

namespace AgentDesktop.Models
{
	public class ReportViewModel
	{
	}

	public class Summary
	{
		public string descr { get; set; }
		public int qty { get; set; }
		public string suma { get; set; }
		public double per { get; set; }
	}

	public class ReportSummary
	{
		[DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMM-yyyy}")]
		public DateTime Date1 { get; set; }

		[DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMM-yyyy}")]
		public DateTime Date2 { get; set; }
		public string ReportType { get; set; }
		public string CommType { get; set; }
		public string Agent { get; set; }
	}

	public class ReportManagement
	{
		[DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMM-yyyy}")]
		public DateTime Date1 { get; set; }

		[DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMM-yyyy}")]
		public DateTime Date2 { get; set; }
		public string ReportType { get; set; }
		public string CommType { get; set; }
		public string Agent { get; set; }
		public string Channel { get; set; }
	}

	public class ReportActivity
	{
		public string tyme { get; set; }
		public int Monday { get; set; }
		public int Tuesday { get; set; }
		public int Wednesday { get; set; }
		public int Thursday { get; set; }
		public int Friday { get; set; }
		public int Saturday { get; set; }
		public int Sunday { get; set; }
	}

	public class ReasonComm
	{
		public string ID { get; set; }
		public string ReasonC { get; set; }
		public string reasoncode { get; set; }
		public string ComChannel { get; set; }
		public int Inbound { get; set; }
		public string inper { get; set; }
		public int outbound { get; set; }
		public string outper { get; set; }
		public int total { get; set; }
		public string totalper { get; set; }
	}

	public class ReasonStat
	{
		public string ID { get; set; }
		public string ReasonC { get; set; }
		public string reasoncode { get; set; }
		public string Status { get; set; }
		public int Inbound { get; set; }
		public string inper { get; set; }
		public int outbound { get; set; }
		public string outper { get; set; }
		public int total { get; set; }
		public string totalper { get; set; }
	}

	public class ReportAging
	{
		[DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMM-yyyy}")]
		public DateTime Date1 { get; set; }

		[DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMM-yyyy}")]
		public DateTime Date2 { get; set; }
		public string ReportType { get; set; }
		public string ReasonCode { get; set; }
		public string Branch { get; set; }
	}

	public class Aging
	{
		public string TicketNo { get; set; }
		public string CustomerName { get; set; }
		public string Activity { get; set; }
		public DateTime DateReceived { get; set; }
		public int Running { get; set; }
		public int Follow { get; set; }
		public string Agent { get; set; }
	}

	public class AgingDetail
	{
		public string DateReceived { get; set; }
		public string TimeReceived { get; set; }
		public string TicketNo { get; set; }
		public string DueDate { get; set; }
		public string CustomerName { get; set; }
		public string CardNo { get; set; }
		public string BranchAccount { get; set; }
		public string TypeOfComplaint { get; set; }
		public string Merchant { get; set; }
		public string ATMUsed { get; set; }
		public string Location { get; set; }
		public string Amount { get; set; }
		public string Currency { get; set; }
		public string DateTrans { get; set; }
		public string Activity { get; set; }
		public string UpdateFromEndorsed { get; set; }
		public string Status { get; set; }
		public string EndoredTo { get; set; }
		public string ResolvedDate { get; set; }
		public string last_update { get; set; }
		public string ReferedBy { get; set; }
		public string created_by { get; set; }
		public string Tagging { get; set; }
		public string aging { get; set; }
		public string Within_Outside_TAT { get; set; }
	}

	public class ReportIRC
	{
		public string reason { get; set; }
		public string reasoncode { get; set; }
		public int cnt { get; set; }
		public int complex { get; set; }
		public int simple { get; set; }
		public int resolved { get; set; }
		public int open	{ get; set; }
		public int cancelled { get; set; }
		public int closed { get; set; }
		public int commitment { get; set; }
		public int endorsed { get; set; }
		public int resolved2 { get; set; }

	}

	public class ReportIRCFilter
	{
		[DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMM-yyyy}")]
		public DateTime Date1 { get; set; }

		[DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMM-yyyy}")]
		public DateTime Date2 { get; set; }

		[Required]
		public string ReportType { get; set; }

		[Required]
		public string ReportTypeDet { get; set; }


	}

}
