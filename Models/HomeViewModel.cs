using System;
using System.ComponentModel.DataAnnotations;

namespace AgentDesktop.Models
{
	public class HomeViewModel
	{
	}

	public class ServiceTicket
	{
		public int TID { get; set; }
		public string description { get; set; }
		public int cnt { get; set; }
		//public string LoginID { get; set; }
	}

	public class ServiceReason
	{
		public int rnid { get; set; }
		public string Description { get; set; }
		public int inbound { get; set; }
		public int outbound { get; set; }
		public string title { get; set; }
		public string RCode { get; set; }
	}

	public class ServiceChannel
	{
		public int cid { get; set; }
		public string Description { get; set; }
		public int inbound { get; set; }
		public int outbound { get; set; }
		public string icon { get; set; }
	}


	public class Escalation
	{
		public string description { get; set; }
		public int cnt { get; set; }
	}

	public class DisplayServiceTicket
	{
		[Key]
		public int RID { get; set; }
		public string TicketNo { get; set; }
		public string CustomerName { get; set; }	
		public DateTime DateReceived { get; set; }
		public string Status { get; set; }
		public DateTime Resolved { get; set; }
		public int elapsed { get; set; }
		public string loginid { get; set; }
		public string agent { get; set; }
	}

	public class DisplayServiceTicketDated
	{
		[Key]
		public int RID { get; set; }
		public string TicketNo { get; set; }
		public string CustomerName { get; set; }
		public DateTime DateReceived { get; set; }
		public string Status { get; set; }
		public DateTime resolved { get; set; }
		public int elapsed { get; set; }
		public string agent { get; set; }
	}

	public class search
	{
		public int w_date { get; set; }
		public DateTime start_date { get; set; }
		public DateTime end_date { get; set; }
		public string customer { get; set; }
		public string card_no { get; set; }
		public string comm_type { get; set; }
		public string comm_channel { get; set; }
		public string reason_code { get; set; }
		public string status { get; set; }
		public string agent { get; set; }
		public string ticket_no { get; set; }
		public string trntype { get; set; }
		public string atmtrans { get; set; }
		public string channel { get; set; }
		public string report { get; set; }
		public string fname { get; set; }
		public string lname { get; set; }
		public string fullname { get; set; }
		public int chk { get; set; }
		public DateTime? stdate { get; set; }
		public DateTime? eddate { get; set; }
		public string client_no { get; set; }
		public string referedby { get; set; }
		public string contact { get; set; }
		public string tagging { get; set; }
		public string abbre { get; set; }
	}

	public class SqlQuery
	{
		public string Query { get; set; }
	}

	
}
