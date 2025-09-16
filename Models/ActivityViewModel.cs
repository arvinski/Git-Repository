using System;
using System.ComponentModel.DataAnnotations;

namespace AgentDesktop.Models
{

    public class ActivityViewModel
    {
    }

    public class CMActivity
    {
        public Int64 RID { get; set; }
        public string TicketNo { get; set; }
        public string InOutBound { get; set; }
        public string CommChannel { get; set; }
        public DateTime DateReceived { get; set; }
        public string abbreviation { get; set; }
        public string CustomerName { get; set; }
        public string CardNo { get; set; }
        public string Destination { get; set; }
        public string TransType { get; set; }
        public string ATMTrans { get; set; }
        public string Location { get; set; }
        public string ATMUsed { get; set; }
        public string PaymentOf { get; set; }
        public string TerminalUsed { get; set; }
        public string BillerName { get; set; }
        public string Merchant { get; set; }
        public string Inter { get; set; }
        public string BancnetUsed { get; set; }
        public string Online { get; set; }
        public string Website { get; set; }
        public string RemitFrom { get; set; }
        public string RemitConcern { get; set; }
        public double Amount { get; set; }
        public DateTime? DateTrans { get; set; }
        public string PLStatus { get; set; }
        public string CardBlocked { get; set; }
        public string EndoredTo { get; set; }
        public string Activity { get; set; }
        public string Status { get; set; }
        public string Agent { get; set; }
        public string created_by { get; set; }
        public string Last_Update { get; set; }
        public string EndorsedFrom { get; set; }
        public DateTime? Resolved { get; set; }
        public string Remarks { get; set; }
        public string ReferedBy { get; set; }
        public string ContactPerson { get; set; }
        public string ReasonCode { get; set; }
        public string Tagging { get; set; }
        public DateTime? ResolvedDate { get; set; }
        public string CallTime { get; set; }
        public string LocalNum { get; set; }
        public string CardPresent { get; set; }
		public string SubStatus { get; set; }
        public string Currency { get; set; }
        public string Branch { get; set; }
    }

    public class ListActivity
    {
        [Key]
        public Int64 RID { get; set; }
        public string TicketNo { get; set; }
        public string InOutBound { get; set; }
        public string CommChannel { get; set; }
        public DateTime DateReceived { get; set; }
        public string abbreviation { get; set; }
        public string CustomerName { get; set; }
        public string CardNo { get; set; }
        public string Destination { get; set; }
        public string TransType { get; set; }
        public string ATMTrans { get; set; }
        public string Location { get; set; }
        public string ATMUsed { get; set; }
        public string PaymentOf { get; set; }
        public string TerminalUsed { get; set; }
        public string BillerName { get; set; }
        public string Merchant { get; set; }
        public string Inter { get; set; }
        public string BancnetUsed { get; set; }
        public string Online { get; set; }
        public string Website { get; set; }
        public string RemitFrom { get; set; }
        public string RemitConcern { get; set; }
        public double Amount { get; set; }
        public DateTime DateTrans { get; set; }
        public string PLStatus { get; set; }
        public string CardBlocked { get; set; }
        public string EndoredTo { get; set; }
        public string Activity { get; set; }
        public string Status { get; set; }
        public string Agent { get; set; }
        public string created_by { get; set; }
        public string Last_Update { get; set; }
        public string EndorsedFrom { get; set; }
        public DateTime? Resolved { get; set; }
        public string Remarks { get; set; }
        public string ReferedBy { get; set; }
        public string ContactPerson { get; set; }
        public string ReasonCode { get; set; }
        public string Tagging { get; set; }
        public DateTime? ResolvedDate { get; set; }
        public string CallTime { get; set; }
        public string LocalNum { get; set; }
        public string CardPresent { get; set; }
    }

#nullable enable
    public class CallReport
    {
        [Key]
        public Int64 RID { get; set; }
        public string? CommChannel { get; set; }
        public DateTime? DateReceived { get; set; }
        public string? ReasonCode { get; set; }
        public string? NameVerified { get; set; }
        public string? CustomerName { get; set; }
        public string? CardNo { get; set; }
        public string? Activity { get; set; }
        public string? Status { get; set; }
        public string? Agent { get; set; }
        public string? TicketNo { get; set; }
        public string? Type { get; set; }
        public string? LastName { get; set; }
        public string? FirstName { get; set; }
        public string? MiddleName { get; set; }
        public string? InOutBound { get; set; }
        public string? TransType { get; set; }
        public string? Location { get; set; }
        public string? ATMUsed { get; set; }
        public string? PaymentOf { get; set; }
        public string? TerminalUsed { get; set; }
        public string? Merchant { get; set; }
        public string? Inter { get; set; }
        public string? BancnetUsed { get; set; }
        public string? Online { get; set; }
        public string? Website { get; set; }
        public double Amount { get; set; }
        public DateTime? DateTrans { get; set; }
        public string? PLStatus { get; set; }
        public string? CardBlocked { get; set; }
        public string? EndoredTo { get; set; }
        public string? Destination { get; set; }
        public string? ATMTrans { get; set; }
        public string? RemitFrom { get; set; }
        public string? RemitConcern { get; set; }
        public string? EndorsedFrom { get; set; }
        public int? Aging { get; set; }
        public DateTime? Resolved { get; set; }
        public string? BillerName { get; set; }
        public string? Remarks { get; set; }
        public string? ReferedBy { get; set; }
        public string? ContactPerson { get; set; }
        public string? Tagging { get; set; }
        public string? CallTime { get; set; }
        public string? LocalNum { get; set; }
        public string? Created_By { get; set; }
        public string? Last_Update { get; set; }
        public string? client_email { get; set; }
        public string? CardPresent { get; set; }
        public string? SubStatus { get; set; }
    }

#nullable disable

    public class NewCallReport
    {

        [Display(Name = "Ticket No")]
        public string TicketNo { get; set; }

        [Display(Name = "Comm Channel")]
        public string CommChannel { get; set; }

        [Display(Name = "Reason Code")]
        public string ReasonCode { get; set; }

        [Display(Name = "Check if name is verified")]
        public string NameVerified { get; set; }

        [Display(Name = "Customer Name")]
        public string CustomerName { get; set; }

        [Display(Name = "Card Number")]
        public string CardNo { get; set; }

        [Display(Name = "Activity")]
        public string Activity { get; set; }

        [Display(Name = "Status")]
        public string Status { get; set; }

        [Display(Name = "Agent")]
        public string Agent { get; set; }

        [Display(Name = "Last Name")]
        public string LastName { get; set; }

        [Display(Name = "First Name")]
        public string FirstName { get; set; }

        [Display(Name = "Middle Name")]
        public string MiddleName { get; set; }

        [Display(Name = "Comm Type")]
        public string InOutBound { get; set; }

        [Display(Name = "Channel")]
        public string TransType { get; set; }

        [Display(Name = "Location")]
        public string Location { get; set; }

        [Display(Name = "Location")]
        public string Location2 { get; set; }

        [Display(Name = "ATM Used")]
        public string ATMUsed { get; set; }

        [Display(Name = "Payment Institution")]
        public string PaymentOf { get; set; }

        [Display(Name = "Terminal Used")]
        public string TerminalUsed { get; set; }

        [Display(Name = "Merchant / Website")]
        public string Merchant { get; set; }

        [Display(Name = "Inter / Intra")]
        public string Inter { get; set; }

        [Display(Name = "Bancnet ATM Used")]
        public string BancnetUsed { get; set; }

        [Display(Name = "Check if online")]
        public string Online { get; set; }

        [Display(Name = "Website")]
        public string Website { get; set; }

        [Display(Name = "Amount")]
        public string Amount { get; set; }

        [Display(Name = "Transaction Date")]
        public string DateTrans { get; set; }

        [Display(Name = "Loan Status")]
        public string PLStatus { get; set; }

        [Display(Name = "Check if Card was Blocked")]
        public string CardBlocked { get; set; }

        [Display(Name = "Endorsed To")]
        public string EndoredTo { get; set; }

        [Display(Name = "Destination")]
        public string Destination { get; set; }

        [Display(Name = "Transaction Type")]
        public string ATMTrans { get; set; }

        [Display(Name = "Remittance From")]
        public string RemitFrom { get; set; }

        [Display(Name = "Remittance Concern")]
        public string RemitConcern { get; set; }

        [Display(Name = "Endorsed From")]
        public string EndorsedFrom { get; set; }

        [Display(Name = "Due Date")]
        public string Resolved { get; set; }

        [Display(Name = "Biller Name")]
        public string BillerName { get; set; }

        [Display(Name = "Nature of Complaint")]
        public string Remarks { get; set; }

        [Display(Name = "Referred By")]
        public string ReferedBy { get; set; }

        [Display(Name = "Contact Person")]
        public string ContactPerson { get; set; }

        [Display(Name = "Classification")]
        public string Tagging { get; set; }

        [Display(Name = "Time of Call")]
        public string CallTime { get; set; }

        [Display(Name = "Local Number")]
        public string LocalNum { get; set; }

        [Display(Name = "Created By")]
        public string Created_By { get; set; }

        [Display(Name = "Last Updated By")]
        public string Last_Update { get; set; }

        [Display(Name = "Card Present / Not Present")]
        public string CardPresent { get; set; }


        [Display(Name = "Sub Status")]
        public string SubStatus { get; set; }

        [Display(Name = "Amount Currency")]
        public string Currency { get; set; }

        [Display(Name = "Branch")]
        public string Branch { get; set; }

        //		[Display(Name = "Search Client")]
        //#pragma warning disable CS8632 // The annotation for nullable reference types should only be used in code within a '#nullable' annotations context.
        //		public string? SearchClient { get; set; }
        //#pragma warning restore CS8632 // The annotation for nullable reference types should only be used in code within a '#nullable' annotations context.
    }


	public class Combo1
    {
		public string ID { get; set; }
		public string Description { get; set; }
	}

    public class Combo2
    {
        public int ID { get; set; }
        public string Description { get; set; }
    }

    public class Combo3
    {
        public string Value { get; set; }
        public string Text { get; set; }
    }

    public class SearchICBA
	{
		public string DESCR { get; set; }
		public string CRLINE { get; set; }
		public string LNAME { get; set; }
		public string FSNAME { get; set; }
		public string MDNAME { get; set; }
		public string LSNAME { get; set; }
		public int CIFKEY { get; set; }
		public string ACCTNO { get; set; }
		public string ACSTS { get; set; }
		public string CLIENT_NAME { get; set; }
	}

    public class SearchBW
    {
        public string CARD_NUMBER { get; set; }
        public string FULL_NAME { get; set; }
        public string LAST_NAME { get; set; }
        public string FIRST_NAME { get; set; }
        public string MID_NAME { get; set; }
        public string CARD_STATUS { get; set; }
        public string NOTE_TEXT { get; set; }
        public string CARD_EXPIRY_DATE { get; set; }
        public string ACCT_NUMBER { get; set; }
        public double BALANCE { get; set; }
        public string CLIENT_NAME { get; set; }
    }

    public class ATM_GROUP_LOC
	{
		public int ID { get; set; }
		public string ATM { get; set; }
		public string ATM_LOCATION { get; set; }
	}

    public class ATMLogs
	{
        public int crid { get; set; }
        public int rid { get; set; }
        public string CommChannel { get; set; }
        public string DateReceived { get; set; }
        public string ReasonCode { get; set; }
        public string NameVerified { get; set; }
        public string CIFKey { get; set; }
        public string CustomerName { get; set; }
        public string CardNo { get; set; }
        public string Activity { get; set; }
        public string Status { get; set; }
        public string Agent { get; set; }
        public string TicketNo { get; set; }
        public string Type { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string InOutBound { get; set; }
        public string TransType { get; set; }
        public string Location { get; set; }
        public string ATMUsed { get; set; }
        public string PaymentOf { get; set; }
        public string TerminalUsed { get; set; }
        public string Merchant { get; set; }
        public string Inter { get; set; }
        public string BancnetUsed { get; set; }
        public string Online { get; set; }
        public string Website { get; set; }
        public string Amount { get; set; }
        public string DateTrans { get; set; }
        public string PLStatus { get; set; }
        public string CardBlocked { get; set; }
        public string EndoredTo { get; set; }
        public string ATMAgent { get; set; }
        public string Destination { get; set; }
        public string ATMTrans { get; set; }
        public string RemitFrom { get; set; }
        public string RemitConcern { get; set; }
        public string Resolved { get; set; }
        public string BillerName { get; set; }
        public string Remarks { get; set; }
        public string ReferedBy { get; set; }
        public string ContactPerson { get; set; }
        public string Tagging { get; set; }

    }

    public class ATMOps
	{
        public int RID { get; set; }
        public string TicketNo { get; set; }
        public string InOutBound { get; set; }
        public string CommChannel { get; set; }
        public DateTime DateReceived { get; set; }
        public string CustomerName { get; set; }
        public string CardNo { get; set; }
        public string Destination { get; set; }
        public string TransType { get; set; }
        public string ATMTrans { get; set; }
        public string Location { get; set; }
        public string ATMUsed { get; set; }
        public string PaymentOf { get; set; }
        public string TerminalUsed { get; set; }
        public string BillerName { get; set; }
        public string Merchant { get; set; }
        public string Inter { get; set; }
        public string BancnetUsed { get; set; }
        public string Online { get; set; }
        public string Website { get; set; }
        public string RemitFrom { get; set; }
        public string RemitConcern { get; set; }
        public double Amount { get; set; }
        public string DateTrans { get; set; }
        public string PLStatus { get; set; }
        public string CardBlocked { get; set; }
        public string EndoredTo { get; set; }
        public string Activity { get; set; }
        public string Status { get; set; }
        public string Agent { get; set; }
        public string Resolved { get; set; }
        public string Remarks { get; set; }
        public string ReferedBy { get; set; }
        public string ContactPerson { get; set; }
        public string abbreviation { get; set; }
        public string Tagging { get; set; }
        public string ResolvedDate { get; set; }
        public string CardPresent { get; set; }
        public string SubStatus { get; set; }
		public string Currency { get; set; }
        public string Branch { get; set; }
        public string Created_By { get; set; }
    }

    public class SearchCardNo
    {
        public string TicketNo { get; set; }
        public string DateReceived { get; set; }
        public string CustomerName { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string CardNo { get; set; }
        public string Status { get; set; }
		public string CLIENT_NAME { get; set; }
	}

    public class Show_Ticket
    {
        public string ticket { get; set; }
    }

    public class EditCallReport
    {
        [Key]
		public int RID { get; set; }

		[Display(Name = "Ticket No")]
        public string TicketNo { get; set; }

        [Display(Name = "Comm Channel")]
        public string CommChannel { get; set; }

        [Display(Name = "Reason Code")]
        public string ReasonCode { get; set; }

        [Display(Name = "Customer Name")]
        public string CustomerName { get; set; }

        [Display(Name = "Card Number")]
        public string CardNo { get; set; }

        [Display(Name = "Activity")]
        public string Activity { get; set; }

        [Display(Name = "Status")]
        public string Status { get; set; }

        [Display(Name = "Agent")]
        public string Agent { get; set; }

        [Display(Name = "Last Name")]
        public string LastName { get; set; }

        [Display(Name = "First Name")]
        public string FirstName { get; set; }

        [Display(Name = "Middle Name")]
        public string MiddleName { get; set; }

        [Display(Name = "Comm Type")]
        public string InOutBound { get; set; }

        [Display(Name = "Channel")]
        public string TransType { get; set; }

        [Display(Name = "Location")]
        public string Location { get; set; }

        [Display(Name = "Location")]
        public string Location2 { get; set; }

        [Display(Name = "ATM Used")]
        public string ATMUsed { get; set; }

        [Display(Name = "Payment Institution")]
        public string PaymentOf { get; set; }

        [Display(Name = "Terminal Used")]
        public string TerminalUsed { get; set; }

        [Display(Name = "Merchant / Website")]
        public string Merchant { get; set; }

        [Display(Name = "Inter / Intra")]
        public string Inter { get; set; }

        [Display(Name = "Bancnet ATM Used")]
        public string BancnetUsed { get; set; }

        [Display(Name = "Check if online")]
        public bool Online { get; set; }

        [Display(Name = "Website")]
        public string Website { get; set; }

        [Display(Name = "Amount")]
        public string Amount { get; set; }

        [Display(Name = "Transaction Date")]
        public string DateTrans { get; set; }

        [Display(Name = "Loan Status")]
        public string PLStatus { get; set; }

        [Display(Name = "Check if Card was Blocked")]
        public bool CardBlocked { get; set; }

        [Display(Name = "Endorsed To")]
        public string EndoredTo { get; set; }

        [Display(Name = "Destination")]
        public string Destination { get; set; }

        [Display(Name = "Transaction Type")]
        public string ATMTrans { get; set; }

        [Display(Name = "Remittance From")]
        public string RemitFrom { get; set; }

        [Display(Name = "Remittance Concern")]
        public string RemitConcern { get; set; }

        [Display(Name = "Endorsed From")]
        public string EndorsedFrom { get; set; }

        [Display(Name = "Due Date")]
        public string Resolved { get; set; }

        [Display(Name = "Biller Name")]
        public string BillerName { get; set; }

        [Display(Name = "Nature of Complaint")]
        public string Remarks { get; set; }

        [Display(Name = "Referred By")]
        public string ReferedBy { get; set; }

        [Display(Name = "Contact Person")]
        public string ContactPerson { get; set; }

        [Display(Name = "Classification")]
        public string Tagging { get; set; }

        [Display(Name = "Time of Call")]
        public string CallTime { get; set; }

        [Display(Name = "Local Number")]
        public string LocalNum { get; set; }

        [Display(Name = "Created By")]
        public string Created_By { get; set; }

        [Display(Name = "Last Updated By")]
        public string Last_Update { get; set; }

        [Display(Name = "Card Present / Not Present")]
        public string CardPresent { get; set; }


        [Display(Name = "Sub Status")]
        public string SubStatus { get; set; }

        [Display(Name = "Currency")]
        public string Currency { get; set; }

        [Display(Name = "Branch")]
        public string Branch { get; set; }

    }

    public class IRemitCode
    {
        public string RCode { get; set; }
    }

    public class CType
	{
		public int RID { get; set; }
		public string Type { get; set; }
	}

#nullable enable
    public class ATMReport
    {
        [Key]
        public Int64 RID { get; set; }
        public string? CommChannel { get; set; }
        public DateTime? DateReceived { get; set; }
        public string? ReasonCode { get; set; }
        public string? NameVerified { get; set; }
        public string? CustomerName { get; set; }
        public string? CardNo { get; set; }
        public string? Activity { get; set; }
        public string? Status { get; set; }
        public string? Agent { get; set; }
        public string? TicketNo { get; set; }
        public string? Type { get; set; }
        public string? LastName { get; set; }
        public string? FirstName { get; set; }
        public string? MiddleName { get; set; }
        public string? InOutBound { get; set; }
        public string? TransType { get; set; }
        public string? Location { get; set; }
        public string? ATMUsed { get; set; }
        public string? PaymentOf { get; set; }
        public string? TerminalUsed { get; set; }
        public string? Merchant { get; set; }
        public string? Inter { get; set; }
        public string? BancnetUsed { get; set; }
        public string? Online { get; set; }
        public string? Website { get; set; }
        public double Amount { get; set; }
        public DateTime? DateTrans { get; set; }
        public string? PLStatus { get; set; }
        public string? CardBlocked { get; set; }
        public string? EndoredTo { get; set; }
        public string? ATMAgent { get; set; }
        public string? Destination { get; set; }
        public string? ATMTrans { get; set; }
        public string? RemitFrom { get; set; }
        public string? RemitConcern { get; set; }
        //public string? EndorsedFrom { get; set; }
        //public int? Aging { get; set; }
        //public DateTime? Resolved { get; set; }
        public string? BillerName { get; set; }
        public string? Remarks { get; set; }
        public string? ReferedBy { get; set; }
        public string? ContactPerson { get; set; }
        public string? Tagging { get; set; }
        //public string? CallTime { get; set; }
        //public string? LocalNum { get; set; }
        //public string? Created_By { get; set; }
        //public string? Last_Update { get; set; }
        //public string? client_email { get; set; }
        public string? CardPresent { get; set; }
        public string? SubStatus { get; set; }
    }

#nullable disable

    public class ATMActivity
    {
        public Int64 RID { get; set; }
        public string TicketNo { get; set; }
        public string InOutBound { get; set; }
        public string CommChannel { get; set; }
        public DateTime DateReceived { get; set; }
        public string abbreviation { get; set; }
        public string CustomerName { get; set; }
        public string CardNo { get; set; }
        public string Destination { get; set; }
        public string TransType { get; set; }
        public string ATMTrans { get; set; }
        public string Location { get; set; }
        public string ATMUsed { get; set; }
        public string PaymentOf { get; set; }
        public string TerminalUsed { get; set; }
        public string BillerName { get; set; }
        public string Merchant { get; set; }
        public string Inter { get; set; }
        public string BancnetUsed { get; set; }
        public string Online { get; set; }
        public string Website { get; set; }
        public string RemitFrom { get; set; }
        public string RemitConcern { get; set; }
        public double Amount { get; set; }
        public DateTime? DateTrans { get; set; }
        public string PLStatus { get; set; }
        public string CardBlocked { get; set; }
        public string EndoredTo { get; set; }
        public string Activity { get; set; }
        public string Status { get; set; }
        public string Agent { get; set; }
		public string Created_By { get; set; }
		public string Last_Update { get; set; }
		//public string EndorsedFrom { get; set; }
		public DateTime? Resolved { get; set; }
        public string Remarks { get; set; }
        public string ReferedBy { get; set; }
        public string ContactPerson { get; set; }
        public string ReasonCode { get; set; }
        public string Tagging { get; set; }
        public DateTime? ResolvedDate { get; set; }
        //public string CallTime { get; set; }
        //public string LocalNum { get; set; }
        public string CardPresent { get; set; }
        public string SubStatus { get; set; }

    }

    public class ATMUsed
	{
		public int ATM_ID { get; set; }
	}

    public class NewATMReport
    {
        [Key]
		public int RID { get; set; }

		[Display(Name = "Ticket No")]
        public string TicketNo { get; set; }

        [Required]
        [Display(Name = "Comm Channel")]
        public string CommChannel { get; set; }

        [Required]
        [Display(Name = "Reason Code")]
        public string ReasonCode { get; set; }

        [Display(Name = "Check if name is verified")]
        public string NameVerified { get; set; }

        [Display(Name = "Customer Name")]
        public string CustomerName { get; set; }

        [Display(Name = "Card Number")]
        public string CardNo { get; set; }

        [Required]
        [Display(Name = "Activity")]
        public string Activity { get; set; }

        [Required]
        [Display(Name = "Status")]
        public string Status { get; set; }

        [Display(Name = "Agent")]
        public string Agent { get; set; }

        [Display(Name = "Last Name")]
        public string LastName { get; set; }

        [Display(Name = "First Name")]
        public string FirstName { get; set; }

        [Display(Name = "Middle Name")]
        public string MiddleName { get; set; }

        [Required]
        [Display(Name = "Comm Type")]
        public string InOutBound { get; set; }

        [Display(Name = "Channel")]
        public string TransType { get; set; }

        [Display(Name = "Location")]
        public string Location { get; set; }

        [Display(Name = "ATM Used")]
        public string ATMUsed { get; set; }

        [Display(Name = "Payment Institution")]
        public string PaymentOf { get; set; }

        [Display(Name = "Terminal Used")]
        public string TerminalUsed { get; set; }

        [Display(Name = "Merchant / Website")]
        public string Merchant { get; set; }

        [Display(Name = "Inter / Intra")]
        public string Inter { get; set; }

        [Display(Name = "Bancnet ATM Used")]
        public string BancnetUsed { get; set; }

        [Display(Name = "Check if online")]
        public string Online { get; set; }

        [Display(Name = "Website")]
        public string Website { get; set; }

        [Display(Name = "Amount")]
        public string Amount { get; set; }

        [Display(Name = "Transaction Date")]
        public string DateTrans { get; set; }

        [Display(Name = "Loan Status")]
        public string PLStatus { get; set; }

        [Display(Name = "Check if Card was Blocked")]
        public string CardBlocked { get; set; }

        [Display(Name = "Endorsed To")]
        public string EndoredTo { get; set; }

        [Display(Name = "Destination")]
        public string Destination { get; set; }

        [Display(Name = "Transaction Type")]
        public string ATMTrans { get; set; }

        [Display(Name = "Remittance From")]
        public string RemitFrom { get; set; }

        [Display(Name = "Remittance Concern")]
        public string RemitConcern { get; set; }

        //[Display(Name = "Endorsed From")]
        //public string EndorsedFrom { get; set; }

        //[Display(Name = "Due Date")]
        //public string Resolved { get; set; }

        [Display(Name = "Biller Name")]
        public string BillerName { get; set; }

        [Display(Name = "Nature of Complaint")]
        public string Remarks { get; set; }

        [Display(Name = "Referred By")]
        public string ReferedBy { get; set; }

        [Display(Name = "Contact Person")]
        public string ContactPerson { get; set; }

        [Display(Name = "Classification")]
        public string Tagging { get; set; }

        //[Display(Name = "Time of Call")]
        //public string CallTime { get; set; }

        //[Display(Name = "Local Number")]
        //public string LocalNum { get; set; }

        //[Display(Name = "Created By")]
        //public string Created_By { get; set; }

        //[Display(Name = "Last Updated By")]
        //public string Last_Update { get; set; }

        [Display(Name = "Card Present / Not Present")]
        public string CardPresent { get; set; }

        [Display(Name = "Sub Status")]
        public string SubStatus { get; set; }

    }

    public class EditATMReport
    {
        [Key]
        public int RID { get; set; }

        [Display(Name = "Ticket No")]
        public string TicketNo { get; set; }

        [Display(Name = "Comm Channel")]
        public string CommChannel { get; set; }

        [Display(Name = "Reason Code")]
        public string ReasonCode { get; set; }

        [Display(Name = "Customer Name")]
        public string CustomerName { get; set; }

        [Display(Name = "Card Number")]
        public string CardNo { get; set; }

        [Display(Name = "Activity")]
        public string Activity { get; set; }

        [Display(Name = "Status")]
        public string Status { get; set; }

        [Display(Name = "Agent")]
        public string Agent { get; set; }

        [Display(Name = "Last Name")]
        public string LastName { get; set; }

        [Display(Name = "First Name")]
        public string FirstName { get; set; }

        [Display(Name = "Middle Name")]
        public string MiddleName { get; set; }

        [Display(Name = "Comm Type")]
        public string InOutBound { get; set; }

        [Display(Name = "Channel")]
        public string TransType { get; set; }

        [Display(Name = "Location")]
        public string Location { get; set; }

        [Display(Name = "ATM Used")]
        public string ATMUsed { get; set; }

        [Display(Name = "Payment Institution")]
        public string PaymentOf { get; set; }

        [Display(Name = "Terminal Used")]
        public string TerminalUsed { get; set; }

        [Display(Name = "Merchant / Website")]
        public string Merchant { get; set; }

        [Display(Name = "Inter / Intra")]
        public string Inter { get; set; }

        [Display(Name = "Bancnet ATM Used")]
        public string BancnetUsed { get; set; }

        [Display(Name = "Check if online")]
        public bool Online { get; set; }

        [Display(Name = "Website")]
        public string Website { get; set; }

        [Display(Name = "Amount")]
        public string Amount { get; set; }

        [Display(Name = "Transaction Date")]
        public string DateTrans { get; set; }

        [Display(Name = "Loan Status")]
        public string PLStatus { get; set; }

        [Display(Name = "Check if Card was Blocked")]
        public bool CardBlocked { get; set; }

        [Display(Name = "Endorsed To")]
        public string EndoredTo { get; set; }

        [Display(Name = "Destination")]
        public string Destination { get; set; }

        [Display(Name = "Transaction Type")]
        public string ATMTrans { get; set; }

        [Display(Name = "Remittance From")]
        public string RemitFrom { get; set; }

        [Display(Name = "Remittance Concern")]
        public string RemitConcern { get; set; }

		//[Display(Name = "Endorsed From")]
		//public string EndorsedFrom { get; set; }

		[Display(Name = "Due Date")]
		public string Resolved { get; set; }

		[Display(Name = "Biller Name")]
        public string BillerName { get; set; }

        [Display(Name = "Nature of Complaint")]
        public string Remarks { get; set; }

        [Display(Name = "Referred By")]
        public string ReferedBy { get; set; }

        [Display(Name = "Contact Person")]
        public string ContactPerson { get; set; }

        [Display(Name = "Classification")]
        public string Tagging { get; set; }

		//[Display(Name = "Time of Call")]
		//public string CallTime { get; set; }

		//[Display(Name = "Local Number")]
		//public string LocalNum { get; set; }

		[Display(Name = "Created By")]
		public string Created_By { get; set; }

		//[Display(Name = "Last Updated By")]
		//public string Last_Update { get; set; }

		[Display(Name = "Card Present / Not Present")]
        public string CardPresent { get; set; }


        [Display(Name = "Sub Status")]
        public string SubStatus { get; set; }

        [Display(Name = "Currency")]
        public string Currency { get; set; }

        [Display(Name = "Branch")]
        public string Branch { get; set; }

        [Display(Name = "Check if name is verified")]
        public string NameVerified { get; set; }

    }

    public class ActivitySearch
    {
		public bool DateSearch { get; set; }
		public string Date1 { get; set; }
        public string Date2 { get; set; }
        public string CustomerName { get; set; }
        public string CardNo { get; set; }
        public string CommType { get; set; }
        public string CommChannel { get; set; }
        public string ReasonCode { get; set; }
        public string Status { get; set; }
        public string Agent { get; set; }
        public string TicketNo { get; set; }
        public string ReferredBy { get; set; }
        public string ContactPerson { get; set; }

    }

    public class RemoveFileModel
    {
        public string File { get; set; }
    }

    public class MaxRID
	{
        public int RID { get; set; }
	}
}
