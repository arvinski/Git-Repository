using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace AgentDesktop.Models
{
	public class InquiryViewModel
	{
	}

	public class Inquiry
	{
#pragma warning disable CS8632 // The annotation for nullable reference types should only be used in code within a '#nullable' annotations context.
		public string? CardNo { get; set; }
		public string? FirstName { get; set; }
		public string? LastName { get; set; }
		public string? FullName { get; set; }
#pragma warning restore CS8632 // The annotation for nullable reference types should only be used in code within a '#nullable' annotations context.

	}

	public class InquiryICBA
	{
		public string acctno { get; set; }
		public string lname { get; set; }
		public string descr { get; set; }
		public string crline { get; set; }
		public string acsts { get; set; }
		public string cifkey { get; set; }

	}

	public class MasterDetailFile
	{
		public List<displayCustomer> vwCustomerInfo { get; set; }
		public List<displayAccountInfo> vwAccountInfo { get; set; }
		public List<displayAccountCollateral> vwAccountCollateral { get; set; }
		public List<displayOSBalance> vwOSBalance { get; set; }
		public List<displayCollection> vwCollection { get; set; }
		public List<displayPostDated> vwPostDated { get; set; }
		public List<displayOverdue> vwOverdue { get; set; }
		public List<displayRepayment> vwRepayment { get; set; }
		public List<displayRepaymentTable> vwRepaymentTable { get; set; }
		public List<displayRepaymentDetail> vwRepaymentDetail { get; set; }
		public List<displayLoanHist> vwLoanHistory { get; set; }
		public List<displayLoanSummary> vwLoanSummary { get; set; }
		public List<displayLNInitialRate> vwLoanInitRate { get; set; }

	}

	public class displayCustomer
	{
		public int cifkey { get; set; }
		public string branch { get; set; }
		public string idno { get; set; }
		public DateTime cifcrtdt { get; set; }
		public DateTime lstupdt { get; set; }
		public string custcat { get; set; }
		public string title { get; set; }
		public string fsname { get; set; }
		public string mdname { get; set; }
		public string lsname { get; set; }
		public string lname { get; set; }
		public string sname { get; set; }
		public string sex { get; set; }
		public string maritsts { get; set; }
		public string presentadd { get; set; }
		public string preseq { get; set; }
		public string permanentadd { get; set; }
		public string perseq { get; set; }
		public string officeadd { get; set; }
		public string offseq { get; set; }
		public string otheradd { get; set; }
		public string othseq { get; set; }
		public string telxadd { get; set; }
		public string tlxseq { get; set; }
		public string emailadd { get; set; }
		public string emseq { get; set; }
		public string corradrseq { get; set; }
		public DateTime birthdt { get; set; }
		public string age { get; set; }
		public string nation { get; set; }
		public string cntryorigin { get; set; }
		public string costcentre { get; set; }
		public string corrlan { get; set; }
		public string pname { get; set; }
		public string employer { get; set; }
		public string custsts { get; set; }
		public string staff { get; set; }
		public string blklisted { get; set; }
		public string foreigner { get; set; }
		public string sourceoffund { get; set; }
		public string hometel { get; set; }
		public string mobile { get; set; }
		public string office { get; set; }
		public string fax { get; set; }
		public string othertel1 { get; set; }
		public string othertel2 { get; set; }
		public string taxno { get; set; }
		public string ssno { get; set; }
		public string custcategory { get; set; }
		public string idtype { get; set; }
		public string indcode { get; set; }
		public string sectorcode { get; set; }
		public string relofficer { get; set; }
		public string relofficer2 { get; set; }
		public string acofficer { get; set; }
		public string custclass { get; set; }
		public string lrgcust { get; set; }
		public string comstmt { get; set; }
		public DateTime relbnkdt { get; set; }
		public string crrating { get; set; }

	}

	public class displayAccountInfo
	{
		public string acsts { get; set; }
		public string acstsdesc { get; set; }
		public string product { get; set; }
		public string prodescr { get; set; }
	}

	public class displayAccountCollateral
	{
		public string eqpdesc { get; set; }
		public string make { get; set; }
		public string model { get; set; }
		public string unitcode { get; set; }
		public double eqpcost { get; set; }
		public double finamt { get; set; }
		public string eqpcond { get; set; }
		public string rcflg { get; set; }
		public double markval { get; set; }
		public double realval { get; set; }
		public double insamt { get; set; }
		public DateTime repodt { get; set; }
		public string engineno { get; set; }
		public string chassisno { get; set; }
		public string serregno { get; set; }
		public string yrmanu { get; set; }
		public string yrmanu2 { get; set; }
		public string gdsts { get; set; }
		public string typofcover { get; set; }
		public DateTime insexpdt { get; set; }

	}

	public class displayOSBalance
	{
		public double grsbal { get; set; }
		public double prinbal { get; set; }
		public double intbal { get; set; }
		public double othos { get; set; }
		public double unern { get; set; }
		public double netbal { get; set; }
		public double agcunern { get; set; }

		public string riskflg { get; set; }
		public string suspintflg { get; set; }
		public string spcproflg { get; set; }
		public string fourthschedule { get; set; }
		public string niissued { get; set; }
		public string roissued { get; set; }
		public string blklisted { get; set; }
		public string litigatsts { get; set; }
		public string frezmode { get; set; }
		public DateTime frezdt { get; set; }
		public DateTime terminatedt { get; set; }
		public string variation { get; set; }
		public string nplflg { get; set; }
		public DateTime npldt { get; set; }
		public string nplsts { get; set; }
		public string remarks { get; set; }
	}

	public class displayCollection
	{
		public double installment { get; set; }
		public double principal { get; set; }
		public double interest { get; set; }
		public double grace_int { get; set; }
		public double month_col { get; set; }
		public double int_serv { get; set; }
		public double pri_serv { get; set; }
		public double part_setl { get; set; }
		public double advamt { get; set; }

		public DateTime last_coldt { get; set; }
		public double prepdno { get; set; }
		public double prepdamt { get; set; }
		public string grcpay { get; set; }
		public double prepdrev { get; set; }
		public string adminofficer { get; set; }

	}

	public class displayPostDated
	{
		public string bnkcd { get; set; }
		public DateTime chqdt { get; set; }
		public string chqno { get; set; }
		public double chqamt { get; set; }
	}

	public class displayOverdue
	{
		public double arrear_amt { get; set; }
		public double prin_arrear { get; set; }
		public DateTime nxt_repay_due { get; set; }

		public double arrear_no { get; set; }
		public double arrear_mth { get; set; }
		public double arrear_int { get; set; }
		public double nxt_repay_amt { get; set; }

		public string rem_typ { get; set; }
		public DateTime issue_dt { get; set; }
		public string no_nii { get; set; }
		public string no_ro { get; set; }

		public DateTime f4thschdt { get; set; }
		public DateTime ni_dt { get; set; }
		public DateTime ro_dt { get; set; }
		public DateTime penalty_acc { get; set; }
		public double penalty_acc_amt { get; set; }
		public string notday1 { get; set; }

	}

	public class displayRepayment
	{
		public string reptyp { get; set; }
		public string reptypdesc { get; set; }
		public string repcat { get; set; }
		public string repcatdesc { get; set; }
		public string repscheme { get; set; }
		public string repschemedesc { get; set; }
		public string tstmth { get; set; }

	}

	public class displayRepaymentTable
	{
		public double freq { get; set; }
		public string repschtyp { get; set; }
		public double frmprd { get; set; }
		public double toprd { get; set; }
		public double repamt { get; set; }
		public double dueday { get; set; }
		public double prcamt { get; set; }
		public double pgiamt { get; set; }
	}

	public class displayRepaymentDetail
	{
		public DateTime duedt { get; set; }
		public double amtdue { get; set; }
		public double intdue { get; set; }
		public double amtpd { get; set; }
		public double intpd { get; set; }
		public double prnpd { get; set; }
		public string repcd { get; set; }
		public DateTime trndt { get; set; }
		public int latefor { get; set; }
	}

	public class displayLoanHist
	{
		public int ID { get; set; }
		public DateTime TRNDT { get; set; }
		public int SEQNO { get; set; }
		public string TRNCD { get; set; }
		public string ITEMCD { get; set; }
		public double TRNAMT { get; set; }
		public double BALUSED { get; set; }
		public string DRTYP { get; set; }
		public string USRID { get; set; }
		public string SVRID { get; set; }
		public string SVRID2 { get; set; }
		public double INTAMT { get; set; }
		public string REFNO { get; set; }
		public string RECEIPTNO { get; set; }
		public string TRNDESCR { get; set; }
		public string GLACNO { get; set; }
		public string GLACNO2 { get; set; }
		public DateTime ACTUALDT { get; set; }
		public string BRNCD { get; set; }
		public string BDSBRNCD { get; set; }
		public string BDSTRNCD { get; set; }
		public string BDSJNLNO { get; set; }
		public double BALANCE { get; set; }
		public double DBBAL { get; set; }
	}

	public class displayLoanSummary
	{
		public double lnamt { get; set; }
		public double prd { get; set; }
		public double grcterm { get; set; }
		public double effrt { get; set; }
		public double effyield { get; set; }
		public string commencecd { get; set; }
		public string revolveflg { get; set; }
		public string payint { get; set; }
		public string pgiupfront { get; set; }
		public double pgirt { get; set; }
		public string pgirtcd { get; set; }
		public string pgirest { get; set; }
		public string pgiovdrttyp { get; set; }
		public double pgiovdrt { get; set; }
		public string pgiprovb { get; set; }
		public double pgiovdgrc { get; set; }
		public double minovdamt { get; set; }
		public string grcpay { get; set; }
		public double intrt { get; set; }
		public string intrtcd { get; set; }
		public string resttyp { get; set; }
		public string ovdrttyp { get; set; }
		public double ovdrt { get; set; }
		public string afrprovb { get; set; }
		public double ovdgrc { get; set; }
	}

	public class displayLNInitialRate
	{
		public string typ { get; set; }
		public string frprd { get; set; }
		public string toprd { get; set; }
		public string rtcd { get; set; }
		public double repamt { get; set; }
		public double basert { get; set; }
		public double mrt { get; set; }
		public double totrt { get; set; }
		public double intamt { get; set; }
		public double openbal { get; set; }
	}

	public class MasterCASAFile
	{
		public List<displayCustomer> vwCustomerInfo { get; set; }
		public List<dispCASATotal> vwTotalBalance { get; set; }
		public List<dispCASAHistory> vwCASAHistory { get; set; }
	}

	public class dispCASATotal
	{
		public double avbal { get; set; }
		public double lgrbal { get; set; }
		public double psbkbal { get; set; }
		public double fltamt { get; set; }
		public double markamt { get; set; }
		public double mcramt { get; set; }
		public double mdramt { get; set; }
		public DateTime lstdt { get; set; }
		public string lstseqno { get; set; }
		public string oldacno { get; set; }
		public string frozen { get; set; }
		public string frezovr { get; set; }
		public DateTime frezexpdt { get; set; }
		public double adb { get; set; }
	}

	public class dispCASAHistory
	{
		public int ID { get; set; }
		public string TRNBRN { get; set; }
		public string TRNCD { get; set; }
		public DateTime TRNDT { get; set; }
		public int SEQNO { get; set; }
		public int SEQCNT { get; set; }
		public string BRSTN { get; set; }
		public string REFDOCNO { get; set; }
		public double TRNAMT { get; set; }
		public double BALANCE { get; set; }
		public DateTime TIME { get; set; }
		public int CIFKEY { get; set; }
		public string POST { get; set; }
		public string MISC { get; set; }
		public string TLRID { get; set; }
		

	}

	public class MasterFDFile
	{
		public List<displayCustomer> vwCustomerInfo { get; set; }
		public List<dispFDCertInfo> vwCertInfo { get; set; }
		public List<dispFDHistory> vwFDHistory { get; set; }
		public List<dispFDSummary> vwFDSummary { get; set; }
	}

	public class dispFDCertInfo
	{
		public string crtno { get; set; }
		public double prcamt { get; set; }
		public string period { get; set; }
		public string fldchar { get; set; }
		public DateTime issdt { get; set; }
		public DateTime effdt { get; set; }
		public DateTime duedt { get; set; }
		public DateTime wdldt { get; set; }
		public string intcd { get; set; }
		public string creditno { get; set; }
		public string certsts { get; set; }
		public string prodcd { get; set; }
		public double intrt { get; set; }
		public double totint { get; set; }
		public DateTime intpddt { get; set; }
		public double intpaid { get; set; }
		public double whtax { get; set; }
		public string advice { get; set; }
		public string certsr { get; set; }

	}

	public class dispFDHistory
	{
		public int ID { get; set; }
		public DateTime TRNDT { get; set; }
		public string TRNCD { get; set; }
		public string ITEMCD { get; set; }
		public int SEQNO { get; set; }
		public int SEQCNT { get; set; }
		public string ERRFLG { get; set; }
		public string CRTNO { get; set; }
		public double TRNAMT { get; set; }
		public double BALANCE { get; set; }
		public string DR { get; set; }
	}

	public class dispFDSummary
	{
		public string prodcd { get; set; }
		public double certbal { get; set; }
		public double mdepamt { get; set; }
		public double mdepcnt { get; set; }
		public double mwdlamt { get; set; }
		public double mwdlcnt { get; set; }
		public double fltamt { get; set; }
		public double mincetvpd { get; set; }
		public double mzakatpd { get; set; }
	}

	//INQUIRY BANKWORKS
	public class InquiryPrepaid
	{
#pragma warning disable CS8632 // The annotation for nullable reference types should only be used in code within a '#nullable' annotations context.
		public string? CardNo { get; set; }
		public string? FirstName { get; set; }
		public string? LastName { get; set; }
#pragma warning restore CS8632 // The annotation for nullable reference types should only be used in code within a '#nullable' annotations context.

	}

	public class InquiryBW
	{
		public string CARD_NUMBER { get; set; }
		public string CNO { get; set; }
		public string EMBOSS_LINE_1 { get; set; }
		public string CARD_STATUS { get; set; }
		public string EXPIRY_DATE { get; set; }
		public string ACCT_NUMBER { get; set; }
		public string LAST_NAME { get; set; }
		public string FIRST_NAME { get; set; }
		public string service_contract_id { get; set; }
		public string cstat { get; set; }
		public string balance { get; set; }
		public string note_text { get; set; }
	}

	public class MasterBWFile
	{
		public List<bwClienInfo> BWClientInfo { get; set; }
		public List<bwAddress> BWAddress { get; set; }
		public List<bwAccountDetails> BWAccountDetails { get; set; }
		public List<bwHistory> BWHistory { get; set; }
		public List<bwCardUsage> BWCardUsage { get; set; }
		public List<bwAuthorization> BWOnlineAuth { get; set; }
		public List<bwContactInfo> BWContactInfo { get; set; }
		public List<bwStatus> BWStatus { get; set; }
		public List<bwBalance> BWBalance { get; set; }
	}


	public class bwClienInfo
	{
		public string salutation { get; set; }
		public string first_name { get; set; }
		public string last_name { get; set; }
		public string BIRTH_NAME { get; set; }
		public string SHORT_NAME { get; set; }
		public string tin { get; set; }
		public string sss { get; set; }
		public string PASSPORT_NUMBER { get; set; }
		public string birth_date { get; set; }
		public string BIRTH_PLACE { get; set; }
		public string NATIONALITY { get; set; }
		public string CLIENT_LANGUAGE { get; set; }
		public string CARD_STATUS { get; set; }
		public string CLIENT_BRANCH { get; set; }
		public string TEL_PRIVATE { get; set; }
		public string TEL_WORK { get; set; }
		public string FAX_PRIVATE { get; set; }
		public string FAX_WORK { get; set; }
		public string CLIENT_PASSWORD { get; set; }
		public string CIFNO { get; set; }
		public string EMAIL_ADDR { get; set; }
		public string tel_home { get; set; }
		public string tel_other { get; set; }
		public string ADDR_LINE_1 { get; set; }
		public string ADDR_LINE_2 { get; set; }
		public string ADDR_LINE_3 { get; set; }
		public string ADDR_CLIENT_CITY { get; set; }
		public string POST_CODE { get; set; }
		public string client_country { get; set; }
		public string delivery_method { get; set; }

	}
	public class bwAddress
	{
		public string address_category { get; set; }
		public string effective_date { get; set; }
		public string expiry_date { get; set; }
		public string addr_line_1 { get; set; }
		public string addr_line_2 { get; set; }
		public string addr_line_3 { get; set; }
		public string addr_line_4 { get; set; }
		public string addr_line_5 { get; set; }
		public string post_code { get; set; }
		public string addr_client_city { get; set; }
		public string client_country { get; set; }
		public string contact_name { get; set; }
		public string tel_work { get; set; }
		public string fax_work { get; set; }

	}

	public class bwAccountDetails
	{
		public string card_number { get; set; }
		public string emboss_line_1 { get; set; }
		public string emboss_line_2 { get; set; }
		public string CARD_STATUS { get; set; }
		public string last_issued_date { get; set; }
		public string effective_date { get; set; }
		public string expiry_date { get; set; }
		public string username { get; set; }
	}

	public class bwHistory
	{
		public string effective_date { get; set; }
		public string time { get; set; }
		public string expiry_date { get; set; }
		public string original_card_status { get; set; }
		public string card_status { get; set; }
		public string note_text { get; set; }
		public string pin_required { get; set; }
		public string production_status { get; set; }
		public string fee_charged { get; set; }
		public string prod_file_number { get; set; }
		public string card_event_status { get; set; }
		public string delivery_method { get; set; }
		public string file_number { get; set; }
		public string last_fee_date { get; set; }
		public string username { get; set; }
	}

	public class bwCardUsage
	{
		public string transaction_type { get; set; }
		public string wrong_pin_tries { get; set; }
		public string wrong_pin_limit { get; set; }
		public string daily_limit_amt { get; set; }
		public string daily_limit_freq { get; set; }
		public string daily_usage_amt { get; set; }
		public string daily_usage_freq { get; set; }
		public string window_limit_amt { get; set; }
		public string window_limit_freq { get; set; }
		public string window_usage_amt { get; set; }
		public string window_usage_freq { get; set; }
		public string last_amendment_date { get; set; }
		public string accum_currency { get; set; }
		public string window_no_of_days { get; set; }

	}

	public class bwAuthorization
	{
		public string TRANSACTION_DATE { get; set; }
		public string TIME_TRANSACTION { get; set; }
		public string MERCHANT_NAME { get; set; }
		public string MERCHANT_COUNTRY { get; set; }
		public string DR_CR_INDICATOR { get; set; }
		public string TRANSACTION_TYPE { get; set; }
		public string ACCOUNT_AMOUNT_GR { get; set; }
		public string ACCT_CURRENCY { get; set; }
		public string RESPONSE { get; set; }
		public string TRAN_AMOUNT_GR { get; set; }
		public string TRAN_CURRENCY { get; set; }
		public string RATE_FX_LOCAL_ACCT { get; set; }
		public string TRANSACTION_STATUS { get; set; }
		public string REVERSAL_FLAG { get; set; }
		public string TRANSACTION_CATEGORY { get; set; }
		public string TRANSACTION_CLASS { get; set; }
		public string CARD_NUMBER { get; set; }
		public string AUTH_CODE { get; set; }
		public string BUSINESS_CLASS { get; set; }
		public string TRANSACTION_SOURCE { get; set; }
		public string ACCOUNT_AMOUNT_NET { get; set; }
		public string ACCOUNT_AMOUNT_CHG { get; set; }
		public string TRANSACTION_SLIP { get; set; }
		public string retrieval_reference { get; set; }
		public string transaction_id { get; set; }
		public string pan_entry { get; set; }
	}

	public class bwContactInfo
	{
		public string service_contract { get; set; }
		public string client_tariff { get; set; }
		public string settlement_method { get; set; }
		public string account_level { get; set; }

	}
	public class bwStatus
	{
		public string client_status { get; set; }
		public string client_date { get; set; }
		public string card_status { get; set; }
		public string card_date { get; set; }
		public string acct_status { get; set; }
		public string acct_date { get; set; }
		public string contract_status { get; set; }
		public string contract_date { get; set; }
		public string limit_currency { get; set; }
		public string client_limit { get; set; }

	}

	public class bwBalance
	{
		public string swift_code { get; set; }
		public string begin_balance { get; set; }
		public string current_balance { get; set; }
		public string pending_auths { get; set; }
		public string availability { get; set; }
	}

	public class IRemitCard
	{
		[Key]
		public int ID { get; set; }
		public string CardNo { get; set; }
	}


	public class ImportCard
	{
		public string CardNo { get; set; }
	}

}
