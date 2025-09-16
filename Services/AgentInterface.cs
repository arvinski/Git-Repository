using System;
using System.Collections.Generic;
using System.Net;
using AgentDesktop.Models;
using Microsoft.AspNetCore.Http;

namespace AgentDesktop.Services
{
	public interface AgentInterface : IDisposable
	{
		#region Login
		bool ChkUser(string _usr);
		EmailForgot GetEmail(string _usr);
		bool SendEmailForgot(string _email, string _pass, string _fname);
		bool UpdateStatus(string _userid, string _pass);
		bool ChkPass(string _usr, string _pass);
		LoginPass GetLoginID(string usrName, string usrPass);
		bool LockUser(string _username);
		bool ChkTmpPass(string _userid, string _pass);
		bool ChkPassList(string _pass);
		bool ChkPassListSeq(string _pass);
		bool ChkPassName(string _id, string _pass);
		CountPass ChkPassCount(string _id, string _pass);
		bool UpdatePass(string _username, string _pass, int _usrtyp);
		bool UserClosed(int _inactive, int _unlock);
		bool LogUserOut(int _userid, string _mod);
		bool Activity(string _loginid);
		bool SendEmailUnlock(string _name, string _email, string _url, string _usrname);
		bool UserUnlock(string _username);

		#endregion

		#region Admin
		IEnumerable<ViewUsers> ListUsers();
		List<GroupName> ListGroup();
		List<Departments> ListDept();
		bool CreateUser(NewUser usr);
		UserPass GetUserPass(string _usr);
		bool SendEmailNewUser(string _usrid, string _email, string _fname, string _pass);
		EditUser ViewUsers(int _id);
		bool UpdateUser(EditUser edt);
		bool DeleteUser(int _id);
		IEnumerable<ReasonCode> ListReasonCodes();
		bool ChkRCode(string _rcd);
		bool CreateReasonCode(NewReason rcd);
		EditReason ViewReasonCode(int _id);
		bool UpdateReasonCode(EditReason edt);
		bool ChkRCodeActivity(string _rcd);
		bool DeleteReason(int _id);
		IEnumerable<ATM> ListATM(string _typ);
		bool CreateATM(string _desc, string _typ);
		bool ChkATM(string _desc, string _typ);
		EditATM ViewATM(int _id);
		bool UpdateATM(EditATM edt);
		bool ChkPaymentOfActivity(string _desc);
		bool DeleteATM(int _id);
		bool DeleteATMLocation(int _id, int _loc_id);
		bool ChkBancnetActivity(string _desc);
		bool ChkSubStatusActivity(string _desc);
		bool ChkRemarksActivity(string _desc);
		bool ChkReferredByActivity(string _desc);
		bool ChkATMActivity(string _desc);
		IEnumerable<ATM_GROUP_LOC> ListATMLocation();
		bool CreateATMLocation(int _id, string _desc);
		EditATMLoc ViewATMLocation(int _id);
		bool UpdateATMLocation(string _atmloc, int _locid);
		IEnumerable<UserMenu> ListUserMenu();
		bool ChkMenu(string _view, string _page);
		bool CreateMenu(NewMenu mnu);
		EditMenu ViewMenu(int _id);
		bool UpdateMenu(EditMenu edt);
		bool DeleteMenu(int _id);
		bool ChkEndorsedActivity(string _desc);
		IEnumerable<UserRight> ListUserRights(int id);
		bool InsertRights(int _id);
		bool UpdateRights(int _id, string _idd);
		bool ChkRights(string _uid, string _vw, string _pg);
		bool GetAuditLogs(string _description, string _workstn, string _activity);
		IEnumerable<Audit_Logs> ListAuditLogs(string _dt1, string _dt2);

		#endregion

		#region Home
		bool Populate_Dashboard(string _id);
		IEnumerable<ServiceTicket> ListServiceTicket(string _loginid, string _dep);
		IEnumerable<ServiceTicket> ListServiceStatus(string _loginid, string _dep);
		IEnumerable<Escalation> ListEscalation();
		IEnumerable<ServiceReason> ListServiceReason(string _loginid, string _dep);
		IEnumerable<ServiceChannel> ListServiceChannel(string _loginid, string _dep);
		IEnumerable<DisplayServiceTicket> DisplayServiceTickets(string _loginid, string _dept, int _utype, int _cond);
#nullable enable
		IEnumerable<DisplayServiceTicketDated> View_TransStatus(string _LoginID, string _Stat, int _Mod, string? _Date1, string? _Date2);
		IEnumerable<DisplayServiceTicketDated> View_TransChannel(string _LoginID, string _Stat, int _Mod, string? _Date1, string? _Date2);
		IEnumerable<DisplayServiceTicketDated> View_TransReason(string _LoginID, string _Stat, int _Mod, string? _Date1, string? _Date2);

		IEnumerable<DisplayServiceTicket> View_TransEscalate(int _cond);

		#endregion

		#region Activity
		//IEnumerable<CMActivity> View_CMActivity(string _dt1, string _dt2, string _rno);
		IEnumerable<CMActivity> View_CMActivity(ActivitySearch _sch);
		IEnumerable<CMActivity> View_CMActivityHist(ActivitySearch _sch);
		IEnumerable<ListActivity> ListActivity(string _ticketno, string _dep);
#nullable disable
		List<Combo1> ListReasonCode();
		List<Combo1> Get_Dropdown(string typ);
		List<Combo2> Get_Dropdown2(string typ);
		List<Combo1> Get_Dropdown3();
		List<Combo2> Get_Dropdown4();
		IEnumerable<CMActivity> View_Ticket(string _ticketno);
		IEnumerable<SearchICBA> Search_ICBA(string _clientname);
		IEnumerable<SearchBW> Search_BW(string _clientname);
		List<Combo3> Get_Dropdown_Location(int _id);
		IEnumerable<ATMLogs> ListATMLogs();
		IEnumerable<ATMOps> ListATMOps(string _dt1, string _dt2);
		IEnumerable<SearchCardNo> Search_Card(string _cardno);
		bool CreateNewActivity(string CommChannel, string ReasonCode, string CustomerName, string CardNo,
			string Activity, string Status, string Agent, string TicketNo, string LastName, string FirstName, string MiddleName, string InOutBound,
			string TransType, string Location, string ATMUsed, string PaymentOf, string TerminalUsed, string Merchant, string Inter, string BancnetUsed, string Online,
			string Website, double Amount, string DateTrans, string PLStatus, string CardBlocked, string EndoredTo, string Destination, string ATMTrans,
			string RemitFrom, string RemitConcern, string EndorsedFrom, string Resolved, string BillerName, string Remarks, string ReferedBy,
			string ContactPerson, string Tagging, string CallTime, string LocalNum, string Created_By, string Last_Update, string CardPresent, string SubStatus, string Currency, string Branch
		 );

		Show_Ticket DisplayTickets(string _rcode);
		bool UploadInputFile(string file, string ticketno);
		EditCallReport ViewCMEntry(int _id);
		IEnumerable<CMActivity> View_TicketNo(string _ticketno);
		bool SendEndorsedMail(string _tgi, string _irmt, string _emailto, string _ticketno, string _remarks, string _lastname, string _firstname, string _midname, string _cardno, string _activity, string _fullname, string _referredby, string _contactperson, List<string> attachments);
		bool SendEndorsedMailATM(string _irmt, string _ticketno, string _remarks, string _lastname, string _firstname, string _midname, string _cardno, string _activity, string _fullname, string _referredby, string _contactperson, List<string> attachments);
		bool ChkIremitCode(string _rcd);

		bool CreateNewATMActivity(string CommChannel, string ReasonCode, string CustomerName, string CardNo,
			string Activity, string Status, string Agent, string TicketNo, string LastName, string FirstName, string MiddleName, string InOutBound,
			string TransType, string Location, string ATMUsed, string PaymentOf, string TerminalUsed, string Merchant, string Inter, string BancnetUsed, string Online,
			string Website, double Amount, string DateTrans, string PLStatus, string CardBlocked, string EndoredTo, string Destination, string ATMTrans,
			string RemitFrom, string RemitConcern, string EndorsedFrom, string Resolved, string BillerName, string Remarks, string ReferedBy,
			string ContactPerson, string Tagging, string CallTime, string LocalNum, string Created_By, string Last_Update, string CardPresent, string SubStatus, string Currency, string Branch
		 );

		bool UpdateType(int _id, string _typ);
		bool DeleteCMActivity(string _ticketno);
		bool ChkActivityInATM(string _ticketno);
		bool EditNewActivity(EditCallReport edt);
		IEnumerable<ATMOps> View_ATMActivity(ActivitySearch _sch);
		bool ChkATMUsed(string _atmused);
		EditATMReport ViewATMEntry(int _id);
		IEnumerable<ATMOps> View_TicketATM(string _ticketno);

		bool CreateATMActivity(string CommChannel, string ReasonCode, string CustomerName, string CardNo,
			string Activity, string Status, string Agent, string TicketNo, string LastName, string FirstName, string MiddleName, string InOutBound,
			string TransType, string Location, string ATMUsed, string PaymentOf, string TerminalUsed, string Merchant, string Inter, string BancnetUsed, string Online,
			string Website, double Amount, string DateTrans, string PLStatus, string CardBlocked, string EndoredTo, string Destination, string ATMTrans,
			string RemitFrom, string RemitConcern, string BillerName, string Remarks, string ReferedBy,
			string ContactPerson, string Tagging, string CardPresent, string SubStatus, string Currency, string Branch, string NameVerified, string Resolved, string Last_Update, string Created_By
		 );

		IEnumerable<ATMOps> ListATMOpsResolved(string _dt1, string _dt2);
		IEnumerable<ATMOps> ListATMOpsUnResolved(string _dt1, string _dt2);
		IEnumerable<CMActivity> View_ReasonCodes(string _dt1, string _dt2, string _rno);
		MaxRID Get_RID(string _ticketno, string _dept);
		bool UpdateLocation(string _ticketno);

		#endregion

		#region Report

		IEnumerable<Combo1> ListAgent();
		IEnumerable<Summary> View_ReportSummary(string StartDate, string EndDate, string Bound, string Agent, string Type, string Comm);
		LoginPass GetAgent(int id);
		IEnumerable<ReportActivity> View_ActivityLogs(string StartDate, string EndDate, string comm_channel, string comm_type, string agent, string dept);
		IEnumerable<ReasonComm> View_ReasonComm(string StartDate, string EndDate, string agent, string dept);
		IEnumerable<ReasonStat> View_ReasonStat(string StartDate, string EndDate, string agent, string dept);
		IEnumerable<Aging> View_Aging(string StartDate, string EndDate, string RCode, string Type, string Branch);
		IEnumerable<AgingDetail> View_AgingDetail(string StartDate, string EndDate, string RCode, string Type, string Branch);
		IEnumerable<Aging> View_AgingResolved(string StartDate, string EndDate, string RCode, string Type);
		IEnumerable<ReportIRC> List_ReportIRC(string dt1, string dt2, string typ);
		IEnumerable<CMActivity> List_ReportIRCDetail(string dt1, string dt2, string typ);
		IEnumerable<ReportIRC> List_ReportIRCATM(string dt1, string dt2, string typ);
		#endregion

		#region Inquiry
		//ICBA
		IEnumerable<InquiryICBA> ListInquiryICBA(string _fil, string _cardno, string _fname, string _lname, string _fullname);
		IEnumerable<MasterDetailFile> GetLoanDetails(string _acctno, string _cifkey);
		IEnumerable<MasterCASAFile> GetCASADetails(string _acctno, string _cifkey);
		IEnumerable<MasterFDFile> GetFDDetails(string _acctno, string _cifkey);

		//BW
		IEnumerable<InquiryBW> ListInquiryBW(string _fil, string _cardno, string _fname, string _lname);
		IEnumerable<MasterBWFile> GetBWDetails(string _acctno, string _cardno, string _st, string _et);
		IEnumerable<bwAuthorization> Get_BWOnlineAuth(string _cardno, string _start, string _end);
		bool ImportExcel(string _path);

		bool ClearCards();
		bool InserToCard(string _cardno);
		string Documentupload(IFormFile fromFiles);
		IEnumerable<InquiryBW> ListInquiryBWIRemit();
		#endregion
		string GetIPAdd();
	}
}
