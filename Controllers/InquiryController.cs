using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using AgentDesktop.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Globalization;
using AgentDesktop.Services;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using System.Data;

namespace AgentDesktop.Controllers
{
	public class InquiryController : Controller
	{

		private readonly AgentInterface _db;
		private readonly IWebHostEnvironment _webHostEnvironment;

		public InquiryController(AgentInterface db, IWebHostEnvironment webHostEnvironment)
		{
			_db = db;
			_webHostEnvironment = webHostEnvironment;
		}

		public IActionResult Index()
		{

			return View();
		}

		[HttpGet]
		public IActionResult InquiryICBA()
		{
			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			bool usrights = _db.ChkRights(ses.LoginID, "Inquiry", "InquiryICBA");

			if (!usrights)
			{
				TempData["msgAuth"] = "You are not authorized to use this feature";
				return RedirectToAction("Index", "Home");
			}

			//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

			bool isLog = _db.GetAuditLogs("Inquiry ICBA", ses.WrkStn, "Search Customer Details in ICBA");

			return View();
		}

		[HttpPost]
		public IActionResult InquiryICBA(string CardNo, string FirstName, string LastName, string FullName)
		{
			ViewBag.Page = "ICBA";
			string fil;

			fil = (((string.IsNullOrEmpty(CardNo)) ? "0" : "1")
				+ (((string.IsNullOrEmpty(FirstName)) ? "0" : "1")
				+ (((string.IsNullOrEmpty(LastName)) ? "0" : "1")
				+ ((string.IsNullOrEmpty(FullName)) ? "0" : "1"))));

			if (fil == "0000")
			{
				TempData["msginfo"] = "Please enter one of the parameters.";

			}
			else
			{

				ViewBag.Inquiry = "Meron";

				ViewBag.ListICBAClients = _db.ListInquiryICBA(fil, CardNo ?? "", FirstName == null ? "" : FirstName.Trim().ToUpper(), LastName == null ? "" : LastName.Trim().ToUpper(), FullName == null ? "" : FullName.Trim().ToUpper());

			}

			return View();
		}

		public IActionResult InquiryRedirect(string actno, string crline, string cif)
		{

			switch (crline)
			{
				case "SA":
					return RedirectToAction("InquirySavings", "Inquiry", new { cif = cif, acc = actno });
				case "CA":
					return RedirectToAction("InquiryCurrent", "Inquiry", new { cif = cif, acc = actno });
				case "FD":
					return RedirectToAction("InquiryFixed", "Inquiry", new { cif = cif, acc = actno });
				case "LN":
					return RedirectToAction("InquiryLoans", "Inquiry", new { cif = cif, acc = actno });
			}

			return View();
		}

		public IActionResult InquiryLoans(string cif, string acc)
		{

			HttpContext.Session.SetString("accountno", acc);
			HttpContext.Session.SetString("cifkey", cif);

			MasterDetailFile LoanData = new MasterDetailFile();

			List<MasterDetailFile> MasterData = _db.GetLoanDetails(acc, cif).ToList();

			LoanData.vwCustomerInfo = MasterData[0].vwCustomerInfo;
			LoanData.vwAccountInfo = MasterData[0].vwAccountInfo;
			LoanData.vwAccountCollateral = MasterData[0].vwAccountCollateral;
			LoanData.vwOSBalance = MasterData[0].vwOSBalance;
			LoanData.vwCollection = MasterData[0].vwCollection;
			LoanData.vwPostDated = MasterData[0].vwPostDated;
			LoanData.vwOverdue = MasterData[0].vwOverdue;
			LoanData.vwRepayment = MasterData[0].vwRepayment;
			LoanData.vwRepaymentTable = MasterData[0].vwRepaymentTable;
			LoanData.vwRepaymentDetail = MasterData[0].vwRepaymentDetail;
			LoanData.vwLoanHistory = MasterData[0].vwLoanHistory;
			LoanData.vwLoanSummary = MasterData[0].vwLoanSummary;
			LoanData.vwLoanInitRate = MasterData[0].vwLoanInitRate;

			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

			bool isLog = _db.GetAuditLogs("Inquiry ICBA - Loans", ses.WrkStn, "View Loans Customer Details in ICBA");

			return View(LoanData);
		}

		public IActionResult InquirySavings(string cif, string acc)
		{

			HttpContext.Session.SetString("accountno", acc);
			HttpContext.Session.SetString("cifkey", cif);

			MasterCASAFile SavingsData = new MasterCASAFile();

			List<MasterCASAFile> MasterData = _db.GetCASADetails(acc, cif).ToList();

			SavingsData.vwCustomerInfo = MasterData[0].vwCustomerInfo;
			SavingsData.vwTotalBalance = MasterData[0].vwTotalBalance;
			SavingsData.vwCASAHistory = MasterData[0].vwCASAHistory;

			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

			bool isLog = _db.GetAuditLogs("Inquiry ICBA - Savings", ses.WrkStn, "View Savings Customer Details in ICBA");


			return View(SavingsData);
		}

		public IActionResult InquiryCurrent(string cif, string acc)
		{

			HttpContext.Session.SetString("accountno", acc);
			HttpContext.Session.SetString("cifkey", cif);

			MasterCASAFile CurrentData = new MasterCASAFile();

			List<MasterCASAFile> MasterData = _db.GetCASADetails(acc, cif).ToList();

			CurrentData.vwCustomerInfo = MasterData[0].vwCustomerInfo;
			CurrentData.vwTotalBalance = MasterData[0].vwTotalBalance;
			CurrentData.vwCASAHistory = MasterData[0].vwCASAHistory;

			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

			bool isLog = _db.GetAuditLogs("Inquiry ICBA - Current", ses.WrkStn, "View Current Customer Details in ICBA");


			return View(CurrentData);
		}

		public IActionResult InquiryFixed(string cif, string acc)
		{

			HttpContext.Session.SetString("accountno", acc);
			HttpContext.Session.SetString("cifkey", cif);

			MasterFDFile FixedData = new MasterFDFile();

			List<MasterFDFile> MasterData = _db.GetFDDetails(acc, cif).ToList();

			FixedData.vwCustomerInfo = MasterData[0].vwCustomerInfo;
			FixedData.vwCertInfo = MasterData[0].vwCertInfo;
			FixedData.vwFDHistory = MasterData[0].vwFDHistory;
			FixedData.vwFDSummary = MasterData[0].vwFDSummary;

			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

			bool isLog = _db.GetAuditLogs("Inquiry ICBA - Fixed Deposit", ses.WrkStn, "View Fixed Deposit Customer Details in ICBA");


			return View(FixedData);
		}

		[HttpGet]
		public IActionResult InquiryPrepaid()
		{
			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			bool usrights = _db.ChkRights(ses.LoginID, "Inquiry", "InquiryPrepaid");

			if (!usrights)
			{
				TempData["msgAuth"] = "You are not authorized to use this feature";
				return RedirectToAction("Index", "Home");
			}

			//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

			bool isLog = _db.GetAuditLogs("Inquiry Prepaid", ses.WrkStn, "View Prepaid Customers in Bankworks");


			return View();
		}

		[HttpPost]
		public IActionResult InquiryPrepaid(string CardNo, string FirstName, string LastName, IFormFile files)
		{
			ViewBag.Page = "Prepaid";

			if (files == null)
			{
				string fil;

				fil = (((string.IsNullOrEmpty(CardNo)) ? "0" : "1")
					+ (((string.IsNullOrEmpty(FirstName)) ? "0" : "1")
					+ ((string.IsNullOrEmpty(LastName)) ? "0" : "1")));

				if (fil == "000")
				{
					TempData["msginfo"] = "Please enter one of the parameters.";
				}
				else
				{
					HttpContext.Session.SetString("fil", fil);
					HttpContext.Session.SetString("carding", CardNo ?? "");
					HttpContext.Session.SetString("fname", FirstName == null ? "" : FirstName.Trim().ToUpper());
					HttpContext.Session.SetString("lname", LastName == null ? "" : LastName.Trim().ToUpper());
					HttpContext.Session.SetString("posting", "A");

					ViewBag.InqBW = "Meron";

					ViewBag.ListBWClients = _db.ListInquiryBW(fil, CardNo ?? "", FirstName == null ? "" : FirstName.Trim().ToUpper(), LastName == null ? "" : LastName.Trim().ToUpper());

				}
			}
			else
			{
					HttpContext.Session.SetString("posting", "B");

					string path = _db.Documentupload(files);

					bool import_excel = _db.ImportExcel(path);

					if (!import_excel)
					{
						TempData["msgErr"] = "Worksheet out of range";
						return View();
					}

					ViewBag.InqBW = "Meron";

					ViewBag.ListBWClients = _db.ListInquiryBWIRemit();

			}



			return View();
		}

		public IActionResult InquiryBW(string actno, string cardno, string st, string et)
		{

			HttpContext.Session.SetString("acctno", actno);
			HttpContext.Session.SetString("cardno", cardno);

			MasterBWFile BWData = new MasterBWFile();

			List<MasterBWFile> MasterData = _db.GetBWDetails(actno, cardno, st, et).ToList();

			BWData.BWClientInfo = MasterData[0].BWClientInfo;
			BWData.BWAddress = MasterData[0].BWAddress;
			BWData.BWAccountDetails = MasterData[0].BWAccountDetails;
			BWData.BWHistory = MasterData[0].BWHistory;
			BWData.BWCardUsage = MasterData[0].BWCardUsage;
			BWData.BWOnlineAuth = MasterData[0].BWOnlineAuth;
			BWData.BWContactInfo = MasterData[0].BWContactInfo;
			BWData.BWStatus = MasterData[0].BWStatus;
			BWData.BWBalance = MasterData[0].BWBalance;

			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

			bool isLog = _db.GetAuditLogs("Inquiry Prepaid", ses.WrkStn, "View Prepaid Customer Details in Bankworks");

			return View(BWData);
		}

		public JsonResult GetAuthorization(string dt1, string dt2)
		{

			HttpContext.Session.SetString("dt1", dt1 ?? "");
			HttpContext.Session.SetString("dt2", dt2 ?? "");

			return Json(new { success = true });
		}

		public IActionResult OnlineAuth()
		{
			string _dt1 = HttpContext.Session.GetString("dt1");
			string _dt2 = HttpContext.Session.GetString("dt2");
			string _card = HttpContext.Session.GetString("cardno");
			string _acct = HttpContext.Session.GetString("acctno");

			return RedirectToAction("InquiryBW", "Inquiry", new { actno = _acct, cardno = _card, st = _dt1, et = _dt2 } );
			
		}

		public IActionResult ExportPrepaid()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var posting = HttpContext.Session.GetString("posting");
			var fil = HttpContext.Session.GetString("fil");
			var carding = HttpContext.Session.GetString("carding");
			var fname = HttpContext.Session.GetString("fname");
			var lname = HttpContext.Session.GetString("lname");

			string fnam = string.Empty;

			var v = (posting == "A") ? _db.ListInquiryBW(fil, carding, fname, lname) : _db.ListInquiryBWIRemit();

			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("InquiryPrepaid", "Inquiry");
			}


			using (ExcelPackage package = new(stream))
			{

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.CARD_NUMBER.Contains(myInput) ||
								 p.EMBOSS_LINE_1.ToLower().Contains(myInput.ToLower()) ||
								 p.CARD_STATUS.ToLower().Contains(myInput.ToLower()) ||
								 p.note_text.ToString().Contains(myInput) ||
								 p.EXPIRY_DATE.Contains(myInput) ||
								 p.balance.Contains(myInput) 
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "CARD NUMBER";
				ws.Cells["B1"].Value = "ACCOUNT NAME";
				ws.Cells["C1"].Value = "CARD STATUS";
				ws.Cells["D1"].Value = "NOTE";
				ws.Cells["E1"].Value = "EXPIRY DATE";
				ws.Cells["F1"].Value = "ACCOUNT BALANCE";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.CARD_NUMBER;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.EMBOSS_LINE_1;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.CARD_STATUS;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.note_text;
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.EXPIRY_DATE;
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.balance;
					rowStart++;
				}

				ws.Cells["A:F"].AutoFitColumns();

				package.Save();
			}

			fnam = "Filter Inquiry - Prepaid Cards " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

	}
}
