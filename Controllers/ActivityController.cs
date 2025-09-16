using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Http;
using AgentDesktop.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using AgentDesktop.Services;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using System.Text;
using Microsoft.Data.SqlClient;
using System.Data;

namespace AgentDesktop.Controllers
{

	public class ActivityController : Controller
	{
		private readonly AgentInterface _db;
		private readonly IWebHostEnvironment _webHostEnvironment;

		private IDbConnection con = new SqlConnection(new SqlConnectionStringBuilder()
		{
			DataSource = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppDBConnection:SQLNM"].ToString()),
			InitialCatalog = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppDBConnection:DBNM"].ToString()),
			PersistSecurityInfo = true,
			UserID = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppDBConnection:LGID"].ToString()),
			Password = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppDBConnection:PAWD"].ToString()),
			MultipleActiveResultSets = true,
		}.ConnectionString);

#nullable enable
		private string? TranDate { get; set; }
		private string rcd { get; set; }
#nullable disable



		public ActivityController(AgentInterface db, IWebHostEnvironment webHostEnvironment)
		{
			_db = db;
			_webHostEnvironment = webHostEnvironment;
		}

		[HttpGet]
		public IActionResult CustomerService()
		{
			//ITO ANG LATEST
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Activity", "CustomerService");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Activity";
				ViewBag.Page = "CustomerService";

				ActivitySearch s = new ActivitySearch();

				var vw = _db.View_CMActivity(s);

				ViewBag.ListCMActivity = vw;

				string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Customer Service - Activity", wrkstn, "View List of Customer Service Activities");

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");

				var ser = new ActivitySearch()
				{
					DateSearch = s.DateSearch,
					Date1 = s.Date1,
					Date2 = s.Date2,
					CustomerName = s.CustomerName,
					CardNo = s.CardNo,
					CommType = s.CommType,
					CommChannel = s.CommChannel,
					ReasonCode = s.ReasonCode,
					Status = s.Status,
					Agent = s.Agent,
					TicketNo = s.TicketNo,
					ReferredBy = s.ReferredBy,
					ContactPerson = s.ContactPerson
				};

				HttpContext.Session.SetString("SearchActivity", JsonConvert.SerializeObject(ser));

				return View(s);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		[HttpPost]
		public IActionResult CustomerService(ActivitySearch s)
		{
			try
			{
				ViewBag.Interface = "Activity";
				ViewBag.Page = "CustomerService";

				var vw = _db.View_CMActivity(s);

				ViewBag.ListCMActivity = vw;

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");

				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool isLog = _db.GetAuditLogs("Customer Service Activity", ses.WrkStn, "View List of Customer Service Activities");

				var ser = new ActivitySearch()
				{
					DateSearch = s.DateSearch,
					Date1 = s.Date1,
					Date2 = s.Date2,
					CustomerName = s.CustomerName,
					CardNo = s.CardNo,
					CommType = s.CommType,
					CommChannel = s.CommChannel,
					ReasonCode = s.ReasonCode,
					Status = s.Status,
					Agent = s.Agent,
					TicketNo = s.TicketNo,
					ReferredBy = s.ReferredBy,
					ContactPerson = s.ContactPerson
				};

				HttpContext.Session.SetString("SearchActivity", JsonConvert.SerializeObject(ser));
			}
			catch (Exception ex)
			{
				string controllerName = ControllerContext.ActionDescriptor.ControllerName;
				string actionName = ControllerContext.ActionDescriptor.ActionName;

				string err = "Error in " + controllerName + " - " + actionName + " : " + ex.ToString();

				TempData["msgErr"] = err;
			}

			return View(s);
		}

		public IActionResult ExportCMActivity(int id)
		{
			//try
			//{
				var stream = new System.IO.MemoryStream();

				var ses = JsonConvert.DeserializeObject<ActivitySearch>(HttpContext.Session.GetString("SearchActivity"));

				var myInput = HttpContext.Session.GetString("SearchValue");


				string fnam = string.Empty;

				var v = (id == 0) ? _db.View_CMActivity(ses) : _db.View_CMActivityHist(ses);

				int cnt = v.Count();

				if (cnt == 0)
				{
					TempData["msginfo"] = "There are no transactions to export in excel";
					return RedirectToAction("CustomerService", "Activity");
				}


				using (ExcelPackage package = new(stream))
				{

					if (!string.IsNullOrEmpty(myInput))
					{
						v = v.Where(p => p.TicketNo.ToLower().Contains(myInput.ToLower()) ||
									 p.InOutBound.ToLower().Contains(myInput.ToLower()) ||
									 p.CommChannel.ToLower().Contains(myInput.ToLower()) ||
									 p.DateReceived.ToString().Contains(myInput) ||
									 p.abbreviation.ToLower().Contains(myInput.ToLower()) ||
									 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
									 p.CardNo.ToLower().Contains(myInput.ToLower()) ||
									 p.Branch.ToLower().Contains(myInput.ToLower()) ||
									 p.Destination.ToLower().Contains(myInput.ToLower()) ||
									 p.TransType.ToLower().Contains(myInput.ToLower()) ||
									 p.ATMTrans.ToLower().Contains(myInput.ToLower()) ||
									 p.Location.ToLower().Contains(myInput.ToLower()) ||
									 p.ATMUsed.ToLower().Contains(myInput.ToLower()) ||
									 p.PaymentOf.ToLower().Contains(myInput.ToLower()) ||
									 p.TerminalUsed.ToLower().Contains(myInput.ToLower()) ||
									 p.BillerName.ToLower().Contains(myInput.ToLower()) ||
									 p.Merchant.ToLower().Contains(myInput.ToLower()) ||
									 p.Inter.ToLower().Contains(myInput.ToLower()) ||
									 p.BancnetUsed.ToLower().Contains(myInput.ToLower()) ||
									 p.Online.ToLower().Contains(myInput.ToLower()) ||
									 p.Website.ToLower().Contains(myInput.ToLower()) ||
									 p.RemitFrom.ToLower().Contains(myInput.ToLower()) ||
									 p.RemitConcern.ToLower().Contains(myInput.ToLower()) ||
									 p.Amount.ToString().Contains(myInput) ||
									 p.Currency.ToString().Contains(myInput) ||
									 p.DateTrans.ToString().Contains(myInput) ||
									 p.PLStatus.ToLower().Contains(myInput.ToLower()) ||
									 p.CardBlocked.ToLower().Contains(myInput.ToLower()) ||
									 p.EndoredTo.ToLower().Contains(myInput.ToLower()) ||
									 p.Activity.ToLower().Contains(myInput.ToLower()) ||
									 p.Status.ToLower().Contains(myInput.ToLower()) ||
									 p.SubStatus.ToLower().Contains(myInput.ToLower()) ||
									 p.Agent.ToLower().Contains(myInput.ToLower()) ||
									 p.created_by.ToLower().Contains(myInput.ToLower()) ||
									 p.Last_Update.ToLower().Contains(myInput.ToLower()) ||
									 p.EndorsedFrom.ToLower().Contains(myInput.ToLower()) ||
									 p.Resolved.ToString().Contains(myInput) ||
									 p.Remarks.ToLower().Contains(myInput.ToLower()) ||
									 p.ReferedBy.ToLower().Contains(myInput.ToLower()) ||
									 p.ContactPerson.ToLower().Contains(myInput.ToLower()) ||
									 p.ReasonCode.ToLower().Contains(myInput.ToLower()) ||
									 p.Tagging.ToLower().Contains(myInput.ToLower()) ||
									 p.ResolvedDate.ToString().Contains(myInput) ||
									 p.CallTime.ToLower().Contains(myInput.ToLower()) ||
									 p.LocalNum.ToLower().Contains(myInput.ToLower()) ||
									 p.CardPresent.ToLower().Contains(myInput.ToLower())
						);
					}

					ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


					ws.Row(1).Height = 20;
					ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					ws.Row(1).Style.Font.Bold = true;

					ws.Cells["A1"].Value = "TICKET NO";
					ws.Cells["B1"].Value = "INBOUND/OUTBOUND";
					ws.Cells["C1"].Value = "COMM CHANNEL";
					ws.Cells["D1"].Value = "DATE RECEIVED";
					ws.Cells["E1"].Value = "TIME RECEIVED";
					ws.Cells["F1"].Value = "REASON CODE";
					ws.Cells["G1"].Value = "CUSTOMER NAME";
					ws.Cells["H1"].Value = "CARD NO";
					ws.Cells["I1"].Value = "DESTINATION";
					ws.Cells["J1"].Value = "TRANS TYPE";
					ws.Cells["K1"].Value = "ATM TRANS";
					ws.Cells["L1"].Value = "LOCATION";
					ws.Cells["M1"].Value = "ATM USED";
					ws.Cells["N1"].Value = "PAYMENT OF";
					ws.Cells["O1"].Value = "TERMINAL USED";
					ws.Cells["P1"].Value = "BILLERNAME";
					ws.Cells["Q1"].Value = "MERCHANT";
					ws.Cells["R1"].Value = "INTER";
					ws.Cells["S1"].Value = "BANCNET USED";
					ws.Cells["T1"].Value = "ONLINE";
					ws.Cells["U1"].Value = "WEB SITE";
					ws.Cells["V1"].Value = "REMIT FROM";
					ws.Cells["W1"].Value = "REMIT CONCERN";
					ws.Cells["X1"].Value = "AMOUNT";
					ws.Cells["Y1"].Value = "CURRENCY";
					ws.Cells["Z1"].Value = "DATE TRANS";
					ws.Cells["AA1"].Value = "PL STATUS";
					ws.Cells["AB1"].Value = "CARD BLOCKED";
					ws.Cells["AC1"].Value = "ACTIVITY";
					ws.Cells["AD1"].Value = "STATUS";
					ws.Cells["AE1"].Value = "SUB STATUS";
					ws.Cells["AF1"].Value = "CREATED BY";
					ws.Cells["AG1"].Value = "LAST UPDATED BY";
					ws.Cells["AH1"].Value = "ENDORSED TO";
					ws.Cells["AI1"].Value = "ENDORSED FROM";
					ws.Cells["AJ1"].Value = "DUE DATE";
					ws.Cells["AK1"].Value = "RESOLVED DATE";
					ws.Cells["AL1"].Value = "NATURE OF COMPLAINT";
					ws.Cells["AM1"].Value = "REFERRED BY";
					ws.Cells["AN1"].Value = "CONTACT PERSON";
					ws.Cells["AO1"].Value = "CLASSIFICATION";
					ws.Cells["AP1"].Value = "TIME OF CALL";
					ws.Cells["AQ1"].Value = "LOCAL NUMBER";
					ws.Cells["AR1"].Value = "CARD PRESENT";
					ws.Cells["AS1"].Value = "BRANCH";

				int rowStart = 2;
					foreach (var item in v)
					{
						ws.Cells[string.Format("A{0}", rowStart)].Value = item.TicketNo;
						ws.Cells[string.Format("B{0}", rowStart)].Value = item.InOutBound;
						ws.Cells[string.Format("C{0}", rowStart)].Value = item.CommChannel;
						ws.Cells[string.Format("D{0}", rowStart)].Value = item.DateReceived.ToString("dd-MMM-yyyy");
						ws.Cells[string.Format("E{0}", rowStart)].Value = item.DateReceived.ToString("HH:mm:ss");
						ws.Cells[string.Format("F{0}", rowStart)].Value = item.abbreviation;
						ws.Cells[string.Format("G{0}", rowStart)].Value = item.CustomerName;
						ws.Cells[string.Format("H{0}", rowStart)].Value = item.CardNo;
						ws.Cells[string.Format("I{0}", rowStart)].Value = item.Destination;
						ws.Cells[string.Format("J{0}", rowStart)].Value = item.TransType;
						ws.Cells[string.Format("K{0}", rowStart)].Value = item.ATMTrans;
						ws.Cells[string.Format("L{0}", rowStart)].Value = item.Location;
						ws.Cells[string.Format("M{0}", rowStart)].Value = item.ATMUsed;
						ws.Cells[string.Format("N{0}", rowStart)].Value = item.PaymentOf;
						ws.Cells[string.Format("O{0}", rowStart)].Value = item.TerminalUsed;
						ws.Cells[string.Format("P{0}", rowStart)].Value = item.BillerName;
						ws.Cells[string.Format("Q{0}", rowStart)].Value = item.Merchant;
						ws.Cells[string.Format("R{0}", rowStart)].Value = item.Inter;
						ws.Cells[string.Format("S{0}", rowStart)].Value = item.BancnetUsed;
						ws.Cells[string.Format("T{0}", rowStart)].Value = item.Online;
						ws.Cells[string.Format("U{0}", rowStart)].Value = item.Website;
						ws.Cells[string.Format("V{0}", rowStart)].Value = item.RemitFrom;
						ws.Cells[string.Format("W{0}", rowStart)].Value = item.RemitConcern;
						ws.Cells[string.Format("X{0}", rowStart)].Value = item.Amount;
						ws.Cells[string.Format("Y{0}", rowStart)].Value = item.Currency;
						ws.Cells[string.Format("Z{0}", rowStart)].Value = (item.DateTrans == null) ? "" : Convert.ToDateTime(item.DateTrans).ToString("dd-MMM-yyyy"); //item.DateTrans.ToString("dd-MMM-yyyy");
						ws.Cells[string.Format("AA{0}", rowStart)].Value = item.PLStatus;
						ws.Cells[string.Format("AB{0}", rowStart)].Value = item.CardBlocked;
						ws.Cells[string.Format("AC{0}", rowStart)].Value = item.Activity;
						ws.Cells[string.Format("AD{0}", rowStart)].Value = item.Status;
						ws.Cells[string.Format("AE{0}", rowStart)].Value = item.SubStatus;
						ws.Cells[string.Format("AF{0}", rowStart)].Value = item.created_by;
						ws.Cells[string.Format("AG{0}", rowStart)].Value = item.Last_Update;
						ws.Cells[string.Format("AH{0}", rowStart)].Value = item.EndoredTo;
						ws.Cells[string.Format("AI{0}", rowStart)].Value = item.EndorsedFrom;
						ws.Cells[string.Format("AJ{0}", rowStart)].Value = (item.Resolved == null) ? "" : Convert.ToDateTime(item.Resolved).ToString("dd-MMM-yyyy");
						ws.Cells[string.Format("AK{0}", rowStart)].Value = (item.ResolvedDate == null) ? "" : Convert.ToDateTime(item.ResolvedDate).ToString("dd-MMM-yyyy"); // item.ResolvedDate;
						ws.Cells[string.Format("AL{0}", rowStart)].Value = item.Remarks;
						ws.Cells[string.Format("AM{0}", rowStart)].Value = item.ReferedBy;
						ws.Cells[string.Format("AN{0}", rowStart)].Value = item.ContactPerson;
						ws.Cells[string.Format("AO{0}", rowStart)].Value = item.Tagging;
						ws.Cells[string.Format("AP{0}", rowStart)].Value = item.CallTime;
						ws.Cells[string.Format("AQ{0}", rowStart)].Value = item.LocalNum;
						ws.Cells[string.Format("AR{0}", rowStart)].Value = item.CardPresent;
						ws.Cells[string.Format("AS{0}", rowStart)].Value = item.Branch;

					rowStart++;
					}

					ws.Cells["A:AS"].AutoFitColumns();
					ws.Cells["X:X"].Style.Numberformat.Format = "#,##0.00";

					package.Save();
				}

				if (id == 1)
				{
					fnam = "CM Activity with History " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
				}
				else
				{
					fnam = "CM Activity " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
				}

				string fileName = fnam;
				string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				stream.Position = 0;
				return File(stream, fileType, fileName);


			//}
			//catch (Exception ex)
			//{
			//	string msg = ex.ToString();

			//	TempData["msgErr"] = msg;

			//	throw;
			//}


			
		}

		[HttpGet]
		public IActionResult ATMLogs()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Activity", "ATMLogs");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Activity";
				ViewBag.Page = "ATMLogs";

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Customer Service - Activity", ses.WrkStn, "View List of ATM Logs");

				var vw = _db.ListATMLogs();

				ViewBag.ListATMLog = vw;

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}

		}


		[HttpGet]
		public IActionResult NewCMEntry(string name)
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Activity", "NewCMEntry");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}


				if (name != null)
				{

					string[] v = name.Split(new char[] { ',' });

					string cardno = v[0];
					string lastname = v[1];
					string firstname = v[2];
					string midname = v[3].Replace(";", "");
					string custname = (midname == "" || midname == null) ? firstname + " " + lastname : firstname + " " + midname + " " + lastname;

					HttpContext.Session.SetString("CardNo", cardno);
					HttpContext.Session.SetString("LastName", lastname);
					HttpContext.Session.SetString("FirstName", firstname);
					HttpContext.Session.SetString("MidName", midname);
					HttpContext.Session.SetString("CustName", custname);
				}

				HttpContext.Session.SetString("createdby", ses.LoginID);

				ViewBag.Interface = "Activity";
				ViewBag.Page = "CustomerService";

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbChannel = new SelectList(_db.Get_Dropdown("CL"), "ID", "Description");
				ViewBag.cmbTransType = new SelectList(_db.Get_Dropdown("AT"), "ID", "Description");
				ViewBag.cmbPayInst = new SelectList(_db.Get_Dropdown("IT"), "ID", "Description");
				ViewBag.cmbTermUsed = new SelectList(_db.Get_Dropdown("TU"), "ID", "Description");
				ViewBag.cmbBancUsed = new SelectList(_db.Get_Dropdown("BN"), "ID", "Description");
				ViewBag.cmbATMUsed = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");
				ViewBag.cmbLoanStat = new SelectList(_db.Get_Dropdown("PS"), "ID", "Description");
				ViewBag.cmbEndorsedTo = new SelectList(_db.Get_Dropdown("ET"), "ID", "Description");
				ViewBag.cmbEndrosedFrom = new SelectList(_db.Get_Dropdown("EF"), "ID", "Description");
				ViewBag.cmbCommType = new SelectList(_db.Get_Dropdown("CT"), "ID", "Description");
				ViewBag.cmbCommChannel = new SelectList(_db.Get_Dropdown("CC"), "ID", "Description");
				ViewBag.cmbStatus = new SelectList(_db.Get_Dropdown("ST"), "ID", "Description");
				ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("CS"), "ID", "Description");
				ViewBag.cmbRemarks = new SelectList(_db.Get_Dropdown("NT"), "ID", "Description");
				ViewBag.cmbReferedBy = new SelectList(_db.Get_Dropdown("RB"), "ID", "Description");
				ViewBag.cmbCurrency = new SelectList(_db.Get_Dropdown("CR"), "ID", "Description");
				ViewBag.cmbBranch = new SelectList(_db.Get_Dropdown3(), "ID", "Description");
				ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown4(), "ID", "Description");

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}

			
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public async Task<IActionResult> NewCMEntry(NewCallReport cm, List<IFormFile> files)
		{
			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			try
			{

				if (string.IsNullOrEmpty(cm.LastName) && string.IsNullOrEmpty(cm.FirstName))
				{
					ModelState.AddModelError(nameof(cm.CustomerName), "Customer Name required");

					TempData["msgErr"] = "Mandatory fields required!";
				}

				if (string.IsNullOrEmpty(cm.Branch))
				{
					ModelState.AddModelError(nameof(cm.Branch), "Branch required");

					TempData["msgErr"] = "Mandatory fields required!";
				}

				if (string.IsNullOrEmpty(cm.CommChannel))
				{
					ModelState.AddModelError(nameof(cm.CommChannel), "Comm Channel required");

					TempData["msgErr"] = "Mandatory fields required!";
				}

				if (string.IsNullOrEmpty(cm.InOutBound))
				{
					ModelState.AddModelError(nameof(cm.InOutBound), "Comm Type required");

					TempData["msgErr"] = "Mandatory fields required!";
				}

				if (string.IsNullOrEmpty(cm.ReasonCode))
				{
					ModelState.AddModelError(nameof(cm.ReasonCode), "Reason Code required");

					TempData["msgErr"] = "Mandatory fields required!";
				}


				//else
				//{
				//	switch (cm.ReasonCode)
				//	{
				//		case "SPR":
				//		case "NCC":
				//		case "IPC":
				//		case "PNC":

				//			if (string.IsNullOrEmpty(cm.Remarks))
				//			{
				//				ModelState.AddModelError(nameof(cm.Remarks), "Nature of Complaint required");

				//				TempData["msgErr"] = "Mandatory fields required!";
				//			}

				//			break;
				//	}
				//}


				if (string.IsNullOrEmpty(cm.Status))
				{
					ModelState.AddModelError(nameof(cm.Status), "Status required");

					TempData["msgErr"] = "Mandatory fields required!";
				}

				if (!string.IsNullOrEmpty(cm.Amount))
				{
					double amt = Convert.ToDouble(cm.Amount);

					if (amt != 0)
					{
						if (string.IsNullOrEmpty(cm.Currency))
						{
							ModelState.AddModelError(nameof(cm.Currency), "Currency required");
							TempData["msgErr"] = "Mandatory fields required!";

						}
					}

				}

				if (cm.Status == "Endorsed")
				{

					if (string.IsNullOrEmpty(cm.EndoredTo))
					{
						ModelState.AddModelError(nameof(cm.EndoredTo), "Endorsed To required");

						TempData["msgErr"] = "Mandatory fields required!";
					}


					if (string.IsNullOrEmpty(cm.Remarks))
					{
						ModelState.AddModelError(nameof(cm.Remarks), "Nature of Complaint required");

						TempData["msgErr"] = "Mandatory fields required!";

					}

				}

				if (!string.IsNullOrEmpty(cm.EndoredTo))
				{
					switch (cm.EndoredTo)
					{
						case "Others":
						case "ATM Ops":

							if (string.IsNullOrEmpty(cm.CardNo))
							{
								ModelState.AddModelError(nameof(cm.CardNo), "Card Number required");

								TempData["msgErr"] = "Mandatory fields required!";
							}

							break;
					}

					if (string.IsNullOrEmpty(cm.ReferedBy))
					{
						ModelState.AddModelError(nameof(cm.ReferedBy), "Reffered By required");
						TempData["msgErr"] = "Mandatory fields required!";
					}

					if (string.IsNullOrEmpty(cm.ContactPerson))
					{
						ModelState.AddModelError(nameof(cm.ContactPerson), "Contact Person required");
						TempData["msgErr"] = "Mandatory fields required!";
					}

				}

				if (string.IsNullOrEmpty(cm.Activity))
				{
					ModelState.AddModelError(nameof(cm.Activity), "Activity required");

					TempData["msgErr"] = "Mandatory fields required!";
				}

				//GET THE TICKET NO.
				var tickets = _db.DisplayTickets(cm.ReasonCode);

				string TicketNo = cm.ReasonCode + '-' + string.Format("{0:yyMM}", DateTime.Today).ToString() + '-' + tickets.ticket.ToString();

				string cust = (string.IsNullOrEmpty(cm.MiddleName)) ? cm.FirstName + ' ' + cm.LastName : cm.FirstName + ' ' + cm.MiddleName + ' ' + cm.LastName;

				string loc = (cm.ATMUsed == "325") ? cm.Location : cm.Location2;

				//cm.CustomerName = cust;

				HttpContext.Session.SetString("CardNo", (string.IsNullOrEmpty(cm.CardNo)) ? string.Empty : cm.CardNo);
				HttpContext.Session.SetString("LastName", (string.IsNullOrEmpty(cm.LastName)) ? string.Empty : cm.LastName);
				HttpContext.Session.SetString("FirstName", (string.IsNullOrEmpty(cm.FirstName)) ? string.Empty : cm.FirstName);
				HttpContext.Session.SetString("MidName", (string.IsNullOrEmpty(cm.MiddleName)) ? string.Empty : cm.MiddleName);
				HttpContext.Session.SetString("CustName", (string.IsNullOrEmpty(cm.CustomerName)) ? string.Empty : cm.CustomerName);
				HttpContext.Session.SetString("createdby", ses.LoginID);

				if (ModelState.IsValid)
				{
					bool isCreated = _db.CreateNewActivity(
						cm.CommChannel,
						cm.ReasonCode,
						cust,
						cm.CardNo,
						cm.Activity,
						cm.Status,
						cm.Created_By,
						TicketNo,
						cm.LastName,
						cm.FirstName,
						cm.MiddleName,
						cm.InOutBound,
						cm.TransType,
						loc,
						cm.ATMUsed,
						cm.PaymentOf,
						cm.TerminalUsed,
						cm.Merchant,
						cm.Inter,
						cm.BancnetUsed,
						cm.Online,
						cm.Website,
						Convert.ToDouble(cm.Amount),
						cm.DateTrans,
						cm.PLStatus,
						cm.CardBlocked,
						cm.EndoredTo,
						cm.Destination,
						cm.ATMTrans,
						cm.RemitFrom,
						cm.RemitConcern,
						cm.EndorsedFrom,
						cm.Resolved,
						cm.BillerName,
						cm.Remarks,
						cm.ReferedBy,
						cm.ContactPerson,
						cm.Tagging,
						cm.CallTime,
						cm.LocalNum,
						cm.Created_By,
						cm.Last_Update,
						cm.CardPresent,
						cm.SubStatus,
						cm.Currency,
						cm.Branch
						);


					if (isCreated)
					{

						long size = files.Sum(f => f.Length);

						var filePaths = new List<string>();

						foreach (var formFile in files)
						{
							if (formFile.Length > 0)
							{
								var filePath = Path.Combine(_webHostEnvironment.ContentRootPath, @"wwwroot/files", formFile.FileName);

								if (System.IO.File.Exists(filePath))
								{
									System.IO.File.Delete(filePath);
								}

								filePaths.Add(filePath);

								using (var stream = new FileStream(filePath, FileMode.Create))
								{
									bool insertFile = _db.UploadInputFile(TicketNo, formFile.FileName);

									await formFile.CopyToAsync(stream);
								}

							}
						}

						//SendEmail if Endorsed From enabled
						bool chkRCode = _db.ChkIremitCode(cm.ReasonCode);

						if (chkRCode)
						{
							rcd = "Y";
						}
						else
						{
							rcd = "N";
						}

						//var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

						switch (cm.EndoredTo)
						{
							case "Others":
								bool SendEmail = _db.SendEndorsedMail("0", rcd, cm.EndorsedFrom, TicketNo, cm.Remarks, cm.LastName, cm.FirstName, cm.MiddleName, cm.CardNo, cm.Activity, ses.FullName, cm.ReferedBy, cm.ContactPerson, filePaths);

								break;
							case "ATM Ops":

								bool isATMCreated = _db.CreateNewATMActivity(
									cm.CommChannel,
									cm.ReasonCode,
									cust,
									cm.CardNo,
									cm.Activity,
									cm.Status,
									cm.Created_By,
									TicketNo,
									cm.LastName,
									cm.FirstName,
									cm.MiddleName,
									cm.InOutBound,
									cm.TransType,
									loc,
									cm.ATMUsed,
									cm.PaymentOf,
									cm.TerminalUsed,
									cm.Merchant,
									cm.Inter,
									cm.BancnetUsed,
									cm.Online,
									cm.Website,
									Convert.ToDouble(cm.Amount),
									cm.DateTrans,
									cm.PLStatus,
									cm.CardBlocked,
									cm.EndoredTo,
									cm.Destination,
									cm.ATMTrans,
									cm.RemitFrom,
									cm.RemitConcern,
									cm.EndorsedFrom,
									cm.Resolved,
									cm.BillerName,
									cm.Remarks,
									cm.ReferedBy,
									cm.ContactPerson,
									cm.Tagging,
									cm.CallTime,
									cm.LocalNum,
									cm.Created_By,
									cm.Last_Update,
									cm.CardPresent,
									cm.SubStatus,
									cm.Currency,
									cm.Branch
									);

								//bool SendEmailATM = _db.SendEndorsedMail("0", rcd, "atm.ops@sterlingbankasia.com", TicketNo, cm.Remarks, cm.LastName, cm.FirstName, cm.MiddleName, cm.CardNo, cm.Activity, ses.FullName, cm.ReferedBy, cm.ContactPerson, filePaths);

								//string cardno = cm.CardNo.Substring(0, 6) + "******" + cm.CardNo.Substring(cm.CardNo.Length - 4, 4);

								bool SendEmailATM = _db.SendEndorsedMail("0", rcd, "atm.ops@sterlingbankasia.com", TicketNo, cm.Remarks, cm.LastName, cm.FirstName, cm.MiddleName, cm.CardNo, cm.Activity, ses.FullName, cm.ReferedBy, cm.ContactPerson, filePaths);
								break;
						}

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("Customer Service - Activity", ses.WrkStn, "Add New Activity Ticket No. " + TicketNo);

						TempData["msgOk"] = "Ticket Number " + TicketNo + " has been successfully saved.";

						HttpContext.Session.SetString("CardNo", string.Empty);
						HttpContext.Session.SetString("LastName", string.Empty);
						HttpContext.Session.SetString("FirstName", string.Empty);
						HttpContext.Session.SetString("MidName", string.Empty);
						HttpContext.Session.SetString("CustName", string.Empty);
						HttpContext.Session.SetString("createdby", string.Empty);

						return RedirectToAction("CustomerService");
					}
					else
					{
						TempData["msgErr"] = "Creation of new ticket was not successfull";

						return View(cm);
					}



				}


			}
			catch (Exception ex)
			{
				string controllerName = ControllerContext.ActionDescriptor.ControllerName;
				string actionName = ControllerContext.ActionDescriptor.ActionName;

				string err = "Error in " + controllerName + " - " + actionName + " : " + ex.ToString();

				TempData["msgErr"] = err;
			}

			ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
			ViewBag.cmbChannel = new SelectList(_db.Get_Dropdown("CL"), "ID", "Description");
			ViewBag.cmbTransType = new SelectList(_db.Get_Dropdown("AT"), "ID", "Description");
			ViewBag.cmbPayInst = new SelectList(_db.Get_Dropdown("IT"), "ID", "Description");
			ViewBag.cmbTermUsed = new SelectList(_db.Get_Dropdown("TU"), "ID", "Description");
			ViewBag.cmbBancUsed = new SelectList(_db.Get_Dropdown("BN"), "ID", "Description");
			ViewBag.cmbATMUsed = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");
			ViewBag.cmbLoanStat = new SelectList(_db.Get_Dropdown("PS"), "ID", "Description");
			ViewBag.cmbEndorsedTo = new SelectList(_db.Get_Dropdown("ET"), "ID", "Description");
			ViewBag.cmbEndrosedFrom = new SelectList(_db.Get_Dropdown("EF"), "ID", "Description");
			ViewBag.cmbCommType = new SelectList(_db.Get_Dropdown("CT"), "ID", "Description");
			ViewBag.cmbCommChannel = new SelectList(_db.Get_Dropdown("CC"), "ID", "Description");
			ViewBag.cmbStatus = new SelectList(_db.Get_Dropdown("ST"), "ID", "Description");
			ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("CS"), "ID", "Description");
			ViewBag.cmbRemarks = new SelectList(_db.Get_Dropdown("NT"), "ID", "Description");
			ViewBag.cmbReferedBy = new SelectList(_db.Get_Dropdown("RB"), "ID", "Description");
			ViewBag.cmbCurrency = new SelectList(_db.Get_Dropdown("CR"), "ID", "Description");
			ViewBag.cmbBranch = new SelectList(_db.Get_Dropdown3(), "ID", "Description");
			//ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown_Location(Convert.ToInt32(cm.ATMUsed)), "Value", "Text");
			ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown4(), "ID", "Description");

			return View(cm);

		}

		[HttpGet]
		public IActionResult CMEntry(int id)
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				string _id = id.ToString();

				if (Security.Encryptor(_id).ToString() != "iuwOJsfeFeMlK/Wa2Ea3mQ==")
				{
					TempData["msginfo"] = "You are not authorized to use this feature.";
					return RedirectToAction("Index", "Home");
				}

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
			
		}

		[HttpPost]
		public IActionResult CMEntry(string query)
		{
			
			if (!string.IsNullOrWhiteSpace(query))
			{
				try
				{
					var table = new StringBuilder();

					SqlConnectionStringBuilder build = new SqlConnectionStringBuilder("Server = {ServerName},1433; Initial Catalog = {Database}; Persist Security Info = False; User ID = {Username}; Password = {Password}; MultipleActiveResultSets = False; Encrypt = True; TrustServerCertificate = False; Authentication = Active Directory Password");


					using (con)
					{
						con.Open();

						SqlCommand cmd = (SqlCommand)con.CreateCommand();

						cmd.CommandTimeout = 0;

						cmd.CommandText = query;

						SqlDataReader reader = cmd.ExecuteReader();

						if (reader.Read())
						{
							table.AppendLine("<div style=\"overflow:auto;\">");
							table.AppendLine("<table class=\"table-sm table-bordered table-hover table-vcenter js-cmentry\">");
							table.AppendLine("<thead class=\"bg-primary-light\"><tr>");
							for (int i = 0; i < reader.FieldCount; i++)
							{
								table.AppendLine("<th class=\"d-none d-sm-table-cell font-size-xs\">");
								table.AppendLine(reader.GetName(i));
								table.AppendLine("</th>");
							}
							table.AppendLine("</tr></thead>");
							table.AppendLine("<tbody>");
							while (reader.Read())
							{
								table.AppendLine("<tr>");

								for (int i = 0; i < reader.FieldCount; i++)
								{
									table.AppendLine("<td class=\"d-none d-sm-table-cell font-size-xs\" style=\"white-space:nowrap\" >");
									table.AppendLine(reader[i].ToString());
									table.AppendLine("</td>");
								}
								table.AppendLine("</tr>");
							}
							table.AppendLine("</tbody>");
							table.AppendLine("</table>");
							table.AppendLine("</div>");
							ViewBag.Result = table.ToString();
						}
						else
						{
							ViewBag.Result = string.Format("{0} records affected", reader.RecordsAffected);
						}

						reader.Close();
					}
				}
				catch (Exception ex)
				{
					ViewBag.Result = ex.Message;
				}
			}

			return View();

		}


		public IActionResult SearchICBA(string id)
		{
			//string id = HttpContext.Session.GetString("ClientSearch");

			if (id == "")
			{
				TempData["msginfo"] = "Please enter client name to search";

			}

			var v = _db.Search_ICBA(id);

			ViewBag.ListICBAClients = v;

			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "Customer does not exist. Please enter client first name or last name only";
			}

			//ViewBag.Wala = id;

			return PartialView("_SearchICBA");
		}

		public IActionResult SearchBW(string id)
		{
			//string id = HttpContext.Session.GetString("ClientSearch");

			if (id == "")
			{
				TempData["msginfo"] = "Please enter client name to search";

				return RedirectToAction("CustomerService");
			}

			ViewBag.ListBWClients = _db.Search_BW(id);

			//ViewBag.Wala = id;

			return PartialView("_SearchBW");
		}

		public JsonResult GetCustomer(string myInput)
		{


			HttpContext.Session.SetString("ClientSearch", myInput ?? "");

			return Json(new { success = true });
		}

		[HttpGet]
		public IActionResult LoadICBA()
		{
			try
			{
				var draw = Request.Form["draw"].FirstOrDefault();
				// Skiping number of Rows count  
				var start = Request.Form["start"].FirstOrDefault();
				// Paging Length 10,20  
				var length = Request.Form["length"].FirstOrDefault();
				// Sort Column Name  
				var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][name]"].FirstOrDefault();
				// Sort Column Direction ( asc ,desc)  
				var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
				// Search Value from (Search box)  
				var searchValue = Request.Form["search[value]"].FirstOrDefault();

				//Paging Size (10,20,50,100)  
				int pageSize = length != null ? Convert.ToInt32(length) : 0;
				int skip = start != null ? Convert.ToInt32(start) : 0;
				int recordsTotal = 0;

				// Getting all Customer data  
				var customerData = _db.Search_ICBA(HttpContext.Session.GetString("ClientSearch")); 

				//Sorting  
				//if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDirection)))
				//{
				//	customerData = customerData.OrderBy(sortColumn + " " + sortColumnDirection);
				//}

				//Search  
				if (!string.IsNullOrEmpty(searchValue))
				{
					customerData = customerData.Where(p => p.ACCTNO.Contains(searchValue) ||
								 p.LNAME.ToLower().Contains(searchValue) ||
								 p.DESCR.ToLower().Contains(searchValue) ||
								 p.CRLINE.ToString().ToLower().Contains(searchValue) ||
								 p.ACSTS.ToLower().Contains(searchValue)
					);
				}

				//total number of rows count   
				recordsTotal = customerData.Count();
				//Paging   
				var data = customerData.Skip(skip).Take(pageSize).ToList();
				//Returning Json Data  
				return Json(new { draw = draw, recordsFiltered = recordsTotal, recordsTotal = recordsTotal, data = data });

			}
			catch (Exception)
			{
				throw;
			}

		}

		[HttpGet]
		public JsonResult GetATMLocation(int ATMUsed)
		{

			//ViewBag.ListLocation = _db.Get_Dropdown_Location(ATMUsed);

			var location = _db.Get_Dropdown_Location(ATMUsed);

			return Json(new SelectList(location, "Value", "Text"));
		}

		public IActionResult ExportATMLogs(int id)
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");

			string fnam = string.Empty;

			var v = _db.ListATMLogs();

			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("CustomerService", "Activity");
			}


			using (ExcelPackage package = new(stream))
			{

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.TicketNo.ToLower().Contains(myInput.ToLower()) ||
									p.InOutBound.ToLower().Contains(myInput.ToLower()) ||
									p.CommChannel.ToLower().Contains(myInput.ToLower()) ||
									p.DateReceived.ToLower().Contains(myInput.ToLower()) ||
									p.ReasonCode.ToLower().Contains(myInput.ToLower()) ||
									p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
									p.CardNo.ToLower().Contains(myInput.ToLower()) ||
									p.Destination.ToLower().Contains(myInput.ToLower()) ||
									p.TransType.ToLower().Contains(myInput.ToLower()) ||
									p.ATMTrans.ToLower().Contains(myInput.ToLower()) ||
									p.Location.ToLower().Contains(myInput.ToLower()) ||
									p.ATMUsed.ToLower().Contains(myInput.ToLower()) ||
									p.PaymentOf.ToLower().Contains(myInput.ToLower()) ||
									p.TerminalUsed.ToLower().Contains(myInput.ToLower()) ||
									p.BillerName.ToLower().Contains(myInput.ToLower()) ||
									p.Merchant.ToLower().Contains(myInput.ToLower()) ||
									p.Inter.ToLower().Contains(myInput.ToLower()) ||
									p.BancnetUsed.ToLower().Contains(myInput.ToLower()) ||
									p.Online.ToLower().Contains(myInput.ToLower()) ||
									p.Website.ToLower().Contains(myInput.ToLower()) ||
									p.RemitFrom.ToLower().Contains(myInput.ToLower()) ||
									p.RemitConcern.ToLower().Contains(myInput.ToLower()) ||
									p.Amount.ToLower().Contains(myInput.ToLower()) ||
									p.DateTrans.ToLower().Contains(myInput.ToLower()) ||
									p.PLStatus.ToLower().Contains(myInput.ToLower()) ||
									p.CardBlocked.ToLower().Contains(myInput.ToLower()) ||
									p.Activity.ToLower().Contains(myInput.ToLower()) ||
									p.Status.ToLower().Contains(myInput.ToLower()) ||
									p.Agent.ToLower().Contains(myInput.ToLower()) ||
									p.EndoredTo.ToLower().Contains(myInput.ToLower()) ||
									p.Resolved.ToLower().Contains(myInput.ToLower())
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "TICKET NO";
				ws.Cells["B1"].Value = "INBOUND/OUTBOUND";
				ws.Cells["C1"].Value = "COMM CHANNEL";
				ws.Cells["D1"].Value = "DATE RECEIVED";
				ws.Cells["E1"].Value = "REASON CODE";
				ws.Cells["F1"].Value = "CUSTOMER NAME";
				ws.Cells["G1"].Value = "CARD NO";
				ws.Cells["H1"].Value = "DESTINATION";
				ws.Cells["I1"].Value = "TRANS TYPE";
				ws.Cells["J1"].Value = "ATM TRANS";
				ws.Cells["K1"].Value = "LOCATION";
				ws.Cells["L1"].Value = "ATM USED";
				ws.Cells["M1"].Value = "PAYMENT OF";
				ws.Cells["N1"].Value = "TERMINAL USED";
				ws.Cells["O1"].Value = "BILLERNAME";
				ws.Cells["P1"].Value = "MERCHANT";
				ws.Cells["Q1"].Value = "INTER";
				ws.Cells["R1"].Value = "BANCNET USED";
				ws.Cells["S1"].Value = "ONLINE";
				ws.Cells["T1"].Value = "WEB SITE";
				ws.Cells["U1"].Value = "REMIT FROM";
				ws.Cells["V1"].Value = "REMIT CONCERN";
				ws.Cells["W1"].Value = "AMOUNT";
				ws.Cells["X1"].Value = "DATE TRANS";
				ws.Cells["Y1"].Value = "PL STATUS";
				ws.Cells["Z1"].Value = "CARD BLOCKED";
				ws.Cells["AA1"].Value = "ACTIVITY";
				ws.Cells["AB1"].Value = "STATUS";
				ws.Cells["AC1"].Value = "AGENT";
				ws.Cells["AD1"].Value = "ENDORSED TO";
				ws.Cells["AE1"].Value = "DUE DATE";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.TicketNo;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.InOutBound;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.CommChannel;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.DateReceived; 
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.ReasonCode;
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.CustomerName;
					ws.Cells[string.Format("G{0}", rowStart)].Value = item.CardNo;
					ws.Cells[string.Format("H{0}", rowStart)].Value = item.Destination;
					ws.Cells[string.Format("I{0}", rowStart)].Value = item.TransType;
					ws.Cells[string.Format("J{0}", rowStart)].Value = item.ATMTrans;
					ws.Cells[string.Format("K{0}", rowStart)].Value = item.Location;
					ws.Cells[string.Format("L{0}", rowStart)].Value = item.ATMUsed;
					ws.Cells[string.Format("M{0}", rowStart)].Value = item.PaymentOf;
					ws.Cells[string.Format("N{0}", rowStart)].Value = item.TerminalUsed;
					ws.Cells[string.Format("O{0}", rowStart)].Value = item.BillerName;
					ws.Cells[string.Format("P{0}", rowStart)].Value = item.Merchant;
					ws.Cells[string.Format("Q{0}", rowStart)].Value = item.Inter;
					ws.Cells[string.Format("R{0}", rowStart)].Value = item.BancnetUsed;
					ws.Cells[string.Format("S{0}", rowStart)].Value = item.Online;
					ws.Cells[string.Format("T{0}", rowStart)].Value = item.Website;
					ws.Cells[string.Format("U{0}", rowStart)].Value = item.RemitFrom;
					ws.Cells[string.Format("V{0}", rowStart)].Value = item.RemitConcern;
					ws.Cells[string.Format("W{0}", rowStart)].Value = item.Amount;
					ws.Cells[string.Format("X{0}", rowStart)].Value = item.DateTrans;
					ws.Cells[string.Format("Y{0}", rowStart)].Value = item.PLStatus;
					ws.Cells[string.Format("Z{0}", rowStart)].Value = item.CardBlocked;
					ws.Cells[string.Format("AA{0}", rowStart)].Value = item.Activity;
					ws.Cells[string.Format("AB{0}", rowStart)].Value = item.Status;
					ws.Cells[string.Format("AC{0}", rowStart)].Value = item.Agent;
					ws.Cells[string.Format("AD{0}", rowStart)].Value = item.EndoredTo;
					ws.Cells[string.Format("AE{0}", rowStart)].Value = item.Resolved;

					rowStart++;
				}

				ws.Cells["A:AE"].AutoFitColumns();
				ws.Cells["W:W"].Style.Numberformat.Format = "#,##0.00";

				package.Save();
			}

			fnam = "ATM Logs " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);

		}


		[HttpGet]
		public IActionResult ATMOps()
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Activity", "ATMOps");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "ATM";
				ViewBag.Page = "ATMOps";

				//HttpContext.Session.SetString("deyt1", DateTime.Now.ToString("dd-MMM-yyyy"));
				//HttpContext.Session.SetString("deyt2", DateTime.Now.ToString("dd-MMM-yyyy"));

				//var vw = _db.ListATMOps(DateTime.Now.ToString("dd-MMM-yyyy"), DateTime.Now.ToString("dd-MMM-yyyy"));

				//ViewBag.ListATMOps = vw;

				ActivitySearch s = new ActivitySearch();

				var vw = _db.View_ATMActivity(s);

				ViewBag.ListATMOps = vw;

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("ATM Operations - Activity", ses.WrkStn, "View List of ATM Ops Activities");

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");

				var ser = new ActivitySearch()
				{
					DateSearch = s.DateSearch,
					Date1 = s.Date1,
					Date2 = s.Date2,
					CustomerName = s.CustomerName,
					CardNo = s.CardNo,
					CommType = s.CommType,
					CommChannel = s.CommChannel,
					ReasonCode = s.ReasonCode,
					Status = s.Status,
					Agent = s.Agent,
					TicketNo = s.TicketNo,
					ReferredBy = s.ReferredBy,
					ContactPerson = s.ContactPerson
				};

				HttpContext.Session.SetString("SearchActivity", JsonConvert.SerializeObject(ser));

				return View(s);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}

		}

		[HttpPost]
		public IActionResult ATMOps(ActivitySearch s)
		{
			try
			{
				ViewBag.Interface = "Activity";
				ViewBag.Page = "ATM Ops";

				var vw = _db.View_ATMActivity(s);

				ViewBag.ListATMOps = vw;

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");


				var ser = new ActivitySearch()
				{
					DateSearch = s.DateSearch,
					Date1 = s.Date1,
					Date2 = s.Date2,
					CustomerName = s.CustomerName,
					CardNo = s.CardNo,
					CommType = s.CommType,
					CommChannel = s.CommChannel,
					ReasonCode = s.ReasonCode,
					Status = s.Status,
					Agent = s.Agent,
					TicketNo = s.TicketNo,
					ReferredBy = s.ReferredBy,
					ContactPerson = s.ContactPerson
				};

				HttpContext.Session.SetString("SearchActivity", JsonConvert.SerializeObject(ser));

			}
			catch (Exception ex)
			{
				string controllerName = ControllerContext.ActionDescriptor.ControllerName;
				string actionName = ControllerContext.ActionDescriptor.ActionName;

				string err = "Error in " + controllerName + " - " + actionName + " : " + ex.ToString();

				TempData["msgErr"] = err;
			}

			return View(s);

		}

		public IActionResult SearchCard(string id)
		{

			if (id == "")
			{
				TempData["msginfo"] = "Please enter card number to search";

				return RedirectToAction("NewCMEntry");
			}

			ViewBag.ListClientsCard = _db.Search_Card(id);


			return PartialView("_SearchCard");
		}

		public IActionResult SearchCardNo(string id)
		{

			if (id == "")
			{
				TempData["msginfo"] = "Please enter card number to search";

				return RedirectToAction("NewCMEntry");
			}

			ViewBag.ListClientsCard = _db.Search_Card(id);


			return PartialView("_SearchCardNo");
		}

		[HttpGet]
		public JsonResult GetSubStatus(string Status)
		{
			var substatus = new SelectList(_db.Get_Dropdown("SS"), "ID", "Description");

			return Json(new SelectList(substatus, "Value", "Text"));
		}


		public IActionResult ActivityOption(int id)
		{

			EditCallReport tk = _db.ViewCMEntry(id);

			ViewBag.TicketNo = tk.TicketNo;
			ViewBag.RID = id.ToString();

			return PartialView("_ActivityOption");
		}

		public IActionResult ActivityATMOption(int id)
		{

			EditATMReport tk = _db.ViewATMEntry(id);

			ViewBag.TicketNo = tk.TicketNo;
			ViewBag.RID = id.ToString();

			return PartialView("_ActivityATMOption");
		}


		[HttpGet]
		public IActionResult UpdateCMEntry(int id)
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Activity", "UpdateCMEntry");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Activity";
				ViewBag.Page = "CustomerService";

				EditCallReport edt = _db.ViewCMEntry(id);

				HttpContext.Session.SetString("updatedby", ses.LoginID);

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbChannel = new SelectList(_db.Get_Dropdown("CL"), "ID", "Description");
				ViewBag.cmbTransType = new SelectList(_db.Get_Dropdown("AT"), "ID", "Description");
				ViewBag.cmbPayInst = new SelectList(_db.Get_Dropdown("IT"), "ID", "Description");
				ViewBag.cmbTermUsed = new SelectList(_db.Get_Dropdown("TU"), "ID", "Description");
				ViewBag.cmbBancUsed = new SelectList(_db.Get_Dropdown("BN"), "ID", "Description");
				ViewBag.cmbATMUsed = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");
				ViewBag.cmbLoanStat = new SelectList(_db.Get_Dropdown("PS"), "ID", "Description");
				ViewBag.cmbEndorsedTo = new SelectList(_db.Get_Dropdown("ET"), "ID", "Description");
				ViewBag.cmbEndrosedFrom = new SelectList(_db.Get_Dropdown("EF"), "ID", "Description");
				ViewBag.cmbCommType = new SelectList(_db.Get_Dropdown("CT"), "ID", "Description");
				ViewBag.cmbCommChannel = new SelectList(_db.Get_Dropdown("CC"), "ID", "Description");
				ViewBag.cmbStatus = new SelectList(_db.Get_Dropdown("ST"), "ID", "Description");
				ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("CS"), "ID", "Description");
				ViewBag.cmbRemarks = new SelectList(_db.Get_Dropdown("NT"), "ID", "Description");
				ViewBag.cmbReferedBy = new SelectList(_db.Get_Dropdown("RB"), "ID", "Description");
				ViewBag.cmbSubStatus = new SelectList(_db.Get_Dropdown("SS"), "ID", "Description");
				ViewBag.cmbCurrency = new SelectList(_db.Get_Dropdown("CR"), "ID", "Description");
				ViewBag.cmbBranch = new SelectList(_db.Get_Dropdown3(), "ID", "Description");
				ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown4(), "ID", "Description");

				if (edt.ATMUsed == "325")
				{
					//bool isATM = _db.ChkATMUsed(edt.ATMUsed);

					//if (isATM)
					//{
					ViewBag.isATMNga = "meron";
					//	ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown_Location(Convert.ToInt32(edt.ATMUsed)), "Value", "Text");
					//}
				}
				else
				{
					edt.Location2 = edt.Location;
					ViewBag.isATMNga = "wala";

				}


				ViewBag.ListTicketNo = _db.View_TicketNo(edt.TicketNo);

				ViewBag.TktNo = edt.TicketNo;

				return View(edt);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}

			
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public async Task<IActionResult> UpdateCMEntry(EditCallReport cm, List<IFormFile> files)
		{

			try
			{

				//if (string.IsNullOrEmpty(cm.LastName) && string.IsNullOrEmpty(cm.FirstName))
				//{
				//	ModelState.AddModelError(nameof(cm.CustomerName), "Customer Name required");

				//	TempData["msgErr"] = "Mandatory fields required!";
				//}


				//if (string.IsNullOrEmpty(cm.CommChannel))
				//{
				//	ModelState.AddModelError(nameof(cm.CommChannel), "Comm Channel required");

				//	TempData["msgErr"] = "Mandatory fields required!";
				//}

				//if (string.IsNullOrEmpty(cm.ReasonCode))
				//{
				//	ModelState.AddModelError(nameof(cm.ReasonCode), "Reason Code required");

				//	TempData["msgErr"] = "Mandatory fields required!";
				//}				

				if (string.IsNullOrEmpty(cm.Activity))
				{
					ModelState.AddModelError(nameof(cm.Activity), "Activity required");

					TempData["msgErr"] = "Mandatory fields required!";
				}

				//if (string.IsNullOrEmpty(cm.Status))
				//{
				//	ModelState.AddModelError(nameof(cm.Status), "Status required");

				//	TempData["msgErr"] = "Mandatory fields required!";
				//}

				//if (string.IsNullOrEmpty(cm.InOutBound))
				//{
				//	ModelState.AddModelError(nameof(cm.InOutBound), "Comm Type required");

				//	TempData["msgErr"] = "Mandatory fields required!";
				//}


				//if (string.IsNullOrEmpty(cm.Branch))
				//{
				//	ModelState.AddModelError(nameof(cm.Branch), "Branch required");

				//	TempData["msgErr"] = "Mandatory fields required!";
				//}

				//if (!string.IsNullOrEmpty(cm.EndoredTo))
				//{
				//	switch (cm.EndoredTo)
				//	{
				//		case "Others":
				//		case "ATM Ops":

				//			if (string.IsNullOrEmpty(cm.CardNo))
				//			{
				//				ModelState.AddModelError(nameof(cm.CardNo), "Card Number required");

				//				TempData["msgErr"] = "Mandatory fields required!";
				//			}

				//			break;
				//	}

				//	if (string.IsNullOrEmpty(cm.ReferedBy))
				//	{
				//		ModelState.AddModelError(nameof(cm.ReferedBy), "Reffered By required");
				//		TempData["msgErr"] = "Mandatory fields required!";
				//	}

				//	if (string.IsNullOrEmpty(cm.ContactPerson))
				//	{
				//		ModelState.AddModelError(nameof(cm.ContactPerson), "Contact Person required");
				//		TempData["msgErr"] = "Mandatory fields required!";
				//	}

				//}


				if (ModelState.IsValid)
				{
					string cust = (string.IsNullOrEmpty(cm.MiddleName)) ? cm.FirstName + ' ' + cm.LastName : cm.FirstName + ' ' + cm.MiddleName + ' ' + cm.LastName;
					cm.CustomerName = cust;

					string loc = (cm.ATMUsed == "325") ? cm.Location : cm.Location2;

					//bool isType = _db.UpdateType(cm.RID, "C");

					//if (isType)
					//{

					bool isCreated = _db.CreateNewActivity(
						cm.CommChannel,
						cm.ReasonCode,
						cust,
						cm.CardNo,
						cm.Activity,
						cm.Status,
						cm.Agent,
						cm.TicketNo,
						cm.LastName,
						cm.FirstName,
						cm.MiddleName,
						cm.InOutBound,
						cm.TransType,
						loc,
						cm.ATMUsed,
						cm.PaymentOf,
						cm.TerminalUsed,
						cm.Merchant,
						cm.Inter,
						cm.BancnetUsed,
						cm.Online == true ? "Y" : "N",
						cm.Website,
						Convert.ToDouble(cm.Amount),
						cm.DateTrans,
						cm.PLStatus,
						cm.CardBlocked == true ? "Yes" : "No",
						cm.EndoredTo,
						cm.Destination,
						cm.ATMTrans,
						cm.RemitFrom,
						cm.RemitConcern,
						cm.EndorsedFrom,
						cm.Resolved,
						cm.BillerName,
						cm.Remarks,
						cm.ReferedBy,
						cm.ContactPerson,
						cm.Tagging,
						cm.CallTime,
						cm.LocalNum,
						cm.Created_By,
						cm.Last_Update,
						cm.CardPresent,
						cm.SubStatus,
						cm.Currency,
						cm.Branch
						);


					if (isCreated)
					{

						long size = files.Sum(f => f.Length);

						var filePaths = new List<string>();

						foreach (var formFile in files)
						{
							if (formFile.Length > 0)
							{
								var filePath = Path.Combine(_webHostEnvironment.ContentRootPath, @"wwwroot/files", formFile.FileName);

								if (System.IO.File.Exists(filePath))
								{
									System.IO.File.Delete(filePath);
								}

								filePaths.Add(filePath);

								using (var stream = new FileStream(filePath, FileMode.Create))
								{
									bool insertFile = _db.UploadInputFile(cm.TicketNo, formFile.FileName);

									await formFile.CopyToAsync(stream);
								}

							}
						}

						//SendEmail if Endorsed From enabled
						bool chkRCode = _db.ChkIremitCode(cm.ReasonCode);

						if (chkRCode)
						{
							rcd = "Y";
						}
						else
						{
							rcd = "N";
						}

						var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

						//string cardno = cm.CardNo.Substring(0, 6) + "******" + cm.CardNo.Substring(cm.CardNo.Length - 4, 4);

						switch (cm.EndoredTo)
						{
							case "Others":
								bool SendEmail = _db.SendEndorsedMail("1", rcd, cm.EndorsedFrom, cm.TicketNo, cm.Remarks, cm.LastName, cm.FirstName, cm.MiddleName, cm.CardNo, cm.Activity, ses.FullName, cm.ReferedBy, cm.ContactPerson, filePaths);

								break;
							case "ATM Ops":

								bool isATMCreated = _db.CreateNewATMActivity(
									cm.CommChannel,
									cm.ReasonCode,
									cust,
									cm.CardNo,
									cm.Activity,
									cm.Status,
									cm.Agent,
									cm.TicketNo,
									cm.LastName,
									cm.FirstName,
									cm.MiddleName,
									cm.InOutBound,
									cm.TransType,
									loc,
									cm.ATMUsed,
									cm.PaymentOf,
									cm.TerminalUsed,
									cm.Merchant,
									cm.Inter,
									cm.BancnetUsed,
									cm.Online == true ? "Y" : "N",
									cm.Website,
									Convert.ToDouble(cm.Amount),
									cm.DateTrans,
									cm.PLStatus,
									cm.CardBlocked == true ? "Yes" : "No",
									cm.EndoredTo,
									cm.Destination,
									cm.ATMTrans,
									cm.RemitFrom,
									cm.RemitConcern,
									cm.EndorsedFrom,
									cm.Resolved,
									cm.BillerName,
									cm.Remarks,
									cm.ReferedBy,
									cm.ContactPerson,
									cm.Tagging,
									cm.CallTime,
									cm.LocalNum,
									cm.Created_By,
									cm.Last_Update,
									cm.CardPresent,
									cm.SubStatus,
									cm.Currency,
									cm.Branch
									);

								//bool SendEmailATM = _db.SendEndorsedMail("0", rcd, "atm.ops@sterlingbankasia.com", TicketNo, cm.Remarks, cm.LastName, cm.FirstName, cm.MiddleName, cm.CardNo, cm.Activity, ses.FullName, cm.ReferedBy, cm.ContactPerson, filePaths);

								//string cardno = cm.CardNo.Substring(0, 6) + "******" + cm.CardNo.Substring(cm.CardNo.Length - 4, 4);

								bool SendEmailATM = _db.SendEndorsedMail("1", rcd, "atm.ops@sterlingbankasia.com", cm.TicketNo, cm.Remarks, cm.LastName, cm.FirstName, cm.MiddleName, cm.CardNo, cm.Activity, ses.FullName, cm.ReferedBy, cm.ContactPerson, filePaths);

								break;
						}

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isType = _db.UpdateType(cm.RID, "C");

						bool isLocation = _db.UpdateLocation(cm.TicketNo);

						bool isLog = _db.GetAuditLogs("Customer Service - Activity", ses.WrkStn, "Update Activity Ticket No. " + cm.TicketNo);

						TempData["msgOk"] = "Ticket Number " + cm.TicketNo + " has been successfully updated.";

						return RedirectToAction("CustomerService");
					}
					else
					{
						TempData["msgErr"] = "Update was not successfull";

						return View(cm);
					}

					//}
					//else
					//{

					//}



				}

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbChannel = new SelectList(_db.Get_Dropdown("CL"), "ID", "Description");
				ViewBag.cmbTransType = new SelectList(_db.Get_Dropdown("AT"), "ID", "Description");
				ViewBag.cmbPayInst = new SelectList(_db.Get_Dropdown("IT"), "ID", "Description");
				ViewBag.cmbTermUsed = new SelectList(_db.Get_Dropdown("TU"), "ID", "Description");
				ViewBag.cmbBancUsed = new SelectList(_db.Get_Dropdown("BN"), "ID", "Description");
				ViewBag.cmbATMUsed = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");
				ViewBag.cmbLoanStat = new SelectList(_db.Get_Dropdown("PS"), "ID", "Description");
				ViewBag.cmbEndorsedTo = new SelectList(_db.Get_Dropdown("ET"), "ID", "Description");
				ViewBag.cmbEndrosedFrom = new SelectList(_db.Get_Dropdown("EF"), "ID", "Description");
				ViewBag.cmbCommType = new SelectList(_db.Get_Dropdown("CT"), "ID", "Description");
				ViewBag.cmbCommChannel = new SelectList(_db.Get_Dropdown("CC"), "ID", "Description");
				ViewBag.cmbStatus = new SelectList(_db.Get_Dropdown("ST"), "ID", "Description");
				ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("CS"), "ID", "Description");
				ViewBag.cmbRemarks = new SelectList(_db.Get_Dropdown("NT"), "ID", "Description");
				ViewBag.cmbReferedBy = new SelectList(_db.Get_Dropdown("RB"), "ID", "Description");
				ViewBag.cmbCurrency = new SelectList(_db.Get_Dropdown("CR"), "ID", "Description");
				ViewBag.cmbBranch = new SelectList(_db.Get_Dropdown3(), "ID", "Description");
				ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown4(), "ID", "Description");

				ViewBag.ListTicketNo = _db.View_TicketNo(cm.TicketNo);

				ViewBag.TktNo = cm.TicketNo;

			}
			catch (Exception ex)
			{
				string controllerName = ControllerContext.ActionDescriptor.ControllerName;
				string actionName = ControllerContext.ActionDescriptor.ActionName;

				string err = "Error in " + controllerName + " - " + actionName + " : " + ex.ToString();

				TempData["msgErr"] = err;
			}

			return View(cm);

		}


		[HttpGet]
		public JsonResult DelActivity(int id)
		{

			EditCallReport del = _db.ViewCMEntry(id);

			bool isExist = _db.ChkATMActivity(del.TicketNo);

			if (!isExist)
			{
				bool isDeleted = _db.DeleteCMActivity(del.TicketNo);

				if (isDeleted)
				{
					var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

					bool isLog = _db.GetAuditLogs("Customer Service - Activity", ses.WrkStn, "Delete Activity Ticket No. " + del.TicketNo);

					return Json(new { success = true, message = "Delete successful" });
				}
				else
				{
					return Json(new { success = false, message = "Error while Deleting" });
				}

			}
			else
			{
				return Json(new { success = false, message = "You cannot delete this because it exist in the ATM activity!" });
			}

		}

		[HttpGet]
		public IActionResult EditCMEntry(int id)
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Activity", "EditCMEntry");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Activity";
				ViewBag.Page = "CustomerService";

				EditCallReport edt = _db.ViewCMEntry(id);

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbChannel = new SelectList(_db.Get_Dropdown("CL"), "ID", "Description");
				ViewBag.cmbTransType = new SelectList(_db.Get_Dropdown("AT"), "ID", "Description");
				ViewBag.cmbPayInst = new SelectList(_db.Get_Dropdown("IT"), "ID", "Description");
				ViewBag.cmbTermUsed = new SelectList(_db.Get_Dropdown("TU"), "ID", "Description");
				ViewBag.cmbBancUsed = new SelectList(_db.Get_Dropdown("BN"), "ID", "Description");
				ViewBag.cmbATMUsed = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");
				ViewBag.cmbLoanStat = new SelectList(_db.Get_Dropdown("PS"), "ID", "Description");
				ViewBag.cmbEndorsedTo = new SelectList(_db.Get_Dropdown("ET"), "ID", "Description");
				ViewBag.cmbEndrosedFrom = new SelectList(_db.Get_Dropdown("EF"), "ID", "Description");
				ViewBag.cmbCommType = new SelectList(_db.Get_Dropdown("CT"), "ID", "Description");
				ViewBag.cmbCommChannel = new SelectList(_db.Get_Dropdown("CC"), "ID", "Description");
				ViewBag.cmbStatus = new SelectList(_db.Get_Dropdown("ST"), "ID", "Description");
				ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("CS"), "ID", "Description");
				ViewBag.cmbRemarks = new SelectList(_db.Get_Dropdown("NT"), "ID", "Description");
				ViewBag.cmbReferedBy = new SelectList(_db.Get_Dropdown("RB"), "ID", "Description");
				ViewBag.cmbSubStatus = new SelectList(_db.Get_Dropdown("SS"), "ID", "Description");
				ViewBag.cmbCurrency = new SelectList(_db.Get_Dropdown("CR"), "ID", "Description");
				ViewBag.cmbBranch = new SelectList(_db.Get_Dropdown3(), "ID", "Description");

				//if (edt.ATMUsed != string.Empty)
				//{
				//	ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown_Location(Convert.ToInt32(edt.ATMUsed)), "Value", "Text");
				//}

				if (edt.ATMUsed == "325")
				{
					ViewBag.isATMNga = "meron";
				}
				else
				{
					edt.Location2 = edt.Location;
					ViewBag.isATMNga = "wala";
				}

				ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown4(), "ID", "Description");

				ViewBag.ListTicketNo = _db.View_TicketNo(edt.TicketNo);

				ViewBag.TktNo = edt.TicketNo;

				return View(edt);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}

			
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public IActionResult EditCMEntry(EditCallReport cm)
		{

			//Valid_Reason
			//if (!string.IsNullOrEmpty(cm.ReasonCode))
			//{
			//	switch (cm.ReasonCode)
			//	{
			//		case "IPR":

			//			if (string.IsNullOrEmpty(cm.DateTrans))
			//				ModelState.AddModelError(nameof(cm.DateTrans), "Transaction date required");


			//			if (string.IsNullOrEmpty(cm.TransType))
			//				ModelState.AddModelError(nameof(cm.TransType), "Transaction type required");


			//			if (string.IsNullOrEmpty(cm.CardNo))
			//				ModelState.AddModelError(nameof(cm.CardNo), "Card Number required");


			//			break;

			//		case "SPR":
			//		case "NCC":

			//			if (string.IsNullOrEmpty(cm.DateTrans))
			//				ModelState.AddModelError(nameof(cm.DateTrans), "Transaction date required");


			//			if (string.IsNullOrEmpty(cm.TransType))
			//				ModelState.AddModelError(nameof(cm.TransType), "Transaction type required");


			//			if (string.IsNullOrEmpty(cm.CardNo))
			//				ModelState.AddModelError(nameof(cm.CardNo), "Card Number required");


			//			break;

			//		case "SGI":
			//		case "FCP":

			//			if (string.IsNullOrEmpty(cm.DateTrans))
			//				ModelState.AddModelError(nameof(cm.DateTrans), "Transaction date required");


			//			break;

			//		case "SAC":
			//		case "FSI":

			//			if (string.IsNullOrEmpty(cm.DateTrans))
			//				ModelState.AddModelError(nameof(cm.DateTrans), "Transaction date required");

			//			if (string.IsNullOrEmpty(cm.CardNo))
			//				ModelState.AddModelError(nameof(cm.CardNo), "Card Number required");

			//			break;

			//		case "RPR":
			//		case "EPR":
			//		case "FCS":

			//			if (string.IsNullOrEmpty(cm.DateTrans))
			//				ModelState.AddModelError(nameof(cm.DateTrans), "Transaction date required");

			//			if (string.IsNullOrEmpty(cm.CardNo))
			//				ModelState.AddModelError(nameof(cm.CardNo), "Card Number required");

			//			break;

			//		case "ECR":
			//		case "PAI":
			//		case "SCB":

			//			if (string.IsNullOrEmpty(cm.DateTrans))
			//				ModelState.AddModelError(nameof(cm.DateTrans), "Transaction date required");

			//			if (string.IsNullOrEmpty(cm.CardNo))
			//				ModelState.AddModelError(nameof(cm.CardNo), "Card Number required");

			//			break;

			//		case "IPC":
			//		case "PNC":

			//			if (string.IsNullOrEmpty(cm.TransType))
			//				ModelState.AddModelError(nameof(cm.TransType), "Transaction type required");

			//			if (string.IsNullOrEmpty(cm.Destination))
			//				ModelState.AddModelError(nameof(cm.Destination), "Destination required");


			//			if (string.IsNullOrEmpty(cm.ATMTrans))
			//				ModelState.AddModelError(nameof(cm.ATMTrans), "Destination required");

			//			if (string.IsNullOrEmpty(cm.Amount))
			//				ModelState.AddModelError(nameof(cm.Amount), "Amount required");

			//			if (string.IsNullOrEmpty(cm.Currency))
			//				ModelState.AddModelError(nameof(cm.Currency), "Currency required");

			//			break;
			//	}
			//}

			//if (string.IsNullOrEmpty(cm.TransType))
			//{

			//	switch (cm.ReasonCode)
			//	{
			//		case "IPC":
			//		case "PNC":
			//			if (string.IsNullOrEmpty(cm.Destination))
			//				ModelState.AddModelError(nameof(cm.Destination), "Destination required");

			//			if (string.IsNullOrEmpty(cm.ATMTrans))
			//				ModelState.AddModelError(nameof(cm.ATMTrans), "Transaction Type required");
			//			break;
			//	}

			//}
			//else
			//{
			//	switch (cm.TransType)
			//	{
			//		case "ATM":

			//			if (string.IsNullOrEmpty(cm.Location))
			//				ModelState.AddModelError(nameof(cm.Location), "Location required");

			//			if (string.IsNullOrEmpty(cm.ATMUsed))
			//				ModelState.AddModelError(nameof(cm.ATMUsed), "Location required");

			//			if (string.IsNullOrEmpty(cm.Amount))
			//				ModelState.AddModelError(nameof(cm.Amount), "Amount required");

			//			if (string.IsNullOrEmpty(cm.Currency))
			//				ModelState.AddModelError(nameof(cm.Currency), "Currency required");


			//			switch (cm.ReasonCode)
			//			{
			//				case "IPC":
			//				case "PNC":

			//					if (string.IsNullOrEmpty(cm.Destination))
			//						ModelState.AddModelError(nameof(cm.Destination), "Destination required");

			//					if (string.IsNullOrEmpty(cm.ATMTrans))
			//						ModelState.AddModelError(nameof(cm.ATMTrans), "Transaction Type required");

			//					break;
			//				default:

			//					if (string.IsNullOrEmpty(cm.ATMTrans))
			//						ModelState.AddModelError(nameof(cm.ATMTrans), "Transaction Type required");

			//					break;
			//			}

			//			break;

			//		case "BILLS PAYMENT":

			//			if (string.IsNullOrEmpty(cm.Amount))
			//				ModelState.AddModelError(nameof(cm.Amount), "Amount required");

			//			if (string.IsNullOrEmpty(cm.Currency))
			//				ModelState.AddModelError(nameof(cm.Currency), "Currency required");

			//			if (string.IsNullOrEmpty(cm.PaymentOf))
			//				ModelState.AddModelError(nameof(cm.PaymentOf), "Payment Institution required");

			//			if (string.IsNullOrEmpty(cm.BillerName))
			//				ModelState.AddModelError(nameof(cm.BillerName), "Biller Name required");

			//			if (string.IsNullOrEmpty(cm.Merchant))
			//				ModelState.AddModelError(nameof(cm.Merchant), "Merchant / Website required");

			//			if (string.IsNullOrEmpty(cm.BancnetUsed))
			//				ModelState.AddModelError(nameof(cm.BancnetUsed), "Bancnet ATM Used required");

			//			switch (cm.ReasonCode)
			//			{
			//				case "IPC":
			//				case "PNC":

			//					if (string.IsNullOrEmpty(cm.Destination))
			//						ModelState.AddModelError(nameof(cm.Destination), "Destination required");

			//					if (string.IsNullOrEmpty(cm.ATMTrans))
			//						ModelState.AddModelError(nameof(cm.ATMTrans), "Transaction Type required");

			//					break;

			//			}

			//			break;

			//		case "POS":

			//			if (string.IsNullOrEmpty(cm.Amount))
			//				ModelState.AddModelError(nameof(cm.Amount), "Amount required");

			//			if (string.IsNullOrEmpty(cm.Currency))
			//				ModelState.AddModelError(nameof(cm.Currency), "Currency required");

			//			if (string.IsNullOrEmpty(cm.TerminalUsed))
			//				ModelState.AddModelError(nameof(cm.TerminalUsed), "Terminal Used required");

			//			if (string.IsNullOrEmpty(cm.BillerName))
			//				ModelState.AddModelError(nameof(cm.BillerName), "Biller Name required");


			//			if (string.IsNullOrEmpty(cm.Merchant))
			//				ModelState.AddModelError(nameof(cm.Merchant), "Merchant / Website required");

			//			switch (cm.ReasonCode)
			//			{
			//				case "IPC":
			//				case "PNC":

			//					if (string.IsNullOrEmpty(cm.Destination))
			//						ModelState.AddModelError(nameof(cm.Destination), "Destination required");

			//					if (string.IsNullOrEmpty(cm.ATMTrans))
			//						ModelState.AddModelError(nameof(cm.ATMTrans), "Transaction Type required");

			//					break;
			//			}



			//			if (string.IsNullOrEmpty(cm.CardPresent))
			//				ModelState.AddModelError(nameof(cm.CardPresent), "Card Present/Not Present required");


			//			break;

			//		case "IBFT":


			//			if (string.IsNullOrEmpty(cm.Amount))
			//				ModelState.AddModelError(nameof(cm.Amount), "Amount required");

			//			if (string.IsNullOrEmpty(cm.Currency))
			//				ModelState.AddModelError(nameof(cm.Currency), "Currency required");

			//			if (string.IsNullOrEmpty(cm.Inter))
			//				ModelState.AddModelError(nameof(cm.Inter), "Ineter/Intra required");

			//			if (string.IsNullOrEmpty(cm.BancnetUsed))
			//				ModelState.AddModelError(nameof(cm.BancnetUsed), "Bancnet ATM Used required");

			//			switch (cm.ReasonCode)
			//			{
			//				case "IPC":
			//				case "PNC":

			//					if (string.IsNullOrEmpty(cm.Destination))
			//						ModelState.AddModelError(nameof(cm.Destination), "Destination required");

			//					if (string.IsNullOrEmpty(cm.ATMTrans))
			//						ModelState.AddModelError(nameof(cm.ATMTrans), "Transaction Type required");

			//					break;
			//				default:
			//					if (string.IsNullOrEmpty(cm.ATMTrans))
			//						ModelState.AddModelError(nameof(cm.ATMTrans), "Transaction Type required");

			//					break;
			//			}


			//			break;

			//		case "REMITTANCE":


			//			switch (cm.ReasonCode)
			//			{
			//				case "IPC":
			//				case "PNC":
			//					if (string.IsNullOrEmpty(cm.Destination))
			//						ModelState.AddModelError(nameof(cm.Destination), "Destination required");

			//					if (string.IsNullOrEmpty(cm.ATMTrans))
			//						ModelState.AddModelError(nameof(cm.ATMTrans), "Transaction Type required");

			//					break;
			//			}

			//			if (string.IsNullOrEmpty(cm.RemitFrom))
			//				ModelState.AddModelError(nameof(cm.RemitFrom), "Remittance From required");

			//			if (string.IsNullOrEmpty(cm.RemitConcern))
			//				ModelState.AddModelError(nameof(cm.RemitConcern), "Remittance concern required");

			//			break;

			//		default:
			//			break;
			//	}
			//}


			//if (!string.IsNullOrEmpty(cm.ReasonCode))
			//{
			//	switch (cm.ReasonCode)
			//	{
			//		case "APR":
			//		case "BOC":
			//		case "BRC":
			//		case "BPR":
			//		case "DPR":
			//		case "EPR":
			//		case "FCP":
			//		case "FCS":
			//		case "NPR":
			//		case "OPR":
			//		case "PPR":
			//		case "RPR":
			//		case "UPR":
			//		case "USR":
			//		case "SPR":
			//		case "NCC":
			//			if (string.IsNullOrEmpty(cm.Tagging))
			//				ModelState.AddModelError(nameof(cm.Tagging), "Classification required");
			//			break;
			//		default:
			//			break;
			//	}
			//}

			//if (string.IsNullOrEmpty(cm.LastName))
			//	ModelState.AddModelError(nameof(cm.LastName), "Last Name required");

			//if (string.IsNullOrEmpty(cm.FirstName))
			//	ModelState.AddModelError(nameof(cm.FirstName), "FirstName Name required");

			try
			{
				//if (!string.IsNullOrEmpty(cm.EndoredTo))
				//{
				//	switch (cm.EndoredTo)
				//	{
				//		case "Others":
				//		case "ATM Ops":

				//			if (string.IsNullOrEmpty(cm.CardNo))
				//				ModelState.AddModelError(nameof(cm.CardNo), "Card Number required");

				//			break;
				//	}
				//}

				if (ModelState.IsValid)
				{

					cm.CustomerName = (string.IsNullOrEmpty(cm.MiddleName)) ? cm.FirstName + ' ' + cm.LastName : cm.FirstName + ' ' + cm.MiddleName + ' ' + cm.LastName;

					//string loc = (cm.ATMUsed == "325") ? cm.Location : cm.Location2;

					bool isEdit = _db.EditNewActivity(cm);

					if (isEdit)
					{

						var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("Customer Service - Activity", ses.WrkStn, "Edit Activity Ticket No. " + cm.TicketNo);

						TempData["msgOk"] = "Ticket " + cm.TicketNo + " successfully Edited.";
						return RedirectToAction("CustomerService");

					}
					else
					{
						TempData["msgErr"] = "Editing of this ticket was not successfull";
						return View(cm);
					}


					//bool isType = _db.UpdateType(cm.RID);

					//if (isType)
					//{

					//	bool isCreated = _db.CreateNewActivity(
					//		cm.CommChannel,
					//		cm.ReasonCode,
					//		cm.CustomerName,
					//		cm.CardNo,
					//		cm.Activity,
					//		cm.Status,
					//		cm.Agent,
					//		cm.TicketNo,
					//		cm.LastName,
					//		cm.FirstName,
					//		cm.MiddleName,
					//		cm.InOutBound,
					//		cm.TransType,
					//		cm.Location,
					//		cm.ATMUsed,
					//		cm.PaymentOf,
					//		cm.TerminalUsed,
					//		cm.Merchant,
					//		cm.Inter,
					//		cm.BancnetUsed,
					//		cm.Online == true ? "Y" : "N",
					//		cm.Website,
					//		Convert.ToDouble(cm.Amount),
					//		cm.DateTrans,
					//		cm.PLStatus,
					//		cm.CardBlocked == true ? "Yes" : "No",
					//		cm.EndoredTo,
					//		cm.Destination,
					//		cm.ATMTrans,
					//		cm.RemitFrom,
					//		cm.RemitConcern,
					//		cm.EndorsedFrom,
					//		cm.Resolved,
					//		cm.BillerName,
					//		cm.Remarks,
					//		cm.ReferedBy,
					//		cm.ContactPerson,
					//		cm.Tagging,
					//		cm.CallTime,
					//		cm.LocalNum,
					//		cm.Created_By,
					//		cm.Last_Update,
					//		cm.CardPresent,
					//		cm.SubStatus
					//		);


					//	if (isCreated)
					//	{

					//		long size = files.Sum(f => f.Length);

					//		var filePaths = new List<string>();

					//		foreach (var formFile in files)
					//		{
					//			if (formFile.Length > 0)
					//			{
					//				var filePath = Path.Combine(_webHostEnvironment.ContentRootPath, @"wwwroot/files", formFile.FileName);

					//				filePaths.Add(filePath);

					//				using (var stream = new FileStream(filePath, FileMode.Create))
					//				{
					//					bool insertFile = _db.UploadInputFile(cm.TicketNo, formFile.FileName);

					//					await formFile.CopyToAsync(stream);
					//				}

					//			}
					//		}

					//		//SendEmail if Endorsed From enabled
					//		bool chkRCode = _db.ChkIremitCode(cm.ReasonCode);

					//		if (chkRCode)
					//		{
					//			rcd = "Y";
					//		}
					//		else
					//		{
					//			rcd = "N";
					//		}

					//		var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					//		switch (cm.EndoredTo)
					//		{
					//			case "Others":
					//				bool SendEmail = _db.SendEndorsedMail("1", rcd, cm.EndorsedFrom, cm.TicketNo, cm.Remarks, cm.LastName, cm.FirstName, cm.MiddleName, cm.CardNo, cm.Activity, ses.FullName, cm.ReferedBy, cm.ContactPerson, filePaths);

					//				break;
					//			case "ATM Ops":

					//				bool isATMCreated = _db.CreateNewATMActivity(
					//					cm.CommChannel,
					//					cm.ReasonCode,
					//					cm.CustomerName,
					//					cm.CardNo,
					//					cm.Activity,
					//					cm.Status,
					//					cm.Agent,
					//					cm.TicketNo,
					//					cm.LastName,
					//					cm.FirstName,
					//					cm.MiddleName,
					//					cm.InOutBound,
					//					cm.TransType,
					//					cm.Location,
					//					cm.ATMUsed,
					//					cm.PaymentOf,
					//					cm.TerminalUsed,
					//					cm.Merchant,
					//					cm.Inter,
					//					cm.BancnetUsed,
					//					cm.Online == true ? "Y" : "N",
					//					cm.Website,
					//					Convert.ToDouble(cm.Amount),
					//					cm.DateTrans,
					//					cm.PLStatus,
					//					cm.CardBlocked == true ? "Yes" : "No",
					//					cm.EndoredTo,
					//					cm.Destination,
					//					cm.ATMTrans,
					//					cm.RemitFrom,
					//					cm.RemitConcern,
					//					cm.EndorsedFrom,
					//					cm.Resolved,
					//					cm.BillerName,
					//					cm.Remarks,
					//					cm.ReferedBy,
					//					cm.ContactPerson,
					//					cm.Tagging,
					//					cm.CallTime,
					//					cm.LocalNum,
					//					cm.Created_By,
					//					cm.Last_Update,
					//					cm.CardPresent,
					//					cm.SubStatus
					//					);

					//				//bool SendEmailATM = _db.SendEndorsedMail("0", rcd, "atm.ops@sterlingbankasia.com", TicketNo, cm.Remarks, cm.LastName, cm.FirstName, cm.MiddleName, cm.CardNo, cm.Activity, ses.FullName, cm.ReferedBy, cm.ContactPerson, filePaths);

					//				bool SendEmailATM = _db.SendEndorsedMail("1", rcd, "aalivara@sterlingbankasia.com", cm.TicketNo, cm.Remarks, cm.LastName, cm.FirstName, cm.MiddleName, cm.CardNo, cm.Activity, ses.FullName, cm.ReferedBy, cm.ContactPerson, filePaths);
					//				break;
					//		}


					//		TempData["msgOk"] = "Ticket " + cm.TicketNo + " successfully updated.";

					//		return RedirectToAction("CustomerService");
					//	}
					//	else
					//	{
					//		TempData["msgErr"] = "Creation of new ticket was not successfull";

					//		return View(cm);
					//	}

					//}

				}

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbChannel = new SelectList(_db.Get_Dropdown("CL"), "ID", "Description");
				ViewBag.cmbTransType = new SelectList(_db.Get_Dropdown("AT"), "ID", "Description");
				ViewBag.cmbPayInst = new SelectList(_db.Get_Dropdown("IT"), "ID", "Description");
				ViewBag.cmbTermUsed = new SelectList(_db.Get_Dropdown("TU"), "ID", "Description");
				ViewBag.cmbBancUsed = new SelectList(_db.Get_Dropdown("BN"), "ID", "Description");
				ViewBag.cmbATMUsed = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");
				ViewBag.cmbLoanStat = new SelectList(_db.Get_Dropdown("PS"), "ID", "Description");
				ViewBag.cmbEndorsedTo = new SelectList(_db.Get_Dropdown("ET"), "ID", "Description");
				ViewBag.cmbEndrosedFrom = new SelectList(_db.Get_Dropdown("EF"), "ID", "Description");
				ViewBag.cmbCommType = new SelectList(_db.Get_Dropdown("CT"), "ID", "Description");
				ViewBag.cmbCommChannel = new SelectList(_db.Get_Dropdown("CC"), "ID", "Description");
				ViewBag.cmbStatus = new SelectList(_db.Get_Dropdown("ST"), "ID", "Description");
				ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("CS"), "ID", "Description");
				ViewBag.cmbRemarks = new SelectList(_db.Get_Dropdown("NT"), "ID", "Description");
				ViewBag.cmbReferedBy = new SelectList(_db.Get_Dropdown("RB"), "ID", "Description");
				ViewBag.cmbCurrency = new SelectList(_db.Get_Dropdown("CR"), "ID", "Description");
				ViewBag.cmbBranch = new SelectList(_db.Get_Dropdown3(), "ID", "Description");
				ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown4(), "ID", "Description");
			}
			catch (Exception ex)
			{
				string controllerName = ControllerContext.ActionDescriptor.ControllerName;
				string actionName = ControllerContext.ActionDescriptor.ActionName;

				string err = "Error in " + controllerName + " - " + actionName + " : " + ex.ToString();

				TempData["msgErr"] = err;
			}

			return View(cm);

		}

		[HttpGet]
		public IActionResult ATMEntry(int id)
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Activity", "ATMEntry");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "ATM";
				ViewBag.Page = "ATMOps";

				EditATMReport edt = _db.ViewATMEntry(id);

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbChannel = new SelectList(_db.Get_Dropdown("CL"), "ID", "Description");
				ViewBag.cmbTransType = new SelectList(_db.Get_Dropdown("AT"), "ID", "Description");
				ViewBag.cmbPayInst = new SelectList(_db.Get_Dropdown("IT"), "ID", "Description");
				ViewBag.cmbTermUsed = new SelectList(_db.Get_Dropdown("TU"), "ID", "Description");
				ViewBag.cmbBancUsed = new SelectList(_db.Get_Dropdown("BN"), "ID", "Description");
				ViewBag.cmbATMUsed = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");
				ViewBag.cmbLoanStat = new SelectList(_db.Get_Dropdown("PS"), "ID", "Description");
				ViewBag.cmbEndorsedTo = new SelectList(_db.Get_Dropdown("ET"), "ID", "Description");
				ViewBag.cmbEndrosedFrom = new SelectList(_db.Get_Dropdown("EF"), "ID", "Description");
				ViewBag.cmbCommType = new SelectList(_db.Get_Dropdown("CT"), "ID", "Description");
				ViewBag.cmbCommChannel = new SelectList(_db.Get_Dropdown("CC"), "ID", "Description");
				ViewBag.cmbStatus = new SelectList(_db.Get_Dropdown("ST"), "ID", "Description");
				ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("CS"), "ID", "Description");
				ViewBag.cmbRemarks = new SelectList(_db.Get_Dropdown("NT"), "ID", "Description");
				ViewBag.cmbReferedBy = new SelectList(_db.Get_Dropdown("RB"), "ID", "Description");
				ViewBag.cmbSubStatus = new SelectList(_db.Get_Dropdown("SS"), "ID", "Description");
				ViewBag.cmbCurrency = new SelectList(_db.Get_Dropdown("CR"), "ID", "Description");
				ViewBag.cmbBranch = new SelectList(_db.Get_Dropdown3(), "ID", "Description");
				ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown4(), "ID", "Description");

				if (edt.ATMUsed != string.Empty)
				{
					bool isATM = _db.ChkATMUsed(edt.ATMUsed);

					if (isATM)
					{
						ViewBag.isATMNga = "meron";
						ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown_Location(Convert.ToInt32(edt.ATMUsed)), "Value", "Text");
					}
					else
					{
						ViewBag.isATMNga = "wala";
					}

				}

				HttpContext.Session.SetString("updatedby", ses.LoginID);

				ViewBag.ListTicketNo = _db.View_TicketATM(edt.TicketNo);

				ViewBag.TktNo = edt.TicketNo;

				return View(edt);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}

			
		}


		[HttpPost]
		[ValidateAntiForgeryToken]
		public async Task<IActionResult> ATMEntry(EditATMReport cm, List<IFormFile> files)
		{			
			try
			{


				if (!string.IsNullOrEmpty(cm.Status))
				{
					switch (cm.Status)
					{
						case "Endorsed":

							if (string.IsNullOrEmpty(cm.SubStatus))
								ModelState.AddModelError(nameof(cm.SubStatus), "Sub Status required");

							break;
					}
				}

				if (!string.IsNullOrEmpty(cm.EndoredTo))
				{
					
					if (string.IsNullOrEmpty(cm.ReferedBy))
					{
						ModelState.AddModelError(nameof(cm.ReferedBy), "Reffered By required");
						TempData["msgErr"] = "Mandatory fields required!";
					}

					if (string.IsNullOrEmpty(cm.ContactPerson))
					{
						ModelState.AddModelError(nameof(cm.ContactPerson), "Contact Person required");
						TempData["msgErr"] = "Mandatory fields required!";
					}
				}

				if (ModelState.IsValid)
				{
					var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					cm.CustomerName = (string.IsNullOrEmpty(cm.MiddleName)) ? cm.FirstName + ' ' + cm.LastName : cm.FirstName + ' ' + cm.MiddleName + ' ' + cm.LastName;

					//bool isType = _db.UpdateType(cm.RID, "A");

					//if (isType)
					//{

					bool isCreated = _db.CreateATMActivity(
						cm.CommChannel,
						cm.ReasonCode,
						cm.CustomerName,
						cm.CardNo,
						cm.Activity,
						cm.Status,
						ses.LoginID,
						cm.TicketNo,
						cm.LastName,
						cm.FirstName,
						cm.MiddleName,
						cm.InOutBound,
						cm.TransType,
						cm.Location,
						cm.ATMUsed,
						cm.PaymentOf,
						cm.TerminalUsed,
						cm.Merchant,
						cm.Inter,
						cm.BancnetUsed,
						cm.Online == true ? "Y" : "N",
						cm.Website,
						Convert.ToDouble(cm.Amount),
						cm.DateTrans,
						cm.PLStatus,
						cm.CardBlocked == true ? "Yes" : "No",
						cm.EndoredTo,
						cm.Destination,
						cm.ATMTrans,
						cm.RemitFrom,
						cm.RemitConcern,
						cm.BillerName,
						cm.Remarks,
						cm.ReferedBy,
						cm.ContactPerson,
						cm.Tagging,
						cm.CardPresent,
						cm.SubStatus,
						cm.Currency,
						cm.Branch,
						cm.NameVerified,
						cm.Resolved,
						ses.LoginID,
						cm.Created_By
					);


					if (isCreated)
					{

						long size = files.Sum(f => f.Length);

						var filePaths = new List<string>();

						foreach (var formFile in files)
						{
							if (formFile.Length > 0)
							{
								var filePath = Path.Combine(_webHostEnvironment.ContentRootPath, @"wwwroot/files", formFile.FileName);

								if (System.IO.File.Exists(filePath))
								{
									System.IO.File.Delete(filePath);
								}

								filePaths.Add(filePath);

								using (var stream = new FileStream(filePath, FileMode.Create))
								{
									bool insertFile = _db.UploadInputFile(cm.TicketNo, formFile.FileName);

									await formFile.CopyToAsync(stream);
								}

							}
						}

						//SendEmail if Endorsed From enabled
						bool chkRCode = _db.ChkIremitCode(cm.ReasonCode);

						if (chkRCode)
						{
							rcd = "Y";
						}
						else
						{
							rcd = "N";
						}



						switch (cm.EndoredTo)
						{
							case "Customer Management":

								MaxRID _rid = _db.Get_RID(cm.TicketNo, "C");

								//int _id = _rid.RID;

								bool isCMCreated = _db.CreateNewActivity(
									cm.CommChannel,
									cm.ReasonCode,
									cm.CustomerName,
									cm.CardNo,
									cm.Activity,
									cm.Status,
									ses.LoginID,
									cm.TicketNo,
									cm.LastName,
									cm.FirstName,
									cm.MiddleName,
									cm.InOutBound,
									cm.TransType,
									cm.Location,
									cm.ATMUsed,
									cm.PaymentOf,
									cm.TerminalUsed,
									cm.Merchant,
									cm.Inter,
									cm.BancnetUsed,
									cm.Online == true ? "Y" : "N",
									cm.Website,
									Convert.ToDouble(cm.Amount),
									cm.DateTrans,
									cm.PLStatus,
									"",
									cm.EndoredTo,
									cm.Destination,
									cm.ATMTrans,
									cm.RemitFrom,
									cm.RemitConcern,
									"",
									cm.Resolved,
									cm.BillerName,
									cm.Remarks,
									cm.ReferedBy,
									cm.ContactPerson,
									cm.Tagging,
									"",
									"",
									cm.Created_By,
									ses.LoginID,
									cm.CardPresent,
									cm.SubStatus,
									cm.Currency,
									cm.Branch
									); ;

								if (isCMCreated)
								{
									bool isTypeC = _db.UpdateType(_rid.RID, "C");

									bool SendEmail = _db.SendEndorsedMailATM(rcd, cm.TicketNo, cm.Remarks, cm.LastName, cm.FirstName, cm.MiddleName, cm.CardNo, cm.Activity, ses.FullName, cm.ReferedBy, cm.ContactPerson, filePaths);
								}

								break;

						}

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isType = _db.UpdateType(cm.RID, "A");

						bool isLog = _db.GetAuditLogs("ATM Ops - Activity", ses.WrkStn, "ATM Ops Activity Update for Ticket No. " + cm.TicketNo);

						TempData["msgOk"] = "Ticket " + cm.TicketNo + " successfully updated.";

						return RedirectToAction("ATMOps");
					}
					else
					{
						TempData["msgErr"] = "Creation of new ticket was not successfull";

						return View(cm);
					}

					//}

				}

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbChannel = new SelectList(_db.Get_Dropdown("CL"), "ID", "Description");
				ViewBag.cmbTransType = new SelectList(_db.Get_Dropdown("AT"), "ID", "Description");
				ViewBag.cmbPayInst = new SelectList(_db.Get_Dropdown("IT"), "ID", "Description");
				ViewBag.cmbTermUsed = new SelectList(_db.Get_Dropdown("TU"), "ID", "Description");
				ViewBag.cmbBancUsed = new SelectList(_db.Get_Dropdown("BN"), "ID", "Description");
				ViewBag.cmbATMUsed = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");
				ViewBag.cmbLoanStat = new SelectList(_db.Get_Dropdown("PS"), "ID", "Description");
				ViewBag.cmbEndorsedTo = new SelectList(_db.Get_Dropdown("ET"), "ID", "Description");
				ViewBag.cmbEndrosedFrom = new SelectList(_db.Get_Dropdown("EF"), "ID", "Description");
				ViewBag.cmbCommType = new SelectList(_db.Get_Dropdown("CT"), "ID", "Description");
				ViewBag.cmbCommChannel = new SelectList(_db.Get_Dropdown("CC"), "ID", "Description");
				ViewBag.cmbStatus = new SelectList(_db.Get_Dropdown("ST"), "ID", "Description");
				ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("CS"), "ID", "Description");
				ViewBag.cmbRemarks = new SelectList(_db.Get_Dropdown("NT"), "ID", "Description");
				ViewBag.cmbReferedBy = new SelectList(_db.Get_Dropdown("RB"), "ID", "Description");
				ViewBag.cmbCurrency = new SelectList(_db.Get_Dropdown("CR"), "ID", "Description");
				ViewBag.cmbBranch = new SelectList(_db.Get_Dropdown3(), "ID", "Description");
				ViewBag.cmbSubStatus = new SelectList(_db.Get_Dropdown("SS"), "ID", "Description");
				ViewBag.cmbLocation = new SelectList(_db.Get_Dropdown4(), "ID", "Description");

			}
			catch (Exception ex)
			{
				string controllerName = ControllerContext.ActionDescriptor.ControllerName;
				string actionName = ControllerContext.ActionDescriptor.ActionName;

				string err = "Error in " + controllerName + " - " + actionName + " : " + ex.ToString();

				TempData["msgErr"] = err;
			}


			return View(cm);

		}

		public IActionResult ExportATMActivity()
		{
			var stream = new System.IO.MemoryStream();

			var ses = JsonConvert.DeserializeObject<ActivitySearch>(HttpContext.Session.GetString("SearchActivity"));

			var myInput = HttpContext.Session.GetString("SearchValue");

			string fnam = string.Empty;

			var v = _db.View_ATMActivity(ses);

			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("ATMOps", "Activity");
			}

			using (ExcelPackage package = new(stream))
			{

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "TICKET NO";
				ws.Cells["B1"].Value = "INBOUND/OUTBOUND";
				ws.Cells["C1"].Value = "COMM CHANNEL";
				ws.Cells["D1"].Value = "DATE RECEIVED";
				ws.Cells["E1"].Value = "TIME RECEIVED";
				ws.Cells["F1"].Value = "REASON CODE";
				ws.Cells["G1"].Value = "CUSTOMER NAME";
				ws.Cells["H1"].Value = "CARD NO";
				ws.Cells["I1"].Value = "DESTINATION";
				ws.Cells["J1"].Value = "TRANS TYPE";
				ws.Cells["K1"].Value = "ATM TRANS";
				ws.Cells["L1"].Value = "LOCATION";
				ws.Cells["M1"].Value = "ATM USED";
				ws.Cells["N1"].Value = "PAYMENT OF";
				ws.Cells["O1"].Value = "TERMINAL USED";
				ws.Cells["P1"].Value = "BILLERNAME";
				ws.Cells["Q1"].Value = "MERCHANT";
				ws.Cells["R1"].Value = "INTER";
				ws.Cells["S1"].Value = "BANCNET USED";
				ws.Cells["T1"].Value = "ONLINE";
				ws.Cells["U1"].Value = "WEB SITE";
				ws.Cells["V1"].Value = "REMIT FROM";
				ws.Cells["W1"].Value = "REMIT CONCERN";
				ws.Cells["X1"].Value = "AMOUNT";
				ws.Cells["Y1"].Value = "CURRENCY";
				ws.Cells["Z1"].Value = "DATE TRANS";
				ws.Cells["AA1"].Value = "PL STATUS";
				ws.Cells["AB1"].Value = "CARD BLOCKED";
				ws.Cells["AC1"].Value = "ACTIVITY";
				ws.Cells["AD1"].Value = "STATUS";
				ws.Cells["AE1"].Value = "SUB STATUS";
				ws.Cells["AF1"].Value = "ENDORSED TO";
				ws.Cells["AG1"].Value = "DUE DATE";
				ws.Cells["AH1"].Value = "RESOLVED DATE";
				ws.Cells["AI1"].Value = "NATURE OF COMPLAINT";
				ws.Cells["AJ1"].Value = "REFERRED BY";
				ws.Cells["AK1"].Value = "CONTACT PERSON";
				ws.Cells["AL1"].Value = "CLASSIFICATION";
				ws.Cells["AM1"].Value = "CARD PRESENT";
				ws.Cells["AN1"].Value = "BRANCH";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.TicketNo;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.InOutBound;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.CommChannel;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.DateReceived.ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.DateReceived.ToString("HH:mm:ss");
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.abbreviation;
					ws.Cells[string.Format("G{0}", rowStart)].Value = item.CustomerName;
					ws.Cells[string.Format("H{0}", rowStart)].Value = item.CardNo;
					ws.Cells[string.Format("I{0}", rowStart)].Value = item.Destination;
					ws.Cells[string.Format("J{0}", rowStart)].Value = item.TransType;
					ws.Cells[string.Format("K{0}", rowStart)].Value = item.ATMTrans;
					ws.Cells[string.Format("L{0}", rowStart)].Value = item.Location;
					ws.Cells[string.Format("M{0}", rowStart)].Value = item.ATMUsed;
					ws.Cells[string.Format("N{0}", rowStart)].Value = item.PaymentOf;
					ws.Cells[string.Format("O{0}", rowStart)].Value = item.TerminalUsed;
					ws.Cells[string.Format("P{0}", rowStart)].Value = item.BillerName;
					ws.Cells[string.Format("Q{0}", rowStart)].Value = item.Merchant;
					ws.Cells[string.Format("R{0}", rowStart)].Value = item.Inter;
					ws.Cells[string.Format("S{0}", rowStart)].Value = item.BancnetUsed;
					ws.Cells[string.Format("T{0}", rowStart)].Value = item.Online;
					ws.Cells[string.Format("U{0}", rowStart)].Value = item.Website;
					ws.Cells[string.Format("V{0}", rowStart)].Value = item.RemitFrom;
					ws.Cells[string.Format("W{0}", rowStart)].Value = item.RemitConcern;
					ws.Cells[string.Format("X{0}", rowStart)].Value = item.Amount;
					ws.Cells[string.Format("Y{0}", rowStart)].Value = item.Currency;
					ws.Cells[string.Format("Z{0}", rowStart)].Value = (item.DateTrans == null) ? "" : Convert.ToDateTime(item.DateTrans).ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("AA{0}", rowStart)].Value = item.PLStatus;
					ws.Cells[string.Format("AB{0}", rowStart)].Value = item.CardBlocked;
					ws.Cells[string.Format("AC{0}", rowStart)].Value = item.Activity;
					ws.Cells[string.Format("AD{0}", rowStart)].Value = item.Status;
					ws.Cells[string.Format("AE{0}", rowStart)].Value = item.SubStatus;
					ws.Cells[string.Format("AF{0}", rowStart)].Value = item.EndoredTo;
					ws.Cells[string.Format("AG{0}", rowStart)].Value = (item.Resolved == null) ? "" : Convert.ToDateTime(item.Resolved).ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("AH{0}", rowStart)].Value = (item.ResolvedDate == null) ? "" : Convert.ToDateTime(item.ResolvedDate).ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("AI{0}", rowStart)].Value = item.Remarks;
					ws.Cells[string.Format("AJ{0}", rowStart)].Value = item.ReferedBy;
					ws.Cells[string.Format("AK{0}", rowStart)].Value = item.ContactPerson;
					ws.Cells[string.Format("AL{0}", rowStart)].Value = item.Tagging;
					ws.Cells[string.Format("AM{0}", rowStart)].Value = item.CardPresent;
					ws.Cells[string.Format("AN{0}", rowStart)].Value = item.Branch;

					rowStart++;
				}

				ws.Cells["A:AM"].AutoFitColumns();
				ws.Cells["X:X"].Style.Numberformat.Format = "#,##0.00";

				package.Save();
			}

			fnam = "ATM OPS Activity " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		[HttpGet]
		public IActionResult ResolvedActivity()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Activity", "ResolvedActivity");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "ATM";
				ViewBag.Page = "ATMOps";

				HttpContext.Session.SetString("deyt1", DateTime.Now.ToString("dd-MMM-yyyy"));
				HttpContext.Session.SetString("deyt2", DateTime.Now.ToString("dd-MMM-yyyy"));

				var vw = _db.ListATMOpsResolved(DateTime.Now.ToString("dd-MMM-yyyy"), DateTime.Now.ToString("dd-MMM-yyyy"));

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("ATM Ops - Resolved Activity", ses.WrkStn, "View List of ATM Ops Resolved Activity");

				ViewBag.ListATMResolved = vw;

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}

		}

		[HttpPost]
		public IActionResult ResolvedActivity(DateTime dt1, DateTime dt2)
		{
			try
			{
				ViewBag.Interface = "ATM";
				ViewBag.Page = "ATMOps";

				HttpContext.Session.SetString("deyt1", dt1.ToString("dd-MMM-yyyy"));
				HttpContext.Session.SetString("deyt2", dt2.ToString("dd-MMM-yyyy"));

				var vw = _db.ListATMOpsResolved(dt1.ToString("dd-MMM-yyyy"), dt2.ToString("dd-MMM-yyyy"));

				ViewBag.ListATMResolved = vw;
				return View(vw);

			}
			catch (Exception ex)
			{
				string controllerName = ControllerContext.ActionDescriptor.ControllerName;
				string actionName = ControllerContext.ActionDescriptor.ActionName;

				string err = "Error in " + controllerName + " - " + actionName + " : " + ex.ToString();

				TempData["msgErr"] = err;
			}

			return View();

		}

		public IActionResult ExportATMResolved()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			//var card = HttpContext.Session.GetString("cardno");

			string fnam = string.Empty;

			var v = _db.ListATMOpsResolved(dt1, dt2);

			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("ATMOps", "Activity");
			}


			using (ExcelPackage package = new(stream))
			{

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.TicketNo.ToLower().Contains(myInput.ToLower()) ||
								 p.InOutBound.ToLower().Contains(myInput.ToLower()) ||
								 p.CommChannel.ToLower().Contains(myInput.ToLower()) ||
								 p.DateReceived.ToString().Contains(myInput) ||
								 p.abbreviation.ToLower().Contains(myInput.ToLower()) ||
								 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
								 p.CardNo.ToLower().Contains(myInput.ToLower()) ||
								 p.Destination.ToLower().Contains(myInput.ToLower()) ||
								 p.TransType.ToLower().Contains(myInput.ToLower()) ||
								 p.ATMTrans.ToLower().Contains(myInput.ToLower()) ||
								 p.Location.ToLower().Contains(myInput.ToLower()) ||
								 p.ATMUsed.ToLower().Contains(myInput.ToLower()) ||
								 p.PaymentOf.ToLower().Contains(myInput.ToLower()) ||
								 p.TerminalUsed.ToLower().Contains(myInput.ToLower()) ||
								 p.BillerName.ToLower().Contains(myInput.ToLower()) ||
								 p.Merchant.ToLower().Contains(myInput.ToLower()) ||
								 p.Inter.ToLower().Contains(myInput.ToLower()) ||
								 p.BancnetUsed.ToLower().Contains(myInput.ToLower()) ||
								 p.Online.ToLower().Contains(myInput.ToLower()) ||
								 p.Website.ToLower().Contains(myInput.ToLower()) ||
								 p.RemitFrom.ToLower().Contains(myInput.ToLower()) ||
								 p.RemitConcern.ToLower().Contains(myInput.ToLower()) ||
								 p.Amount.ToString().Contains(myInput) ||
								 p.DateTrans.ToString().Contains(myInput) ||
								 p.PLStatus.ToLower().Contains(myInput.ToLower()) ||
								 p.CardBlocked.ToLower().Contains(myInput.ToLower()) ||
								 p.EndoredTo.ToLower().Contains(myInput.ToLower()) ||
								 p.Activity.ToLower().Contains(myInput.ToLower()) ||
								 p.Status.ToLower().Contains(myInput.ToLower()) ||
								 p.SubStatus.ToLower().Contains(myInput.ToLower()) ||
								 p.Agent.ToLower().Contains(myInput.ToLower()) ||
								 p.Resolved.ToString().Contains(myInput) ||
								 p.Remarks.ToLower().Contains(myInput.ToLower()) ||
								 p.ReferedBy.ToLower().Contains(myInput.ToLower()) ||
								 p.ContactPerson.ToLower().Contains(myInput.ToLower()) ||
								 p.abbreviation.ToLower().Contains(myInput.ToLower()) ||
								 p.Tagging.ToLower().Contains(myInput.ToLower()) ||
								 p.ResolvedDate.ToString().Contains(myInput) ||
								 p.CardPresent.ToLower().Contains(myInput.ToLower())
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "TICKET NO";
				ws.Cells["B1"].Value = "INBOUND/OUTBOUND";
				ws.Cells["C1"].Value = "COMM CHANNEL";
				ws.Cells["D1"].Value = "DATE RECEIVED";
				ws.Cells["E1"].Value = "TIME RECEIVED";
				ws.Cells["F1"].Value = "REASON CODE";
				ws.Cells["G1"].Value = "CUSTOMER NAME";
				ws.Cells["H1"].Value = "CARD NO";
				ws.Cells["I1"].Value = "DESTINATION";
				ws.Cells["J1"].Value = "TRANS TYPE";
				ws.Cells["K1"].Value = "ATM TRANS";
				ws.Cells["L1"].Value = "LOCATION";
				ws.Cells["M1"].Value = "ATM USED";
				ws.Cells["N1"].Value = "PAYMENT OF";
				ws.Cells["O1"].Value = "TERMINAL USED";
				ws.Cells["P1"].Value = "BILLERNAME";
				ws.Cells["Q1"].Value = "MERCHANT";
				ws.Cells["R1"].Value = "INTER";
				ws.Cells["S1"].Value = "BANCNET USED";
				ws.Cells["T1"].Value = "ONLINE";
				ws.Cells["U1"].Value = "WEB SITE";
				ws.Cells["V1"].Value = "REMIT FROM";
				ws.Cells["W1"].Value = "REMIT CONCERN";
				ws.Cells["X1"].Value = "AMOUNT";
				ws.Cells["Y1"].Value = "DATE TRANS";
				ws.Cells["Z1"].Value = "PL STATUS";
				ws.Cells["AA1"].Value = "CARD BLOCKED";
				ws.Cells["AB1"].Value = "ACTIVITY";
				ws.Cells["AC1"].Value = "STATUS";
				ws.Cells["AD1"].Value = "SUB STATUS";
				ws.Cells["AE1"].Value = "ENDORSED TO";
				ws.Cells["AF1"].Value = "DUE DATE";
				ws.Cells["AG1"].Value = "RESOLVED DATE";
				ws.Cells["AH1"].Value = "NATURE OF COMPLAINT";
				ws.Cells["AI1"].Value = "REFERRED BY";
				ws.Cells["AJ1"].Value = "CONTACT PERSON";
				ws.Cells["AK1"].Value = "CLASSIFICATION";
				ws.Cells["AL1"].Value = "CARD PRESENT";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.TicketNo;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.InOutBound;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.CommChannel;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.DateReceived.ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.DateReceived.ToString("HH:mm:ss");
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.abbreviation;
					ws.Cells[string.Format("G{0}", rowStart)].Value = item.CustomerName;
					ws.Cells[string.Format("H{0}", rowStart)].Value = item.CardNo;
					ws.Cells[string.Format("I{0}", rowStart)].Value = item.Destination;
					ws.Cells[string.Format("J{0}", rowStart)].Value = item.TransType;
					ws.Cells[string.Format("K{0}", rowStart)].Value = item.ATMTrans;
					ws.Cells[string.Format("L{0}", rowStart)].Value = item.Location;
					ws.Cells[string.Format("M{0}", rowStart)].Value = item.ATMUsed;
					ws.Cells[string.Format("N{0}", rowStart)].Value = item.PaymentOf;
					ws.Cells[string.Format("O{0}", rowStart)].Value = item.TerminalUsed;
					ws.Cells[string.Format("P{0}", rowStart)].Value = item.BillerName;
					ws.Cells[string.Format("Q{0}", rowStart)].Value = item.Merchant;
					ws.Cells[string.Format("R{0}", rowStart)].Value = item.Inter;
					ws.Cells[string.Format("S{0}", rowStart)].Value = item.BancnetUsed;
					ws.Cells[string.Format("T{0}", rowStart)].Value = item.Online;
					ws.Cells[string.Format("U{0}", rowStart)].Value = item.Website;
					ws.Cells[string.Format("V{0}", rowStart)].Value = item.RemitFrom;
					ws.Cells[string.Format("W{0}", rowStart)].Value = item.RemitConcern;
					ws.Cells[string.Format("X{0}", rowStart)].Value = item.Amount;
					ws.Cells[string.Format("Y{0}", rowStart)].Value = (item.DateTrans == null) ? "" : Convert.ToDateTime(item.DateTrans).ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("Z{0}", rowStart)].Value = item.PLStatus;
					ws.Cells[string.Format("AA{0}", rowStart)].Value = item.CardBlocked;
					ws.Cells[string.Format("AB{0}", rowStart)].Value = item.Activity;
					ws.Cells[string.Format("AC{0}", rowStart)].Value = item.Status;
					ws.Cells[string.Format("AD{0}", rowStart)].Value = item.SubStatus;
					ws.Cells[string.Format("AE{0}", rowStart)].Value = item.EndoredTo;
					ws.Cells[string.Format("AF{0}", rowStart)].Value = (item.Resolved == null) ? "" : Convert.ToDateTime(item.Resolved).ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("AG{0}", rowStart)].Value = (item.ResolvedDate == null) ? "" : Convert.ToDateTime(item.ResolvedDate).ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("AH{0}", rowStart)].Value = item.Remarks;
					ws.Cells[string.Format("AI{0}", rowStart)].Value = item.ReferedBy;
					ws.Cells[string.Format("AJ{0}", rowStart)].Value = item.ContactPerson;
					ws.Cells[string.Format("AK{0}", rowStart)].Value = item.Tagging;
					ws.Cells[string.Format("AL{0}", rowStart)].Value = item.CardPresent;

					rowStart++;
				}

				ws.Cells["A:AL"].AutoFitColumns();
				ws.Cells["X:X"].Style.Numberformat.Format = "#,##0.00";

				package.Save();
			}

			fnam = "Resolved - ATM Activity " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		[HttpGet]
		public IActionResult UnresolvedActivity()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Activity", "UnresolvedActivity");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "ATM";
				ViewBag.Page = "ATMOps";

				HttpContext.Session.SetString("deyt1", DateTime.Now.ToString("dd-MMM-yyyy"));
				HttpContext.Session.SetString("deyt2", DateTime.Now.ToString("dd-MMM-yyyy"));

				var vw = _db.ListATMOpsUnResolved(DateTime.Now.ToString("dd-MMM-yyyy"), DateTime.Now.ToString("dd-MMM-yyyy"));

				ViewBag.ListATMOpsUnresolved = vw;

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("ATM Ops - UnResolved Activity", ses.WrkStn, "View List of ATM Ops UnResolved Activity");


				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}

		}

		[HttpPost]
		public IActionResult UnresolvedActivity(DateTime dt1, DateTime dt2)
		{
			try
			{
				ViewBag.Interface = "ATM";
				ViewBag.Page = "ATMOps";

				HttpContext.Session.SetString("deyt1", dt1.ToString("dd-MMM-yyyy"));
				HttpContext.Session.SetString("deyt2", dt2.ToString("dd-MMM-yyyy"));

				var vw = _db.ListATMOpsUnResolved(dt1.ToString("dd-MMM-yyyy"), dt2.ToString("dd-MMM-yyyy"));

				ViewBag.ListATMOpsUnresolved = vw;

				return View(vw);
			}
			catch (Exception ex)
			{
				string controllerName = ControllerContext.ActionDescriptor.ControllerName;
				string actionName = ControllerContext.ActionDescriptor.ActionName;

				string err = "Error in " + controllerName + " - " + actionName + " : " + ex.ToString();

				TempData["msgErr"] = err;
			}

			return View();
		}

		public IActionResult ExportATMUnResolved()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			//var card = HttpContext.Session.GetString("cardno");

			string fnam = string.Empty;

			var v = _db.ListATMOpsUnResolved(dt1, dt2);

			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("ATMOps", "Activity");
			}


			using (ExcelPackage package = new(stream))
			{

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.TicketNo.ToLower().Contains(myInput.ToLower()) ||
								 p.InOutBound.ToLower().Contains(myInput.ToLower()) ||
								 p.CommChannel.ToLower().Contains(myInput.ToLower()) ||
								 p.DateReceived.ToString().Contains(myInput) ||
								 p.abbreviation.ToLower().Contains(myInput.ToLower()) ||
								 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
								 p.CardNo.ToLower().Contains(myInput.ToLower()) ||
								 p.Destination.ToLower().Contains(myInput.ToLower()) ||
								 p.TransType.ToLower().Contains(myInput.ToLower()) ||
								 p.ATMTrans.ToLower().Contains(myInput.ToLower()) ||
								 p.Location.ToLower().Contains(myInput.ToLower()) ||
								 p.ATMUsed.ToLower().Contains(myInput.ToLower()) ||
								 p.PaymentOf.ToLower().Contains(myInput.ToLower()) ||
								 p.TerminalUsed.ToLower().Contains(myInput.ToLower()) ||
								 p.BillerName.ToLower().Contains(myInput.ToLower()) ||
								 p.Merchant.ToLower().Contains(myInput.ToLower()) ||
								 p.Inter.ToLower().Contains(myInput.ToLower()) ||
								 p.BancnetUsed.ToLower().Contains(myInput.ToLower()) ||
								 p.Online.ToLower().Contains(myInput.ToLower()) ||
								 p.Website.ToLower().Contains(myInput.ToLower()) ||
								 p.RemitFrom.ToLower().Contains(myInput.ToLower()) ||
								 p.RemitConcern.ToLower().Contains(myInput.ToLower()) ||
								 p.Amount.ToString().Contains(myInput) ||
								 p.DateTrans.ToString().Contains(myInput) ||
								 p.PLStatus.ToLower().Contains(myInput.ToLower()) ||
								 p.CardBlocked.ToLower().Contains(myInput.ToLower()) ||
								 p.EndoredTo.ToLower().Contains(myInput.ToLower()) ||
								 p.Activity.ToLower().Contains(myInput.ToLower()) ||
								 p.Status.ToLower().Contains(myInput.ToLower()) ||
								 p.SubStatus.ToLower().Contains(myInput.ToLower()) ||
								 p.Agent.ToLower().Contains(myInput.ToLower()) ||
								 p.Resolved.ToString().Contains(myInput) ||
								 p.Remarks.ToLower().Contains(myInput.ToLower()) ||
								 p.ReferedBy.ToLower().Contains(myInput.ToLower()) ||
								 p.ContactPerson.ToLower().Contains(myInput.ToLower()) ||
								 p.abbreviation.ToLower().Contains(myInput.ToLower()) ||
								 p.Tagging.ToLower().Contains(myInput.ToLower()) ||
								 p.ResolvedDate.ToString().Contains(myInput) ||
								 p.CardPresent.ToLower().Contains(myInput.ToLower())
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "TICKET NO";
				ws.Cells["B1"].Value = "INBOUND/OUTBOUND";
				ws.Cells["C1"].Value = "COMM CHANNEL";
				ws.Cells["D1"].Value = "DATE RECEIVED";
				ws.Cells["E1"].Value = "TIME RECEIVED";
				ws.Cells["F1"].Value = "REASON CODE";
				ws.Cells["G1"].Value = "CUSTOMER NAME";
				ws.Cells["H1"].Value = "CARD NO";
				ws.Cells["I1"].Value = "DESTINATION";
				ws.Cells["J1"].Value = "TRANS TYPE";
				ws.Cells["K1"].Value = "ATM TRANS";
				ws.Cells["L1"].Value = "LOCATION";
				ws.Cells["M1"].Value = "ATM USED";
				ws.Cells["N1"].Value = "PAYMENT OF";
				ws.Cells["O1"].Value = "TERMINAL USED";
				ws.Cells["P1"].Value = "BILLERNAME";
				ws.Cells["Q1"].Value = "MERCHANT";
				ws.Cells["R1"].Value = "INTER";
				ws.Cells["S1"].Value = "BANCNET USED";
				ws.Cells["T1"].Value = "ONLINE";
				ws.Cells["U1"].Value = "WEB SITE";
				ws.Cells["V1"].Value = "REMIT FROM";
				ws.Cells["W1"].Value = "REMIT CONCERN";
				ws.Cells["X1"].Value = "AMOUNT";
				ws.Cells["Y1"].Value = "DATE TRANS";
				ws.Cells["Z1"].Value = "PL STATUS";
				ws.Cells["AA1"].Value = "CARD BLOCKED";
				ws.Cells["AB1"].Value = "ACTIVITY";
				ws.Cells["AC1"].Value = "STATUS";
				ws.Cells["AD1"].Value = "SUB STATUS";
				ws.Cells["AE1"].Value = "ENDORSED TO";
				ws.Cells["AF1"].Value = "DUE DATE";
				ws.Cells["AG1"].Value = "RESOLVED DATE";
				ws.Cells["AH1"].Value = "NATURE OF COMPLAINT";
				ws.Cells["AI1"].Value = "REFERRED BY";
				ws.Cells["AJ1"].Value = "CONTACT PERSON";
				ws.Cells["AK1"].Value = "CLASSIFICATION";
				ws.Cells["AL1"].Value = "CARD PRESENT";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.TicketNo;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.InOutBound;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.CommChannel;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.DateReceived.ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.DateReceived.ToString("HH:mm:ss");
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.abbreviation;
					ws.Cells[string.Format("G{0}", rowStart)].Value = item.CustomerName;
					ws.Cells[string.Format("H{0}", rowStart)].Value = item.CardNo;
					ws.Cells[string.Format("I{0}", rowStart)].Value = item.Destination;
					ws.Cells[string.Format("J{0}", rowStart)].Value = item.TransType;
					ws.Cells[string.Format("K{0}", rowStart)].Value = item.ATMTrans;
					ws.Cells[string.Format("L{0}", rowStart)].Value = item.Location;
					ws.Cells[string.Format("M{0}", rowStart)].Value = item.ATMUsed;
					ws.Cells[string.Format("N{0}", rowStart)].Value = item.PaymentOf;
					ws.Cells[string.Format("O{0}", rowStart)].Value = item.TerminalUsed;
					ws.Cells[string.Format("P{0}", rowStart)].Value = item.BillerName;
					ws.Cells[string.Format("Q{0}", rowStart)].Value = item.Merchant;
					ws.Cells[string.Format("R{0}", rowStart)].Value = item.Inter;
					ws.Cells[string.Format("S{0}", rowStart)].Value = item.BancnetUsed;
					ws.Cells[string.Format("T{0}", rowStart)].Value = item.Online;
					ws.Cells[string.Format("U{0}", rowStart)].Value = item.Website;
					ws.Cells[string.Format("V{0}", rowStart)].Value = item.RemitFrom;
					ws.Cells[string.Format("W{0}", rowStart)].Value = item.RemitConcern;
					ws.Cells[string.Format("X{0}", rowStart)].Value = item.Amount;
					ws.Cells[string.Format("Y{0}", rowStart)].Value = (item.DateTrans == null) ? "" : Convert.ToDateTime(item.DateTrans).ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("Z{0}", rowStart)].Value = item.PLStatus;
					ws.Cells[string.Format("AA{0}", rowStart)].Value = item.CardBlocked;
					ws.Cells[string.Format("AB{0}", rowStart)].Value = item.Activity;
					ws.Cells[string.Format("AC{0}", rowStart)].Value = item.Status;
					ws.Cells[string.Format("AD{0}", rowStart)].Value = item.SubStatus;
					ws.Cells[string.Format("AE{0}", rowStart)].Value = item.EndoredTo;
					ws.Cells[string.Format("AF{0}", rowStart)].Value = (item.Resolved == null) ? "" : Convert.ToDateTime(item.Resolved).ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("AG{0}", rowStart)].Value = (item.ResolvedDate == null) ? "" : Convert.ToDateTime(item.ResolvedDate).ToString("dd-MMM-yyyy");
					ws.Cells[string.Format("AH{0}", rowStart)].Value = item.Remarks;
					ws.Cells[string.Format("AI{0}", rowStart)].Value = item.ReferedBy;
					ws.Cells[string.Format("AJ{0}", rowStart)].Value = item.ContactPerson;
					ws.Cells[string.Format("AK{0}", rowStart)].Value = item.Tagging;
					ws.Cells[string.Format("AL{0}", rowStart)].Value = item.CardPresent;

					rowStart++;
				}

				ws.Cells["A:AL"].AutoFitColumns();
				ws.Cells["X:X"].Style.Numberformat.Format = "#,##0.00";

				package.Save();
			}

			fnam = "Un-Resolved - ATM Activity " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}
	}
}
