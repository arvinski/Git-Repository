using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Extensions.Logging;
using AgentDesktop.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Http;
using System.Globalization;
using AgentDesktop.Services;
using OfficeOpenXml.Style;
using OfficeOpenXml;

namespace AgentDesktop.Controllers
{
	public class ReportController : Controller
	{

		private readonly AgentInterface _db;

		public ReportController(AgentInterface db)
		{
			_db = db;
		}

		[HttpGet]
		public IActionResult CMSummary()
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Report", "CMSummary");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Report";
				ViewBag.Page = "CMSummary";


				ReportSummary v = new ReportSummary();

				v.Date1 = DateTime.Now;
				v.Date2 = DateTime.Now;

				var vw = _db.View_ReportSummary(v.Date1.ToString("dd-MMM-yyyy"), v.Date2.ToString("dd-MMM-yyyy"), "", "", "", "");

				ViewBag.ListReportSummary = vw;

				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");


				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Customer Service - Summary Report", ses.WrkStn, "View Report Summary");

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		[HttpPost]
		public IActionResult CMSummary(ReportSummary rpt)
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				//if (ses.UType != "A")
				//{
				//	TempData["msgAuth"] = "You are not authorized to use this feature";
				//	return RedirectToAction("Index", "Home");
				//}

				ViewBag.Interface = "Report";
				ViewBag.Page = "CMSummary";


				if (rpt.ReportType == "0")
				{
					ViewBag.ChkType = "WALA";
					ModelState.AddModelError(nameof(rpt.ReportType), "Report Type required");
				}
				else
				{
					switch (rpt.ReportType)
					{
						case "1":

							string bond = rpt.CommType == "1" ? "Inbound" : "Outbound";
							string cotyp = rpt.CommType == "0" ? "" : bond;

							LoginPass v = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string agt = rpt.Agent == null ? "" : v.LoginID;

							ViewBag.ListReportSummary = _db.View_ReportSummary(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), cotyp, agt, "C", "C");

							HttpContext.Session.SetString("commtype", cotyp);
							HttpContext.Session.SetString("agent", agt);
							HttpContext.Session.SetString("rtype", "C");

							break;
						case "2":

							string bound = rpt.CommType == "1" ? "Inbound" : "Outbound";
							string comtyp = rpt.CommType == "0" ? "" : bound;

							LoginPass a = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string agnt = rpt.Agent == null ? "" : a.LoginID;

							ViewBag.ListReportSummary = _db.View_ReportSummary(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), comtyp, agnt, "C", "R");

							HttpContext.Session.SetString("commtype", comtyp);
							HttpContext.Session.SetString("agent", agnt);
							HttpContext.Session.SetString("rtype", "R");


							break;

						case "3":

							string bnd = rpt.CommType == "1" ? "Inbound" : "Outbound";
							string ctyp = rpt.CommType == "0" ? "" : bnd;

							LoginPass x = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string ag = rpt.Agent == null ? "" : x.LoginID;

							ViewBag.ListReportSummary = _db.View_ReportSummary(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), ctyp, ag, "C", "S");

							HttpContext.Session.SetString("commtype", ctyp);
							HttpContext.Session.SetString("agent", ag);
							HttpContext.Session.SetString("rtype", "S");

							break;
					}
				}

				HttpContext.Session.SetString("deyt1", rpt.Date1.ToString("dd-MMM-yyyy"));
				HttpContext.Session.SetString("deyt2", rpt.Date2.ToString("dd-MMM-yyyy"));

				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");

				return View(rpt);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		public IActionResult ExportSummary()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			var commtype = HttpContext.Session.GetString("commtype");
			var agent = HttpContext.Session.GetString("agent");
			var reportype = HttpContext.Session.GetString("rtype");

			string fnam = string.Empty;

			var v = _db.View_ReportSummary(dt1, dt2, commtype, agent, "C", reportype);

			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("CMSummary", "Report");
			}


			using (ExcelPackage package = new(stream))
			{

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.descr.ToLower().Contains(myInput.ToLower()) ||
								 p.qty.ToString().Contains(myInput.ToLower()) ||
								 p.suma.ToLower().Contains(myInput.ToLower()) 
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				ws.Row(1).Style.Font.Bold = true;
				ws.Row(1).Style.Font.Size = 14;

				switch (reportype)
				{
					case "C":
						ws.Cells["A1"].Value = "CUSTOMER MANAGEMENT REPORT - SUMMARY COMMUNICATION CHANNEL";
						break;
					case "R":
						ws.Cells["A1"].Value = "CUSTOMER MANAGEMENT REPORT - SUMMARY REASON CODE";
						break;
					case "S":
						ws.Cells["A1"].Value = "CUSTOMER MANAGEMENT REPORT - SUMMARY STATUS";
						break;
				}

				ws.Cells["A2"].Value = "Date Coverage";
				ws.Cells["B2"].Value = ":";
				ws.Cells["C2"].Value = dt1 + " to " + dt2;

				ws.Cells["A3"].Value = "Communication Type";
				ws.Cells["B3"].Value = ":";
				ws.Cells["C3"].Value = commtype == "" ? "Inbound and Outbound" : commtype;

				ws.Cells["A4"].Value = "Agent";
				ws.Cells["B4"].Value = ":";
				ws.Cells["C4"].Value = agent == "" ? "All Agent" : agent;


				ws.Row(6).Height = 20;
				ws.Row(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(6).Style.Font.Bold = true;

				ws.Cells["A6"].Value = "Description";
				ws.Cells["B6"].Value = "Quantity";
				ws.Cells["C6"].Value = "Percentage";

				int rowStart = 7;

				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.descr;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.qty;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.suma;

					rowStart++;
				}

				ws.Cells["A3:C3"].AutoFitColumns();

				package.Save();
			}

			switch (reportype)
			{
				case "C":

					fnam = "Report Summary - CommChannel " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

				case "R":
					fnam = "Report Summary - ReasonCode " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

				case "S":
					fnam = "Report Summary - Status " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;
			}

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		[HttpGet]
		public IActionResult ATMSummary()
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Report", "ATMSummary");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "ATMReport";
				ViewBag.Page = "ATMSummary";


				ReportSummary v = new ReportSummary();

				v.Date1 = DateTime.Now;
				v.Date2 = DateTime.Now;

				var vw = _db.View_ReportSummary(v.Date1.ToString("dd-MMM-yyyy"), v.Date2.ToString("dd-MMM-yyyy"), "", "", "", "");

				ViewBag.ListATMReportSummary = vw;

				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("ATM Operations - Summary Report", ses.WrkStn, "View Report Summary");


				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		[HttpPost]
		public IActionResult ATMSummary(ReportSummary rpt)
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				//var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				//if (ses.UType != "A")
				//{
				//	TempData["msgAuth"] = "You are not authorized to use this feature";
				//	return RedirectToAction("Index", "Home");
				//}

				ViewBag.Interface = "ATMReport";
				ViewBag.Page = "ATMSummary";


				if (rpt.ReportType == "0")
				{
					ViewBag.ChkType = "WALA";
					ModelState.AddModelError(nameof(rpt.ReportType), "Report Type required");
				}
				else
				{
					switch (rpt.ReportType)
					{
						case "1":

							string bond = rpt.CommType == "1" ? "Inbound" : "Outbound";
							string cotyp = rpt.CommType == "0" ? "" : bond;

							LoginPass v = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string agt = rpt.Agent == null ? "" : v.LoginID;

							ViewBag.ListATMReportSummary = _db.View_ReportSummary(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), cotyp, agt, "A", "C");

							HttpContext.Session.SetString("commtype", cotyp);
							HttpContext.Session.SetString("agent", agt);
							HttpContext.Session.SetString("rtype", "C");

							break;
						case "2":

							string bound = rpt.CommType == "1" ? "Inbound" : "Outbound";
							string comtyp = rpt.CommType == "0" ? "" : bound;

							LoginPass a = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string agnt = rpt.Agent == null ? "" : a.LoginID;

							ViewBag.ListATMReportSummary = _db.View_ReportSummary(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), comtyp, agnt, "A", "R");

							HttpContext.Session.SetString("commtype", comtyp);
							HttpContext.Session.SetString("agent", agnt);
							HttpContext.Session.SetString("rtype", "R");


							break;

						case "3":

							string bnd = rpt.CommType == "1" ? "Inbound" : "Outbound";
							string ctyp = rpt.CommType == "0" ? "" : bnd;

							LoginPass x = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string ag = rpt.Agent == null ? "" : x.LoginID;

							ViewBag.ListATMReportSummary = _db.View_ReportSummary(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), ctyp, ag, "A", "S");

							HttpContext.Session.SetString("commtype", ctyp);
							HttpContext.Session.SetString("agent", ag);
							HttpContext.Session.SetString("rtype", "S");

							break;
					}
				}

				HttpContext.Session.SetString("deyt1", rpt.Date1.ToString("dd-MMM-yyyy"));
				HttpContext.Session.SetString("deyt2", rpt.Date2.ToString("dd-MMM-yyyy"));

				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");

				return View(rpt);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}


		public IActionResult ExportATMSummary()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			var commtype = HttpContext.Session.GetString("commtype");
			var agent = HttpContext.Session.GetString("agent");
			var reportype = HttpContext.Session.GetString("rtype");

			string fnam = string.Empty;

			var v = _db.View_ReportSummary(dt1, dt2, commtype, agent, "A", reportype);

			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("CMSummary", "Report");
			}


			using (ExcelPackage package = new(stream))
			{

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.descr.ToLower().Contains(myInput.ToLower()) ||
								 p.qty.ToString().Contains(myInput.ToLower()) ||
								 p.suma.ToLower().Contains(myInput.ToLower())
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				ws.Row(1).Style.Font.Bold = true;
				ws.Row(1).Style.Font.Size = 14;

				switch (reportype)
				{
					case "C":
						ws.Cells["A1"].Value = "ATM OPS REPORT - SUMMARY COMMUNICATION CHANNEL";
						break;
					case "R":
						ws.Cells["A1"].Value = "ATM OPS REPORT - SUMMARY REASON CODE";
						break;
					case "S":
						ws.Cells["A1"].Value = "ATM OPS REPORT - SUMMARY STATUS";
						break;
				}

				ws.Cells["A2"].Value = "Date Coverage";
				ws.Cells["B2"].Value = ":";
				ws.Cells["C2"].Value = dt1 + " to " + dt2;

				ws.Cells["A3"].Value = "Communication Type";
				ws.Cells["B3"].Value = ":";
				ws.Cells["C3"].Value = commtype == "" ? "Inbound and Outbound" : commtype;

				ws.Cells["A4"].Value = "Agent";
				ws.Cells["B4"].Value = ":";
				ws.Cells["C4"].Value = agent == "" ? "All Agent" : agent;


				ws.Row(6).Height = 20;
				ws.Row(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(6).Style.Font.Bold = true;

				ws.Cells["A6"].Value = "Description";
				ws.Cells["B6"].Value = "Quantity";
				ws.Cells["C6"].Value = "Percentage";

				int rowStart = 7;

				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.descr;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.qty;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.suma;

					rowStart++;
				}

				ws.Cells["A3:C3"].AutoFitColumns();

				package.Save();
			}

			switch (reportype)
			{
				case "C":

					fnam = "ATM Ops Report Summary - CommChannel " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

				case "R":
					fnam = "ATM Ops Report Summary - ReasonCode " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

				case "S":
					fnam = "ATM Ops Report Summary - Status " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;
			}

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);
		}

		[HttpGet]
		public IActionResult CMManagement()
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Report", "CMManagement");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Report";
				ViewBag.Page = "CMManagement";

				ViewBag.Load = "0";

				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");


				bool isLog = _db.GetAuditLogs("Customer Service - Management Report", ses.WrkStn, "View Management Report");


				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		[HttpPost]
		public IActionResult CMManagement(ReportManagement rpt)
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				//var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				//if (ses.UType != "A")
				//{
				//	TempData["msgAuth"] = "You are not authorized to use this feature";
				//	return RedirectToAction("Index", "Home");
				//}

				ViewBag.Interface = "Report";
				ViewBag.Page = "CMManagement";


				if (rpt.ReportType == "0")
				{
					ViewBag.ChkType = "WALA";
					ModelState.AddModelError(nameof(rpt.ReportType), "Report Type required");
				}
				else
				{
					switch (rpt.ReportType)
					{
						case "1": //Activity Logs

							ViewBag.Load = "1";

							string bond = rpt.CommType == "1" ? "Inbound" : "Outbound";
							string cotyp = rpt.CommType == "0" ? "" : bond;

							//LoginPass v = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string agt = rpt.Agent == null ? "" : rpt.Agent; //v.LoginID;

							string comchan = rpt.Channel == "0" ? "" : rpt.Channel;

							ViewBag.ListActivityLogs = _db.View_ActivityLogs(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), comchan, cotyp, agt, "C");

							HttpContext.Session.SetString("commchannel", comchan);
							HttpContext.Session.SetString("commtype", cotyp);
							HttpContext.Session.SetString("agent", agt);
							HttpContext.Session.SetString("rtype", "1");

							break;
						case "2": //Summary by Reason Code and Comm Channel

							ViewBag.Load = "2";

							//LoginPass a = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string agnt = rpt.Agent == null ? "" : rpt.Agent;

							ViewBag.ListReasonComm = _db.View_ReasonComm(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), agnt, "C");

							HttpContext.Session.SetString("agent", agnt);
							HttpContext.Session.SetString("rtype", "2");


							break;

						case "3":

							ViewBag.Load = "3";

							//LoginPass x = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string ag = rpt.Agent == null ? "" : rpt.Agent;

							ViewBag.ListReasonStat = _db.View_ReasonStat(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), ag, "C");

							HttpContext.Session.SetString("agent", ag);
							HttpContext.Session.SetString("rtype", "3");

							break;
					}
				}

				HttpContext.Session.SetString("deyt1", rpt.Date1.ToString("dd-MMM-yyyy"));
				HttpContext.Session.SetString("deyt2", rpt.Date2.ToString("dd-MMM-yyyy"));

				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");

				return View(rpt);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		public IActionResult ExportManagement()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			var commtype = HttpContext.Session.GetString("commtype");
			var agent = HttpContext.Session.GetString("agent");
			var reportype = HttpContext.Session.GetString("rtype");
			var commchannel = HttpContext.Session.GetString("commchannel");

			string fnam = string.Empty;

			var v = (dynamic)null;

			switch (reportype)
			{
				case "1":
					v = _db.View_ActivityLogs(dt1, dt2, commchannel, commtype, agent, "C");
					break;
				case "2":
					v = _db.View_ReasonComm(dt1, dt2, agent, "C");
					break;
				case "3":
					v = _db.View_ReasonStat(dt1, dt2, agent, "C");
					break;
			}

			//int cnt = v.Count();

			//if (cnt == 0)
			//{
			//	TempData["msginfo"] = "There are no transactions to export in excel";
			//	return RedirectToAction("CMManagement", "Report");
			//}


			using (ExcelPackage package = new(stream))
			{

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				ws.Row(1).Style.Font.Bold = true;
				ws.Row(1).Style.Font.Size = 14;

				switch (reportype)
				{
					case "1":
						ws.Cells["A1"].Value = "MANAGEMENT REPORT - ACTIVITY LOGS";

						ws.Cells["A2"].Value = "Date Coverage";
						ws.Cells["B2"].Value = ":";
						ws.Cells["C2"].Value = dt1 + " to " + dt2;

						ws.Cells["A3"].Value = "Communication Type";
						ws.Cells["B3"].Value = ":";
						ws.Cells["C3"].Value = commtype == "" ? "Inbound and Outbound" : commtype;

						ws.Cells["A4"].Value = "Agent";
						ws.Cells["B4"].Value = ":";
						ws.Cells["C4"].Value = agent == "" ? "All Agent" : agent;

						ws.Cells["A5"].Value = "Comm Channel";
						ws.Cells["B5"].Value = ":";
						ws.Cells["C5"].Value = commchannel == "" ? "All Comm Channel" : commchannel;

						ws.Row(7).Height = 20;
						ws.Row(7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Row(7).Style.Font.Bold = true;

						ws.Cells["A7"].Value = "TIME";
						ws.Cells["B7"].Value = "MON";
						ws.Cells["C7"].Value = "TUE";
						ws.Cells["D7"].Value = "WED";
						ws.Cells["E7"].Value = "THU";
						ws.Cells["F7"].Value = "FRI";
						ws.Cells["G7"].Value = "SAT";
						ws.Cells["H7"].Value = "SUN";

						int rowStart = 8;

						foreach (var item in v)
						{
							ws.Cells[string.Format("A{0}", rowStart)].Value = item.tyme;
							ws.Cells[string.Format("B{0}", rowStart)].Value = item.Monday;
							ws.Cells[string.Format("C{0}", rowStart)].Value = item.Tuesday;
							ws.Cells[string.Format("D{0}", rowStart)].Value = item.Wednesday;
							ws.Cells[string.Format("E{0}", rowStart)].Value = item.Thursday;
							ws.Cells[string.Format("F{0}", rowStart)].Value = item.Friday;
							ws.Cells[string.Format("G{0}", rowStart)].Value = item.Saturday;
							ws.Cells[string.Format("H{0}", rowStart)].Value = item.Sunday;

							rowStart++;
						}

						ws.Cells["A3:H3"].AutoFitColumns();
						ws.Cells["B:H"].Style.Numberformat.Format = "0";

						break;

					case "2":

						ws.Cells["A1"].Value = "MANAGEMENT REPORT - SUMMARY REPORT BY REASON CODE AND COMMUNICATION CHANNEL";

						ws.Cells["A2"].Value = "Date Coverage";
						ws.Cells["B2"].Value = ":";
						ws.Cells["C2"].Value = dt1 + " to " + dt2;

						ws.Cells["A3"].Value = "Agent";
						ws.Cells["B3"].Value = ":";
						ws.Cells["C3"].Value = agent == "" ? "All Agent" : agent;

						ws.Row(5).Height = 20;
						ws.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Row(5).Style.Font.Bold = true;

						ws.Cells["A5"].Value = "REASON CODE";
						ws.Cells["B5"].Value = "COMMUNICATION CHANNEL";
						ws.Cells["C5"].Value = "INBOUND";
						ws.Cells["D5"].Value = "INBOUND %";
						ws.Cells["E5"].Value = "OUTBOUND";
						ws.Cells["F5"].Value = "OUTBOUND %";
						ws.Cells["G5"].Value = "TOTAL";
						ws.Cells["H5"].Value = "TOTAL %";

						int rowStart2 = 6;

						foreach (var item in v)
						{
							ws.Cells[string.Format("A{0}", rowStart2)].Value = item.ReasonC;
							ws.Cells[string.Format("B{0}", rowStart2)].Value = item.ComChannel;
							ws.Cells[string.Format("C{0}", rowStart2)].Value = item.Inbound;
							ws.Cells[string.Format("D{0}", rowStart2)].Value = item.inper;
							ws.Cells[string.Format("E{0}", rowStart2)].Value = item.outbound;
							ws.Cells[string.Format("F{0}", rowStart2)].Value = item.outper;
							ws.Cells[string.Format("G{0}", rowStart2)].Value = item.total;
							ws.Cells[string.Format("H{0}", rowStart2)].Value = item.totalper;

							rowStart2++;
						}

						ws.Cells["A6:H6"].AutoFitColumns();

						break;
					case "3":

						ws.Cells["A1"].Value = "MANAGEMENT REPORT - SUMMARY REPORT BY REASON CODE AND STATUS";

						ws.Cells["A2"].Value = "Date Coverage";
						ws.Cells["B2"].Value = ":";
						ws.Cells["C2"].Value = dt1 + " to " + dt2;

						ws.Cells["A3"].Value = "Agent";
						ws.Cells["B3"].Value = ":";
						ws.Cells["C3"].Value = agent == "" ? "All Agent" : agent;

						ws.Row(5).Height = 20;
						ws.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Row(5).Style.Font.Bold = true;

						ws.Cells["A5"].Value = "REASON CODE";
						ws.Cells["B5"].Value = "STATUS";
						ws.Cells["C5"].Value = "INBOUND";
						ws.Cells["D5"].Value = "INBOUND %";
						ws.Cells["E5"].Value = "OUTBOUND";
						ws.Cells["F5"].Value = "OUTBOUND %";
						ws.Cells["G5"].Value = "TOTAL";
						ws.Cells["H5"].Value = "TOTAL %";

						int rowStart3 = 6;

						foreach (var item in v)
						{
							ws.Cells[string.Format("A{0}", rowStart3)].Value = item.ReasonC;
							ws.Cells[string.Format("B{0}", rowStart3)].Value = item.Status;
							ws.Cells[string.Format("C{0}", rowStart3)].Value = item.Inbound;
							ws.Cells[string.Format("D{0}", rowStart3)].Value = item.inper;
							ws.Cells[string.Format("E{0}", rowStart3)].Value = item.outbound;
							ws.Cells[string.Format("F{0}", rowStart3)].Value = item.outper;
							ws.Cells[string.Format("G{0}", rowStart3)].Value = item.total;
							ws.Cells[string.Format("H{0}", rowStart3)].Value = item.totalper;

							rowStart3++;
						}

						ws.Cells["A6:H6"].AutoFitColumns();

						break;
				}

				package.Save();
			}

			switch (reportype)
			{
				case "1":

					fnam = "Summary Activity Logs " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

				case "2":
					fnam = "Summary ReasonCode and Comm Channel " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

				case "3":
					fnam = "Summary ReasonCode and Status " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;
			}

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		[HttpGet]
		public IActionResult ATMManagement()
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Report", "ATMManagement");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "ATMReport";
				ViewBag.Page = "ATMManagement";

				ViewBag.Load = "0";

				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");


				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("ATM Ops - Management Report", ses.WrkStn, "View Management Report");


				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		[HttpPost]
		public IActionResult ATMManagement(ReportManagement rpt)
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				//var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				//if (ses.UType != "A")
				//{
				//	TempData["msgAuth"] = "You are not authorized to use this feature";
				//	return RedirectToAction("Index", "Home");
				//}

				ViewBag.Interface = "ATMReport";
				ViewBag.Page = "ATMManagement";


				if (rpt.ReportType == "0")
				{
					ViewBag.ChkType = "WALA";
					ModelState.AddModelError(nameof(rpt.ReportType), "Report Type required");
				}
				else
				{
					switch (rpt.ReportType)
					{
						case "1": //Activity Logs

							ViewBag.Load = "1";

							string bond = rpt.CommType == "1" ? "Inbound" : "Outbound";
							string cotyp = rpt.CommType == "0" ? "" : bond;

							LoginPass v = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string agt = rpt.Agent == null ? "" : v.LoginID;

							string comchan = rpt.Channel == "0" ? "" : rpt.Channel;

							ViewBag.ListActivityLogs = _db.View_ActivityLogs(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), comchan, cotyp, agt, "A");

							HttpContext.Session.SetString("commchannel", comchan);
							HttpContext.Session.SetString("commtype", cotyp);
							HttpContext.Session.SetString("agent", agt);
							HttpContext.Session.SetString("rtype", "1");

							break;
						case "2": //Summary by Reason Code and Comm Channel

							ViewBag.Load = "2";

							LoginPass a = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string agnt = rpt.Agent == null ? "" : a.LoginID;

							ViewBag.ListReasonComm = _db.View_ReasonComm(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), agnt, "A");

							HttpContext.Session.SetString("agent", agnt);
							HttpContext.Session.SetString("rtype", "2");


							break;

						case "3":

							ViewBag.Load = "3";

							LoginPass x = _db.GetAgent(Convert.ToInt16(rpt.Agent));

							string ag = rpt.Agent == null ? "" : x.LoginID;

							ViewBag.ListReasonStat = _db.View_ReasonStat(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), ag, "A");

							HttpContext.Session.SetString("agent", ag);
							HttpContext.Session.SetString("rtype", "3");

							break;
					}
				}

				HttpContext.Session.SetString("deyt1", rpt.Date1.ToString("dd-MMM-yyyy"));
				HttpContext.Session.SetString("deyt2", rpt.Date2.ToString("dd-MMM-yyyy"));

				ViewBag.cmbAgent = new SelectList(_db.ListAgent(), "ID", "Description");

				return View(rpt);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}


		public IActionResult ExportATMManagement()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			var commtype = HttpContext.Session.GetString("commtype");
			var agent = HttpContext.Session.GetString("agent");
			var reportype = HttpContext.Session.GetString("rtype");
			var commchannel = HttpContext.Session.GetString("commchannel");

			string fnam = string.Empty;

			var v = (dynamic)null;

			switch (reportype)
			{
				case "1":
					v = _db.View_ActivityLogs(dt1, dt2, commchannel, commtype, agent, "C");
					break;
				case "2":
					v = _db.View_ReasonComm(dt1, dt2, agent, "C");
					break;
				case "3":
					v = _db.View_ReasonStat(dt1, dt2, agent, "C");
					break;
			}

			//int cnt = v.Count();

			//if (cnt == 0)
			//{
			//	TempData["msginfo"] = "There are no transactions to export in excel";
			//	return RedirectToAction("CMManagement", "Report");
			//}


			using (ExcelPackage package = new(stream))
			{

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				ws.Row(1).Style.Font.Bold = true;
				ws.Row(1).Style.Font.Size = 14;

				switch (reportype)
				{
					case "1":
						ws.Cells["A1"].Value = "ATM MANAGEMENT REPORT - ACTIVITY LOGS";

						ws.Cells["A2"].Value = "Date Coverage";
						ws.Cells["B2"].Value = ":";
						ws.Cells["C2"].Value = dt1 + " to " + dt2;

						ws.Cells["A3"].Value = "Communication Type";
						ws.Cells["B3"].Value = ":";
						ws.Cells["C3"].Value = commtype == "" ? "Inbound and Outbound" : commtype;

						ws.Cells["A4"].Value = "Agent";
						ws.Cells["B4"].Value = ":";
						ws.Cells["C4"].Value = agent == "" ? "All Agent" : agent;

						ws.Cells["A5"].Value = "Comm Channel";
						ws.Cells["B5"].Value = ":";
						ws.Cells["C5"].Value = commchannel == "" ? "All Comm Channel" : commchannel;

						ws.Row(7).Height = 20;
						ws.Row(7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Row(7).Style.Font.Bold = true;

						ws.Cells["A7"].Value = "TIME";
						ws.Cells["B7"].Value = "MON";
						ws.Cells["C7"].Value = "TUE";
						ws.Cells["D7"].Value = "WED";
						ws.Cells["E7"].Value = "THU";
						ws.Cells["F7"].Value = "FRI";
						ws.Cells["G7"].Value = "SAT";
						ws.Cells["H7"].Value = "SUN";

						int rowStart = 8;

						foreach (var item in v)
						{
							ws.Cells[string.Format("A{0}", rowStart)].Value = item.tyme;
							ws.Cells[string.Format("B{0}", rowStart)].Value = item.Monday;
							ws.Cells[string.Format("C{0}", rowStart)].Value = item.Tuesday;
							ws.Cells[string.Format("D{0}", rowStart)].Value = item.Wednesday;
							ws.Cells[string.Format("E{0}", rowStart)].Value = item.Thursday;
							ws.Cells[string.Format("F{0}", rowStart)].Value = item.Friday;
							ws.Cells[string.Format("G{0}", rowStart)].Value = item.Saturday;
							ws.Cells[string.Format("H{0}", rowStart)].Value = item.Sunday;

							rowStart++;
						}

						ws.Cells["A3:H3"].AutoFitColumns();
						ws.Cells["B:H"].Style.Numberformat.Format = "0";

						break;

					case "2":

						ws.Cells["A1"].Value = "ATM MANAGEMENT REPORT - SUMMARY REPORT BY REASON CODE AND COMMUNICATION CHANNEL";

						ws.Cells["A2"].Value = "Date Coverage";
						ws.Cells["B2"].Value = ":";
						ws.Cells["C2"].Value = dt1 + " to " + dt2;

						ws.Cells["A3"].Value = "Agent";
						ws.Cells["B3"].Value = ":";
						ws.Cells["C3"].Value = agent == "" ? "All Agent" : agent;

						ws.Row(5).Height = 20;
						ws.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Row(5).Style.Font.Bold = true;

						ws.Cells["A5"].Value = "REASON CODE";
						ws.Cells["B5"].Value = "COMMUNICATION CHANNEL";
						ws.Cells["C5"].Value = "INBOUND";
						ws.Cells["D5"].Value = "INBOUND %";
						ws.Cells["E5"].Value = "OUTBOUND";
						ws.Cells["F5"].Value = "OUTBOUND %";
						ws.Cells["G5"].Value = "TOTAL";
						ws.Cells["H5"].Value = "TOTAL %";

						int rowStart2 = 6;

						foreach (var item in v)
						{
							ws.Cells[string.Format("A{0}", rowStart2)].Value = item.ReasonC;
							ws.Cells[string.Format("B{0}", rowStart2)].Value = item.ComChannel;
							ws.Cells[string.Format("C{0}", rowStart2)].Value = item.Inbound;
							ws.Cells[string.Format("D{0}", rowStart2)].Value = item.inper;
							ws.Cells[string.Format("E{0}", rowStart2)].Value = item.outbound;
							ws.Cells[string.Format("F{0}", rowStart2)].Value = item.outper;
							ws.Cells[string.Format("G{0}", rowStart2)].Value = item.total;
							ws.Cells[string.Format("H{0}", rowStart2)].Value = item.totalper;

							rowStart2++;
						}

						ws.Cells["A6:H6"].AutoFitColumns();

						break;
					case "3":

						ws.Cells["A1"].Value = "ATM MANAGEMENT REPORT - SUMMARY REPORT BY REASON CODE AND STATUS";

						ws.Cells["A2"].Value = "Date Coverage";
						ws.Cells["B2"].Value = ":";
						ws.Cells["C2"].Value = dt1 + " to " + dt2;

						ws.Cells["A3"].Value = "Agent";
						ws.Cells["B3"].Value = ":";
						ws.Cells["C3"].Value = agent == "" ? "All Agent" : agent;

						ws.Row(5).Height = 20;
						ws.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Row(5).Style.Font.Bold = true;

						ws.Cells["A5"].Value = "REASON CODE";
						ws.Cells["B5"].Value = "STATUS";
						ws.Cells["C5"].Value = "INBOUND";
						ws.Cells["D5"].Value = "INBOUND %";
						ws.Cells["E5"].Value = "OUTBOUND";
						ws.Cells["F5"].Value = "OUTBOUND %";
						ws.Cells["G5"].Value = "TOTAL";
						ws.Cells["H5"].Value = "TOTAL %";

						int rowStart3 = 6;

						foreach (var item in v)
						{
							ws.Cells[string.Format("A{0}", rowStart3)].Value = item.ReasonC;
							ws.Cells[string.Format("B{0}", rowStart3)].Value = item.Status;
							ws.Cells[string.Format("C{0}", rowStart3)].Value = item.Inbound;
							ws.Cells[string.Format("D{0}", rowStart3)].Value = item.inper;
							ws.Cells[string.Format("E{0}", rowStart3)].Value = item.outbound;
							ws.Cells[string.Format("F{0}", rowStart3)].Value = item.outper;
							ws.Cells[string.Format("G{0}", rowStart3)].Value = item.total;
							ws.Cells[string.Format("H{0}", rowStart3)].Value = item.totalper;

							rowStart3++;
						}

						ws.Cells["A6:H6"].AutoFitColumns();

						break;
				}

				package.Save();
			}

			switch (reportype)
			{
				case "1":

					fnam = "Summary Activity Logs ATM Ops " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

				case "2":
					fnam = "Summary ReasonCode and Comm Channel ATM Ops " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

				case "3":
					fnam = "Summary ReasonCode and Status ATM Ops " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;
			}

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		[HttpGet]
		public IActionResult CMAging()
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Report", "CMAging");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Report";
				ViewBag.Page = "CMAging";

				ViewBag.AgingHead = "AGING REPORT - CUSTOMER MANAGEMENT";

				var vw = _db.View_Aging(DateTime.Now.ToString("dd-MMM-yyyy"), DateTime.Now.ToString("dd-MMM-yyyy"), "", "C", "");

				ViewBag.ListAging = vw;

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbBranch = new SelectList(_db.Get_Dropdown3(), "ID", "Description");


				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Customer Service - Aging Report", ses.WrkStn, "View Aging Report");


				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		[HttpPost]
		public IActionResult CMAging(ReportAging rpt)
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				//var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				//if (ses.UType != "A")
				//{
				//	TempData["msgAuth"] = "You are not authorized to use this feature";
				//	return RedirectToAction("Index", "Home");
				//}

				ViewBag.Interface = "Report";
				ViewBag.Page = "CMAging";


				if (rpt.ReportType == "0")
				{
					ModelState.AddModelError(nameof(rpt.ReportType), "Report Type required");
				}
				else
				{
					switch (rpt.ReportType)
					{
						case "1":

							ViewBag.AgingHead = "AGING REPORT - CUSTOMER MANAGEMENT";

							string rcode = rpt.ReasonCode ?? "";
							string branch = rpt.Branch ?? "";

							ViewBag.ListAging = _db.View_Aging(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), rcode, "C", branch);

							HttpContext.Session.SetString("rcode", rcode);
							HttpContext.Session.SetString("branch", branch);
							HttpContext.Session.SetString("aging", "1");

							break;
						case "2":

							ViewBag.AgingHead = "AGING REPORT RESOLVED - CUSTOMER MANAGEMENT";

							string recode = rpt.ReasonCode ?? "";
							
							ViewBag.ListAging = _db.View_AgingResolved(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), recode, "C");

							HttpContext.Session.SetString("rcode", recode);
							HttpContext.Session.SetString("aging", "2");

							break;

					}
				}

				HttpContext.Session.SetString("deyt1", rpt.Date1.ToString("dd-MMM-yyyy"));
				HttpContext.Session.SetString("deyt2", rpt.Date2.ToString("dd-MMM-yyyy"));

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");
				ViewBag.cmbBranch = new SelectList(_db.Get_Dropdown3(), "ID", "Description");

				return View(rpt);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		public IActionResult ExportAging()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			var rcode = HttpContext.Session.GetString("rcode");
			var aging = HttpContext.Session.GetString("aging");
			var branch = HttpContext.Session.GetString("branch");

			string fnam = string.Empty;

			var v = aging == "1" ? _db.View_Aging(dt1, dt2, rcode, "C", branch) : _db.View_AgingResolved(dt1, dt2, rcode, "C");


			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("CMAging", "Report");
			}


			using (ExcelPackage package = new(stream))
			{

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.TicketNo.ToLower().Contains(myInput.ToLower()) ||
								 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
								 p.Activity.ToLower().Contains(myInput.ToLower()) ||
								 p.DateReceived.ToString().ToLower().Contains(myInput.ToLower()) ||
								 p.Running.ToString().Contains(myInput) ||
								 p.Follow.ToString().Contains(myInput) ||
								 p.Agent.ToLower().Contains(myInput.ToLower())
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				ws.Row(1).Style.Font.Bold = true;
				ws.Row(1).Style.Font.Size = 14;

				switch (aging)
				{
					case "1":
						ws.Cells["A1"].Value = "AGING REPORT - CUSTOMER MANAGEMENT";
						break;
					case "2":
						ws.Cells["A1"].Value = "AGING REPORT RESOLVED - CUSTOMER MANAGEMENT";
						break;
				}

				ws.Cells["A2"].Value = "Date Coverage";
				ws.Cells["B2"].Value = ":";
				ws.Cells["C2"].Value = dt1 + " to " + dt2;

				ws.Cells["A3"].Value = "Reason Code";
				ws.Cells["B3"].Value = ":";
				ws.Cells["C3"].Value = rcode == "" ? "All Reason Code" : rcode;

				ws.Row(5).Height = 20;
				ws.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(5).Style.Font.Bold = true;

				ws.Cells["A5"].Value = "TICKET NO";
				ws.Cells["B5"].Value = "CUSTOMER NAME";
				ws.Cells["C5"].Value = "ACTIVITY";
				ws.Cells["D5"].Value = "DATE/TIME RECEIVED";
				ws.Cells["E5"].Value = "RUNNING DAYS";
				ws.Cells["F5"].Value = "FOLLOW UP";
				ws.Cells["G5"].Value = "AGENT";

				int rowStart = 6;

				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.TicketNo;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.CustomerName;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.Activity;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.DateReceived;
					ws.Cells[string.Format("D{0}", rowStart)].Style.Numberformat.Format = "m/d/yyy h:mm AM/PM";
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.Running;
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.Follow;
					ws.Cells[string.Format("G{0}", rowStart)].Value = item.Agent;

					rowStart++;
				}

				ws.Cells["A2:C2"].AutoFitColumns();

				package.Save();
			}

			switch (aging)
			{
				case "1":

					fnam = "Aging Report - Customer Management " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

				case "2":
					fnam = "Aging Report Resolved - Customer Management " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

			}

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}


		public IActionResult ExportAgingDet()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			var rcode = HttpContext.Session.GetString("rcode");
			var branch = HttpContext.Session.GetString("branch");

			string fnam = string.Empty;

			var v = _db.View_AgingDetail(dt1, dt2, rcode, "C", branch);


			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("CMAging", "Report");
			}


			using (ExcelPackage package = new(stream))
			{

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.DateReceived.ToLower().Contains(myInput.ToLower()) ||
								 p.TimeReceived.ToLower().Contains(myInput.ToLower()) ||
								 p.TicketNo.ToLower().Contains(myInput.ToLower()) ||
								 p.DateReceived.ToString().ToLower().Contains(myInput.ToLower()) ||
								 p.DueDate.ToString().Contains(myInput) ||
								 p.CustomerName.ToString().Contains(myInput) ||
								 p.CardNo.ToLower().Contains(myInput.ToLower()) ||
								 p.BranchAccount.ToLower().Contains(myInput.ToLower()) ||
								 p.TypeOfComplaint.ToLower().Contains(myInput.ToLower()) ||
								 p.Merchant.ToLower().Contains(myInput.ToLower()) ||
								 p.ATMUsed.ToLower().Contains(myInput.ToLower()) ||
								 p.Location.ToLower().Contains(myInput.ToLower()) ||
								 p.Amount.ToLower().Contains(myInput.ToLower()) ||
								 p.Currency.ToLower().Contains(myInput.ToLower()) ||
								 p.DateTrans.ToLower().Contains(myInput.ToLower()) ||
								 p.Activity.ToLower().Contains(myInput.ToLower()) ||
								 p.UpdateFromEndorsed.ToLower().Contains(myInput.ToLower()) ||
								 p.Status.ToLower().Contains(myInput.ToLower()) ||
								 p.CardNo.ToLower().Contains(myInput.ToLower()) ||
								 p.EndoredTo.ToLower().Contains(myInput.ToLower()) ||
								 p.ResolvedDate.ToLower().Contains(myInput.ToLower()) ||
								 p.last_update.ToLower().Contains(myInput.ToLower()) ||
								 p.ReferedBy.ToLower().Contains(myInput.ToLower()) ||
								 p.created_by.ToLower().Contains(myInput.ToLower()) ||
								 p.Tagging.ToLower().Contains(myInput.ToLower()) 
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				ws.Row(1).Style.Font.Bold = true;
				ws.Row(1).Style.Font.Size = 14;

				ws.Cells["A1"].Value = "AGING REPORT WITH TAT - CUSTOMER MANAGEMENT";

				ws.Cells["A2"].Value = "Date Coverage";
				ws.Cells["B2"].Value = ":";
				ws.Cells["C2"].Value = dt1 + " to " + dt2;

				ws.Cells["A3"].Value = "Reason Code";
				ws.Cells["B3"].Value = ":";
				ws.Cells["C3"].Value = rcode == "" ? "All Reason Code" : rcode;

				ws.Row(5).Height = 20;
				ws.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(5).Style.Font.Bold = true;

				ws.Cells["A5"].Value = "Date Received";
				ws.Cells["B5"].Value = "Time Received";
				ws.Cells["C5"].Value = "Ticket Number";
				ws.Cells["D5"].Value = "Due Date";
				ws.Cells["E5"].Value = "Customers Name";
				ws.Cells["F5"].Value = "Card / Account Number";
				ws.Cells["G5"].Value = "SBA Client";
				ws.Cells["H5"].Value = "Branch of Account";
				ws.Cells["I5"].Value = "Type of Complaint";
				ws.Cells["J5"].Value = "Merchant";
				ws.Cells["K5"].Value = "ATMUsed";
				ws.Cells["L5"].Value = "Location";
				ws.Cells["M5"].Value = "Amount";
				ws.Cells["N5"].Value = "Currency";
				ws.Cells["O5"].Value = "Transaction Date";
				ws.Cells["P5"].Value = "Activity";
				ws.Cells["Q5"].Value = "Update from the Endorsed Unit";
				ws.Cells["R5"].Value = "Status";
				ws.Cells["S5"].Value = "Endorsed To";
				ws.Cells["T5"].Value = "Date Resolved";
				ws.Cells["U5"].Value = "Last CSR Handler";
				ws.Cells["V5"].Value = "First Contact Department";
				ws.Cells["W5"].Value = "First Person Contact";
				ws.Cells["X5"].Value = "Classification";
				ws.Cells["Y5"].Value = "Aging";
				ws.Cells["Z5"].Value = "Within/Outside TAT";

				int rowStart = 6;

				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.DateReceived;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.TimeReceived;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.TicketNo;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.DueDate;
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.CustomerName;
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.CardNo;
					ws.Cells[string.Format("G{0}", rowStart)].Value = item.BranchAccount;
					ws.Cells[string.Format("H{0}", rowStart)].Value = "";
					ws.Cells[string.Format("I{0}", rowStart)].Value = item.TypeOfComplaint;
					ws.Cells[string.Format("J{0}", rowStart)].Value = item.Merchant;
					ws.Cells[string.Format("K{0}", rowStart)].Value = item.ATMUsed;
					ws.Cells[string.Format("L{0}", rowStart)].Value = item.Location;
					ws.Cells[string.Format("M{0}", rowStart)].Value = item.Amount;
					ws.Cells[string.Format("N{0}", rowStart)].Value = item.Currency;
					ws.Cells[string.Format("O{0}", rowStart)].Value = item.DateTrans;
					ws.Cells[string.Format("P{0}", rowStart)].Value = item.Activity;
					ws.Cells[string.Format("Q{0}", rowStart)].Value = item.UpdateFromEndorsed;
					ws.Cells[string.Format("R{0}", rowStart)].Value = item.Status;
					ws.Cells[string.Format("S{0}", rowStart)].Value = item.EndoredTo;
					ws.Cells[string.Format("T{0}", rowStart)].Value = item.ResolvedDate;
					ws.Cells[string.Format("U{0}", rowStart)].Value = item.last_update;
					ws.Cells[string.Format("V{0}", rowStart)].Value = item.ReferedBy;
					ws.Cells[string.Format("W{0}", rowStart)].Value = item.created_by;
					ws.Cells[string.Format("X{0}", rowStart)].Value = item.Tagging;
					ws.Cells[string.Format("Y{0}", rowStart)].Value = item.aging	;
					ws.Cells[string.Format("Z{0}", rowStart)].Value = item.Within_Outside_TAT;

					rowStart++;
				}

				ws.Cells["A2:C2"].AutoFitColumns();

				package.Save();
			}

			fnam = "Aging Report Detail with TAT - Customer Management " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}


		[HttpGet]
		public IActionResult ATMAging()
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Report", "ATMAging");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "ATMReport";
				ViewBag.Page = "ATMAging";

				ViewBag.AgingHead = "AGING REPORT - ATM OPERATIONS";

				var vw = _db.View_Aging(DateTime.Now.ToString("dd-MMM-yyyy"), DateTime.Now.ToString("dd-MMM-yyyy"), "", "A", "");

				ViewBag.ListAging = vw;

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");


				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("ATM Ops - Aging Report", ses.WrkStn, "View Aging Report");


				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		[HttpPost]
		public IActionResult ATMAging(ReportAging rpt)
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				//var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				//if (ses.UType != "A")
				//{
				//	TempData["msgAuth"] = "You are not authorized to use this feature";
				//	return RedirectToAction("Index", "Home");
				//}

				ViewBag.Interface = "ATMReport";
				ViewBag.Page = "ATMAging";


				if (rpt.ReportType == "0")
				{
					ModelState.AddModelError(nameof(rpt.ReportType), "Report Type required");
				}
				else
				{
					switch (rpt.ReportType)
					{
						case "1":

							ViewBag.AgingHead = "AGING REPORT - ATM OPERATIONS";
							string rcode = rpt.ReasonCode ?? "";
							string branch = rpt.Branch ?? "";

							ViewBag.ListAging = _db.View_Aging(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), rcode, "A", branch);

							HttpContext.Session.SetString("rcode", rcode);
							HttpContext.Session.SetString("aging", "1");

							break;
						case "2":

							ViewBag.AgingHead = "AGING REPORT RESOLVED - ATM OPERATIONS";

							string recode = rpt.ReasonCode ?? "";

							ViewBag.ListAging = _db.View_AgingResolved(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), recode, "A");

							HttpContext.Session.SetString("rcode", recode);
							HttpContext.Session.SetString("aging", "2");

							break;

					}
				}

				HttpContext.Session.SetString("deyt1", rpt.Date1.ToString("dd-MMM-yyyy"));
				HttpContext.Session.SetString("deyt2", rpt.Date2.ToString("dd-MMM-yyyy"));

				ViewBag.cmbReasonCode = new SelectList(_db.ListReasonCode(), "ID", "Description");

				return View(rpt);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		public IActionResult ExportATMAging()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			var rcode = HttpContext.Session.GetString("rcode");
			var aging = HttpContext.Session.GetString("aging");
			var branch = HttpContext.Session.GetString("branch");

			string fnam = string.Empty;

			var v = aging == "1" ? _db.View_Aging(dt1, dt2, rcode, "A", branch) : _db.View_AgingResolved(dt1, dt2, rcode, "A");


			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("CMAging", "Report");
			}


			using (ExcelPackage package = new(stream))
			{

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.TicketNo.ToLower().Contains(myInput.ToLower()) ||
								 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
								 p.Activity.ToLower().Contains(myInput.ToLower()) ||
								 p.DateReceived.ToString().ToLower().Contains(myInput.ToLower()) ||
								 p.Running.ToString().Contains(myInput) ||
								 p.Follow.ToString().Contains(myInput) ||
								 p.Agent.ToLower().Contains(myInput.ToLower())
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				ws.Row(1).Style.Font.Bold = true;
				ws.Row(1).Style.Font.Size = 14;

				switch (aging)
				{
					case "1":
						ws.Cells["A1"].Value = "AGING REPORT - ATM OPERATIONS";
						break;
					case "2":
						ws.Cells["A1"].Value = "AGING REPORT RESOLVED - ATM OPERATIONS";
						break;
				}

				ws.Cells["A2"].Value = "Date Coverage";
				ws.Cells["B2"].Value = ":";
				ws.Cells["C2"].Value = dt1 + " to " + dt2;

				ws.Cells["A3"].Value = "Reason Code";
				ws.Cells["B3"].Value = ":";
				ws.Cells["C3"].Value = rcode == "" ? "All Reason Code" : rcode;

				ws.Row(5).Height = 20;
				ws.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(5).Style.Font.Bold = true;

				ws.Cells["A5"].Value = "TICKET NO";
				ws.Cells["B5"].Value = "CUSTOMER NAME";
				ws.Cells["C5"].Value = "ACTIVITY";
				ws.Cells["D5"].Value = "DATE/TIME RECEIVED";
				ws.Cells["E5"].Value = "RUNNING DAYS";
				ws.Cells["F5"].Value = "FOLLOW UP";
				ws.Cells["G5"].Value = "AGENT";

				int rowStart = 6;

				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.TicketNo;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.CustomerName;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.Activity;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.DateReceived;
					ws.Cells[string.Format("D{0}", rowStart)].Style.Numberformat.Format = "m/d/yyy h:mm AM/PM";
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.Running;
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.Follow;
					ws.Cells[string.Format("G{0}", rowStart)].Value = item.Agent;

					rowStart++;
				}

				ws.Cells["A2:C2"].AutoFitColumns();

				package.Save();
			}

			switch (aging)
			{
				case "1":

					fnam = "Aging Report - ATM Ops " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

				case "2":
					fnam = "Aging Report Resolved - ATM Ops " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
					break;

			}

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		[HttpGet]
		public IActionResult CMReportIRC()
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Report", "CMReportIRC");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Report";
				ViewBag.Page = "CMReportIRC";


				ReportIRCFilter v = new ReportIRCFilter();

				v.Date1 = DateTime.Now;
				v.Date2 = DateTime.Now;

				var vw = _db.List_ReportIRC(v.Date1.ToString("dd-MMM-yyyy"), v.Date2.ToString("dd-MMM-yyyy"), v.ReportType);

				ViewBag.ListReportIRC = vw;

				ViewBag.ViewTableIRC = "0";

				//string rpt = v.ReportTypeDet ?? "0";

				//switch (rpt)
				//{
				//	case "2":
				//		var vwd = _db.List_ReportIRCDetail(v.Date1.ToString("dd-MMM-yyyy"), v.Date2.ToString("dd-MMM-yyyy"), rpt);

				//		ViewBag.ListReportIRCDetail = vwd;

				//		break;
				//	default:

				//		var vw = _db.List_ReportIRCDetail(v.Date1.ToString("dd-MMM-yyyy"), v.Date2.ToString("dd-MMM-yyyy"), rpt);

				//		ViewBag.ListReportIRC = vw;

				//		break;
				//}


				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Customer Service - Summary Report IRC", ses.WrkStn, "View Report by Reason Codes I/R/C");

				return View(v);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		[HttpPost]
		public IActionResult CMReportIRC(ReportIRCFilter rpt)
		{

			if (rpt.ReportType == "0")
			{
				TempData["msginfo"] = "Please choose report type";

				return RedirectToAction("CMReportIRC", "Report");

			}

			if (rpt.ReportTypeDet == "0")
			{
				TempData["msginfo"] = "Please choose report type details";

				return RedirectToAction("CMReportIRC", "Report");

			}

			ViewBag.Interface = "Report";
			ViewBag.Page = "CMReportIRC"; ;

			HttpContext.Session.SetString("deyt1", rpt.Date1.ToString("dd-MMM-yyyy"));
			HttpContext.Session.SetString("deyt2", rpt.Date2.ToString("dd-MMM-yyyy"));
			HttpContext.Session.SetString("irc", rpt.ReportType);
			HttpContext.Session.SetString("det", rpt.ReportTypeDet);


			switch (rpt.ReportTypeDet.ToString())
			{
				case "2":
					var vwd = _db.List_ReportIRCDetail(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), rpt.ReportType);

					ViewBag.ListReportIRCDetail = vwd;

					ViewBag.ViewTableIRC = "2";

					break;
				default:

					var vw = _db.List_ReportIRC(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), rpt.ReportType);

					ViewBag.ListReportIRC = vw;

					ViewBag.ViewTableIRC = "0";

					break;
			}

			//string detl = rpt.ReportTypeDet;

			//var vw = detl == "2" ? _db.List_ReportIRCDetail(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), rpt.ReportType) : _db.List_ReportIRC(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), rpt.ReportType);




			return View(rpt);


		}

		public IActionResult ExportIRC()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			var irc = HttpContext.Session.GetString("irc");
			var det = HttpContext.Session.GetString("det");

			string fnam = string.Empty;

			var v = _db.List_ReportIRC(dt1, dt2, irc);

			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("CMReportIRC", "Report");
			}


			using (ExcelPackage package = new(stream))
			{

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.reason.ToLower().Contains(myInput.ToLower()) ||
								 p.reasoncode.ToLower().Contains(myInput.ToLower()) 
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				ws.Row(1).Style.Font.Bold = true;
				ws.Row(1).Style.Font.Size = 14;

				ws.Cells["A1"].Value = "SUMMARY REPORT BY REASON CODES (INQUIRY/REQUEST/COMPLAINT)";

				ws.Cells["A2"].Value = "Date Coverage";
				ws.Cells["B2"].Value = ":";
				ws.Cells["C2"].Value = dt1 + " to " + dt2;

				ws.Cells["A3"].Value = "Report Type";
				ws.Cells["B3"].Value = ":";

				switch (irc)
				{
					case "I":
						ws.Cells["C3"].Value = "Inquiries";
						break;
					case "R":
						ws.Cells["C3"].Value = "Requests";
						break;
					default:
						ws.Cells["C3"].Value = "Complaints";
						break;

				}

				ws.Row(5).Height = 20;
				ws.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(5).Style.Font.Bold = true;

				ws.Cells["A5"].Value = "DESCRIPTION";
				ws.Cells["B5"].Value = "REASON CODE";
				ws.Cells["C5"].Value = "COUNT";
				ws.Cells["D5"].Value = "COMPLEX";
				ws.Cells["E5"].Value = "SIMPLE";
				ws.Cells["F5"].Value = "RESOLVED";
				ws.Cells["G5"].Value = "OPEN";
				ws.Cells["H5"].Value = "Cancelled";
				ws.Cells["I5"].Value = "Closed";
				ws.Cells["J5"].Value = "Commitment";
				ws.Cells["K5"].Value = "Endorse";
				ws.Cells["L5"].Value = "Resolved";



				int rowStart = 6;

				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.reason;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.reasoncode;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.cnt;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.complex;
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.simple;
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.resolved;
					ws.Cells[string.Format("G{0}", rowStart)].Value = item.open;
					ws.Cells[string.Format("H{0}", rowStart)].Value = item.cancelled;
					ws.Cells[string.Format("I{0}", rowStart)].Value = item.closed;
					ws.Cells[string.Format("J{0}", rowStart)].Value = item.commitment;
					ws.Cells[string.Format("K{0}", rowStart)].Value = item.endorsed;
					ws.Cells[string.Format("L{0}", rowStart)].Value = item.resolved2;

					rowStart++;
				}

				ws.Cells["A2:C2"].AutoFitColumns();

				package.Save();
			}


			fnam = "CM Summary Report - Reason Codes I/R/C " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		[HttpGet]
		public IActionResult ATMReportIRC()
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Report", "ATMReportIRC");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "ATMReport";
				ViewBag.Page = "ATMReportIRC";


				ReportIRCFilter v = new ReportIRCFilter();

				v.Date1 = DateTime.Now;
				v.Date2 = DateTime.Now;

				var vw = _db.List_ReportIRCATM(v.Date1.ToString("dd-MMM-yyyy"), v.Date2.ToString("dd-MMM-yyyy"), v.ReportType);

				ViewBag.ListReportIRC = vw;


				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("ATM Ops - Summary Report IRC", ses.WrkStn, "View Report by Reason Codes I/R/C");

				return View(v);
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		[HttpPost]
		public IActionResult ATMReportIRC(ReportIRCFilter rpt)
		{

			if (rpt.ReportType == "0")
			{
				TempData["msginfo"] = "Please choose report type";

				return RedirectToAction("ATMReportIRC", "Report");

			}

			ViewBag.Interface = "Report";
			ViewBag.Page = "ATMReportIRC"; ;

			HttpContext.Session.SetString("deyt1", rpt.Date1.ToString("dd-MMM-yyyy"));
			HttpContext.Session.SetString("deyt2", rpt.Date2.ToString("dd-MMM-yyyy"));
			HttpContext.Session.SetString("irc", rpt.ReportType);

			var vw = _db.List_ReportIRCATM(rpt.Date1.ToString("dd-MMM-yyyy"), rpt.Date2.ToString("dd-MMM-yyyy"), rpt.ReportType);

			ViewBag.ListReportIRC = vw;


			return View(rpt);


		}

		public IActionResult ExportIRCATM()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			var irc = HttpContext.Session.GetString("irc");

			string fnam = string.Empty;

			var v = _db.List_ReportIRCATM(dt1, dt2, irc);

			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("ATMReportIRC", "Report");
			}


			using (ExcelPackage package = new(stream))
			{

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.reason.ToLower().Contains(myInput.ToLower()) ||
								 p.reasoncode.ToLower().Contains(myInput.ToLower())
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				ws.Row(1).Style.Font.Bold = true;
				ws.Row(1).Style.Font.Size = 14;

				ws.Cells["A1"].Value = "SUMMARY REPORT BY REASON CODES (INQUIRY/REQUEST/COMPLAINT)";

				ws.Cells["A2"].Value = "Date Coverage";
				ws.Cells["B2"].Value = ":";
				ws.Cells["C2"].Value = dt1 + " to " + dt2;

				ws.Cells["A3"].Value = "Report Type";
				ws.Cells["B3"].Value = ":";

				switch (irc)
				{
					case "I":
						ws.Cells["C3"].Value = "Inquiries";
						break;
					case "R":
						ws.Cells["C3"].Value = "Requests";
						break;
					default:
						ws.Cells["C3"].Value = "Complaints";
						break;

				}

				ws.Row(5).Height = 20;
				ws.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(5).Style.Font.Bold = true;

				ws.Cells["A5"].Value = "DESCRIPTION";
				ws.Cells["B5"].Value = "REASON CODE";
				ws.Cells["C5"].Value = "COUNT";
				ws.Cells["D5"].Value = "COMPLEX";
				ws.Cells["E5"].Value = "SIMPLE";
				ws.Cells["F5"].Value = "RESOLVED";
				ws.Cells["G5"].Value = "OPEN";
				ws.Cells["H5"].Value = "Cancelled";
				ws.Cells["I5"].Value = "Closed";
				ws.Cells["J5"].Value = "Commitment";
				ws.Cells["K5"].Value = "Endorse";
				ws.Cells["L5"].Value = "Resolved";



				int rowStart = 6;

				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.reason;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.reasoncode;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.cnt;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.complex;
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.simple;
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.resolved;
					ws.Cells[string.Format("G{0}", rowStart)].Value = item.open;
					ws.Cells[string.Format("H{0}", rowStart)].Value = item.cancelled;
					ws.Cells[string.Format("I{0}", rowStart)].Value = item.closed;
					ws.Cells[string.Format("J{0}", rowStart)].Value = item.commitment;
					ws.Cells[string.Format("K{0}", rowStart)].Value = item.endorsed;
					ws.Cells[string.Format("L{0}", rowStart)].Value = item.resolved2;

					rowStart++;
				}

				ws.Cells["A2:C2"].AutoFitColumns();

				package.Save();
			}


			fnam = "ATM Ops Summary Report - Reason Codes I/R/C " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		public JsonResult GetReportType(string myInput)
		{

			//HttpContext.Session.SetString("getdet", myInput ?? "");

			return Json(new { getdet = myInput });
		}

		public IActionResult ExportIRCDetail()
		{
			var stream = new System.IO.MemoryStream();

			var dt1 = HttpContext.Session.GetString("deyt1");
			var dt2 = HttpContext.Session.GetString("deyt2");
			var irc = HttpContext.Session.GetString("irc");

			string fnam = string.Empty;

			var v = _db.List_ReportIRCDetail(dt1, dt2, irc);

			int cnt = v.Count();

			if (cnt == 0)
			{
				TempData["msginfo"] = "There are no transactions to export in excel";
				return RedirectToAction("CMReportIRC", "Report");
			}


			using (ExcelPackage package = new(stream))
			{


				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				ws.Row(1).Style.Font.Bold = true;
				ws.Row(1).Style.Font.Size = 14;

				ws.Cells["A1"].Value = "SUMMARY REPORT BY REASON CODES (INQUIRY/REQUEST/COMPLAINT) - DETAIL";

				ws.Cells["A2"].Value = "Date Coverage";
				ws.Cells["B2"].Value = ":";
				ws.Cells["C2"].Value = dt1 + " to " + dt2;

				ws.Cells["A3"].Value = "Report Type";
				ws.Cells["B3"].Value = ":";

				switch (irc)
				{
					case "I":
						ws.Cells["C3"].Value = "Inquiries";
						break;
					case "R":
						ws.Cells["C3"].Value = "Requests";
						break;
					default:
						ws.Cells["C3"].Value = "Complaints";
						break;

				}

				ws.Row(5).Height = 20;
				ws.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(5).Style.Font.Bold = true;

				ws.Cells["A5"].Value = "TICKET NO";
				ws.Cells["B5"].Value = "INBOUND/OUTBOUND";
				ws.Cells["C5"].Value = "COMM CHANNEL";
				ws.Cells["D5"].Value = "DATE RECEIVED";
				ws.Cells["E5"].Value = "TIME RECEIVED";
				ws.Cells["F5"].Value = "REASON CODE";
				ws.Cells["G5"].Value = "CUSTOMER NAME";
				ws.Cells["H5"].Value = "CARD NO";
				ws.Cells["I5"].Value = "DESTINATION";
				ws.Cells["J5"].Value = "TRANS TYPE";
				ws.Cells["K5"].Value = "ATM TRANS";
				ws.Cells["L5"].Value = "LOCATION";
				ws.Cells["M5"].Value = "ATM USED";
				ws.Cells["N5"].Value = "PAYMENT OF";
				ws.Cells["O5"].Value = "TERMINAL USED";
				ws.Cells["P5"].Value = "BILLERNAME";
				ws.Cells["Q5"].Value = "MERCHANT";
				ws.Cells["R5"].Value = "INTER";
				ws.Cells["S5"].Value = "BANCNET USED";
				ws.Cells["T5"].Value = "ONLINE";
				ws.Cells["U5"].Value = "WEB SITE";
				ws.Cells["V5"].Value = "REMIT FROM";
				ws.Cells["W5"].Value = "REMIT CONCERN";
				ws.Cells["X5"].Value = "AMOUNT";
				ws.Cells["Y5"].Value = "CURRENCY";
				ws.Cells["Z5"].Value = "DATE TRANS";
				ws.Cells["AA5"].Value = "PL STATUS";
				ws.Cells["AB5"].Value = "CARD BLOCKED";
				ws.Cells["AC5"].Value = "ACTIVITY";
				ws.Cells["AD5"].Value = "STATUS";
				ws.Cells["AE5"].Value = "SUB STATUS";
				ws.Cells["AF5"].Value = "CREATED BY";
				ws.Cells["AG5"].Value = "LAST UPDATED BY";
				ws.Cells["AH5"].Value = "ENDORSED TO";
				ws.Cells["AI5"].Value = "ENDORSED FROM";
				ws.Cells["AJ5"].Value = "DUE DATE";
				ws.Cells["AK5"].Value = "RESOLVED DATE";
				ws.Cells["AL5"].Value = "NATURE OF COMPLAINT";
				ws.Cells["AM5"].Value = "REFERRED BY";
				ws.Cells["AN5"].Value = "CONTACT PERSON";
				ws.Cells["AO5"].Value = "CLASSIFICATION";
				ws.Cells["AP5"].Value = "TIME OF CALL";
				ws.Cells["AQ5"].Value = "LOCAL NUMBER";
				ws.Cells["AR5"].Value = "CARD PRESENT";
				ws.Cells["AS5"].Value = "BRANCH";


				int rowStart = 6;

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

				ws.Cells["A2:C2"].AutoFitColumns();

				package.Save();
			}


			fnam = "CM Summary Report - Reason Codes I/R/C per detail " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";

			string fileName = fnam;
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}


	}
}
