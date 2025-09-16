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
	public class HomeController : Controller
	{
		private readonly ILogger<HomeController> _logger;
		private readonly AgentInterface _db;
#nullable enable
		private string? TranDate { get; set; }
#nullable disable

		public HomeController(ILogger<HomeController> logger, AgentInterface db)
		{
			_logger = logger;
			_db = db;
		}

		[HttpGet]
		public IActionResult Index()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				ViewBag.FullName = ses.FullName;

				ViewBag.Dept = ses.Dept.ToString();

				if (ses.Stat == "F")
				{
					TempData["msgpass"] = "Please change your password immediately";
				}

				if ((ses.Dept == "M" || ses.Dept == "A"))
				{
					ViewBag.IdentifyMe = "Meron";

					var v = (ses.Dept == "M") ? _db.DisplayServiceTickets(ses.LoginID, ses.Dept, ses.UType, 3) : _db.DisplayServiceTickets(ses.LoginID, ses.Dept, ses.UType, 4);

					int cnt = v.Count();

					ViewBag.ListReminder = v; /*(cnt == 0) ? "Wala" : "Meron";*/

					if (cnt != 0)
					{
						if (ses.Dept == "M" || ses.Dept == "A")
						{
							TempData["msgReminder"] = "meron";
						}
					}

					bool isDashboard = _db.Populate_Dashboard(ses.LoginID.ToString());

					//if (isDashboard)
					//{
					//SERVICE TICKET 
					var ser = _db.ListServiceTicket(ses.LoginID, ses.Dept);

					ViewBag.ListServiceTicket = ser;

					//STATUS
					var st = _db.ListServiceStatus(ses.LoginID, ses.Dept);

					ViewBag.ListServiceStatus = st;

					//ESCALATION
					var esc = _db.ListEscalation();
					ViewBag.ListEscalations = esc;

					//REASON CODES
					var rcode = _db.ListServiceReason(ses.LoginID, ses.Dept);

					ViewBag.ListReasonCode = rcode;

					//CHANNEL
					var chanl = _db.ListServiceChannel(ses.LoginID, ses.Dept);

					ViewBag.ListComChannel = chanl;
				}
				else
				{
					ViewBag.IdentifyMe = "Wala";

				}


				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}

		}

		public IActionResult DisplayTicket(string id, string service, string cnt)
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				ViewBag.Dept = ses.Dept.ToString();

				HttpContext.Session.SetString("ayd", id ?? "");
				HttpContext.Session.SetString("serbis", service ?? "");

				string _id = Security.DeCryptor(id);

				switch (_id)
				{
					case "st":
						switch (service)
						{
							case "TOTAL DUE TODAY":
								if (cnt == "0")
								{
									TempData["msginfo"] = "There are no transactions to display";

									return RedirectToAction("Index", "Home");
								}
								else
								{
									var disp = _db.DisplayServiceTickets(ses.LoginID, ses.Dept, ses.UType, 1);

									ViewBag.DisplayTickets = disp;

									ViewBag.DiplayTitle = "SERVICE TICKETS DUE TODAY";

									//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

									bool isLog = _db.GetAuditLogs("Dashboard", ses.WrkStn, "VIEW SERVICE TICKETS DUE TODAY");

								}


								break;
							case "TOTAL DUE WITHIN THE NEXT 7 DAYS":
								if (cnt == "0")
								{
									TempData["msginfo"] = "There are no transactions to display";

									return RedirectToAction("Index", "Home");
								}
								else
								{
									var disp = _db.DisplayServiceTickets(ses.LoginID, ses.Dept, ses.UType, 2);

									ViewBag.DisplayTickets = disp;

									ViewBag.DiplayTitle = "SERVICE TICKETS DUE WITHIN THE NEXT 7 DAYS";

									//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

									bool isLog = _db.GetAuditLogs("Dashboard", ses.WrkStn, "VIEW SERVICE TICKETS DUE WITHIN THE NEXT 7 DAYS");
								}
								break;
							case "TOTAL PAST DUE":
								if (cnt == "0")
								{
									TempData["msginfo"] = "There are no transactions to display";

									return RedirectToAction("Index", "Home");
								}
								else
								{
									var disp = _db.DisplayServiceTickets(ses.LoginID, ses.Dept, ses.UType, 3);

									ViewBag.DisplayTickets = disp;

									ViewBag.DiplayTitle = "SERVICE TICKETS PAST DUE";

									//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

									bool isLog = _db.GetAuditLogs("Dashboard", ses.WrkStn, "VIEW SERVICE TICKETS PAST DUE");

								}
								break;

						}


						break;
					case "ec":
						switch (service)
						{
							case "TOTAL ESCALATION DUE FOR 3 DAYS":
								if (cnt == "0")
								{
									TempData["msginfo"] = "There are no transactions to display";

									return RedirectToAction("Index", "Home");
								}
								else
								{
									var disp = _db.View_TransEscalate(1);

									ViewBag.DisplayTickets = disp;

									ViewBag.DiplayTitle = "TOTAL ESCALATION DUE FOR 3 DAYS";

									//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

									bool isLog = _db.GetAuditLogs("Dashboard", ses.WrkStn, "VIEW TOTAL ESCALATION DUE FOR 3 DAYS");
								}


								break;
							case "TOTAL ESCALATION DUE FOR 7 DAYS":
								if (cnt == "0")
								{
									TempData["msginfo"] = "There are no transactions to display";

									return RedirectToAction("Index", "Home");
								}
								else
								{
									var disp = _db.View_TransEscalate(2);

									ViewBag.DisplayTickets = disp;

									ViewBag.DiplayTitle = "TOTAL ESCALATION DUE FOR 7 DAYS";

									//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

									bool isLog = _db.GetAuditLogs("Dashboard", ses.WrkStn, "VIEW TOTAL ESCALATION DUE FOR 7 DAYS");

								}
								break;

						}

						break;
				}

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}

			
		}

		[HttpGet]
		public IActionResult DisplayTicketDated(string id, string service, string cnt, string dt1, string dt2)
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				ViewBag.Dept = ses.Dept.ToString();

				var v1 = (dt1 == "XXX") ? TranDate : dt1;
				var v2 = (dt2 == "XXX") ? TranDate : dt2;

				HttpContext.Session.SetString("ayd", id ?? "");
				HttpContext.Session.SetString("serbis", service ?? "");
				HttpContext.Session.SetString("kawnt", cnt ?? "");
				HttpContext.Session.SetString("dt1", dt1 ?? "");
				HttpContext.Session.SetString("dt2", dt2 ?? "");

				string _id = Security.DeCryptor(id);

				switch (_id)
				{
					case "tt":

						if (cnt == "0")
						{
							TempData["msginfo"] = "There are no transactions to display";

							return RedirectToAction("Index", "Home");
						}
						else
						{
							var stat = _db.View_TransStatus(ses.LoginID, service, 0, v1, v2);

							ViewBag.DisplayTickets = stat;

							ViewBag.DiplayTitle = "VIEW TRANSACTION VIA " + service.ToUpper() + " STATUS";

							string wrkstn1 = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog1 = _db.GetAuditLogs("Dashboard", wrkstn1, "VIEW TRANSACTION VIA " + service.ToUpper() + " STATUS");

						}

						break;
					case "ch":

						var chan = _db.View_TransChannel(ses.LoginID, service, 0, v1, v2);

						ViewBag.DisplayTickets = chan;

						ViewBag.DiplayTitle = "VIEW TRANSACTION VIA " + service.ToUpper() + " CHANNEL";

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("Dashboard", ses.WrkStn, "VIEW TRANSACTION VIA " + service.ToUpper() + " CHANNEL");

						break;

					case "rc":

						var reas = _db.View_TransReason(ses.LoginID, service, 0, v1, v2);

						ViewBag.DisplayTickets = reas;

						ViewBag.DiplayTitle = "VIEW TRANSACTION VIA " + cnt.ToUpper() + " REASON CODE";

						string wrkstn2 = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog2 = _db.GetAuditLogs("Dashboard", wrkstn2, "VIEW TRANSACTION VIA " + service.ToUpper() + " REASON CODE");

						break;

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
		public IActionResult DisplayTicketDated(string dt1, string dt2)
		{
			try
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				var v1 = (dt1 == "XXX") ? TranDate : dt1;
				var v2 = (dt2 == "XXX") ? TranDate : dt2;
				int cnt = 0;

				string _id = Security.DeCryptor(HttpContext.Session.GetString("ayd"));
				string _service = HttpContext.Session.GetString("serbis");
				string _cnt = HttpContext.Session.GetString("kawnt");

				if (v1 != null)
				{
					cnt = 1;
				}

				switch (_id)
				{
					case "tt":

						var stat = _db.View_TransStatus(ses.LoginID, _service, cnt, v1, v2);

						ViewBag.DisplayTickets = stat;

						ViewBag.DiplayTitle = "VIEW TRANSACTION VIA " + _service.ToUpper() + " STATUS";

						break;
					case "ch":

						var chan = _db.View_TransChannel(ses.LoginID, _service, cnt, v1, v2);

						ViewBag.DisplayTickets = chan;

						ViewBag.DiplayTitle = "VIEW TRANSACTION VIA " + _service.ToUpper() + " CHANNEL";

						break;

					case "rc":

						var reas = _db.View_TransReason(ses.LoginID, _service, cnt, v1, v2);

						ViewBag.DisplayTickets = reas;

						ViewBag.DiplayTitle = "VIEW TRANSACTION VIA " + _cnt.ToUpper() + " REASON CODE";

						break;
				}

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

		public IActionResult ExportTickets()
		{
			var stream = new System.IO.MemoryStream();
			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));
			var myInput = HttpContext.Session.GetString("SearchValue");
			var dt1 = (HttpContext.Session.GetString("dt1") == "XXX") ? TranDate : HttpContext.Session.GetString("dt1"); 
			var dt2 = (HttpContext.Session.GetString("dt2") == "XXX") ? TranDate : HttpContext.Session.GetString("dt2");
			string id = Security.DeCryptor(HttpContext.Session.GetString("ayd"));
			string service = HttpContext.Session.GetString("serbis");
			string title = string.Empty;

			using (ExcelPackage package = new(stream))
			{

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

				switch (id)
				{
					case "tt":

						title = "Status";

						var v = (dt1 == "XXX") ? _db.View_TransStatus(ses.LoginID, service, 0, dt1, dt2) : _db.View_TransStatus(ses.LoginID, service, 0, dt1, dt2);

						if (!string.IsNullOrEmpty(myInput))
						{
							v = v.Where(p => p.TicketNo.Contains(myInput.ToLower()) ||
										 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
										 p.DateReceived.ToString().Contains(myInput.ToLower()) ||
										 p.Status.ToLower().Contains(myInput.ToLower()) ||
										 p.resolved.ToString().Contains(myInput.ToLower()) ||
										 p.elapsed.ToString().Contains(myInput.ToLower()) ||
										 p.agent.Contains(myInput.ToLower())
							);
						}

						ws.Row(1).Height = 20;
						ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Row(1).Style.Font.Bold = true;

						ws.Cells["A1"].Value = "TICKET NO.";
						ws.Cells["B1"].Value = "CUSTOMER NAME";
						ws.Cells["C1"].Value = "LAST UPDATE";
						ws.Cells["D1"].Value = "STATUS";
						ws.Cells["E1"].Value = "DUE DATE";
						ws.Cells["F1"].Value = "ELAPSED";
						ws.Cells["G1"].Value = "AGENT";


						int rowStart = 2;
						foreach (var item in v)
						{
							ws.Cells[string.Format("A{0}", rowStart)].Value = item.TicketNo;
							ws.Cells[string.Format("B{0}", rowStart)].Value = item.CustomerName;
							ws.Cells[string.Format("C{0}", rowStart)].Value = item.DateReceived.ToString("MM/dd/yyyy hh:mm tt");
							ws.Cells[string.Format("D{0}", rowStart)].Value = item.Status;
							ws.Cells[string.Format("E{0}", rowStart)].Value = item.resolved.ToString("MM/dd/yyyy");
							ws.Cells[string.Format("F{0}", rowStart)].Value = item.elapsed;
							ws.Cells[string.Format("G{0}", rowStart)].Value = item.agent;

							rowStart++;
						}

						break;

					case "ch":

						title = "Channel";

						var c = (dt1 == "XXX") ? _db.View_TransChannel(ses.LoginID, service, 0, dt1, dt2) : _db.View_TransChannel(ses.LoginID, service, 0, dt1, dt2);

						if (!string.IsNullOrEmpty(myInput))
						{
							c = c.Where(p => p.TicketNo.Contains(myInput.ToLower()) ||
										 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
										 p.DateReceived.ToString().Contains(myInput.ToLower()) ||
										 p.Status.ToLower().Contains(myInput.ToLower()) ||
										 p.resolved.ToString().Contains(myInput.ToLower()) ||
										 p.elapsed.ToString().Contains(myInput.ToLower()) ||
										 p.agent.Contains(myInput.ToLower())
							);
						}

						ws.Row(1).Height = 20;
						ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Row(1).Style.Font.Bold = true;

						ws.Cells["A1"].Value = "TICKET NO.";
						ws.Cells["B1"].Value = "CUSTOMER NAME";
						ws.Cells["C1"].Value = "LAST UPDATE";
						ws.Cells["D1"].Value = "STATUS";
						ws.Cells["E1"].Value = "DUE DATE";
						ws.Cells["F1"].Value = "ELAPSED";
						ws.Cells["G1"].Value = "AGENT";


						int rowChan = 2;

						foreach (var item in c)
						{
							ws.Cells[string.Format("A{0}", rowChan)].Value = item.TicketNo;
							ws.Cells[string.Format("B{0}", rowChan)].Value = item.CustomerName;
							ws.Cells[string.Format("C{0}", rowChan)].Value = item.DateReceived.ToString("MM/dd/yyyy hh:mm tt");
							ws.Cells[string.Format("D{0}", rowChan)].Value = item.Status;
							ws.Cells[string.Format("E{0}", rowChan)].Value = item.resolved.ToString("MM/dd/yyyy");
							ws.Cells[string.Format("F{0}", rowChan)].Value = item.elapsed;
							ws.Cells[string.Format("G{0}", rowChan)].Value = item.agent;

							rowChan++;
						}

						break;
					case "rc":

						title = "Reason Code";

						var r = (dt1 == "XXX") ? _db.View_TransReason(ses.LoginID, service, 0, dt1, dt2) : _db.View_TransReason(ses.LoginID, service, 0, dt1, dt2);

						if (!string.IsNullOrEmpty(myInput))
						{
							r = r.Where(p => p.TicketNo.Contains(myInput.ToLower()) ||
										 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
										 p.DateReceived.ToString().Contains(myInput.ToLower()) ||
										 p.Status.ToLower().Contains(myInput.ToLower()) ||
										 p.resolved.ToString().Contains(myInput.ToLower()) ||
										 p.elapsed.ToString().Contains(myInput.ToLower()) ||
										 p.agent.Contains(myInput.ToLower())
							);
						}

						ws.Row(1).Height = 20;
						ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Row(1).Style.Font.Bold = true;

						ws.Cells["A1"].Value = "TICKET NO.";
						ws.Cells["B1"].Value = "CUSTOMER NAME";
						ws.Cells["C1"].Value = "LAST UPDATE";
						ws.Cells["D1"].Value = "STATUS";
						ws.Cells["E1"].Value = "DUE DATE";
						ws.Cells["F1"].Value = "ELAPSED";
						ws.Cells["G1"].Value = "AGENT";


						int rowOne = 2;

						foreach (var item in r)
						{
							ws.Cells[string.Format("A{0}", rowOne)].Value = item.TicketNo;
							ws.Cells[string.Format("B{0}", rowOne)].Value = item.CustomerName;
							ws.Cells[string.Format("C{0}", rowOne)].Value = item.DateReceived.ToString("MM/dd/yyyy hh:mm tt");
							ws.Cells[string.Format("D{0}", rowOne)].Value = item.Status;
							ws.Cells[string.Format("E{0}", rowOne)].Value = item.resolved.ToString("MM/dd/yyyy");
							ws.Cells[string.Format("F{0}", rowOne)].Value = item.elapsed;
							ws.Cells[string.Format("G{0}", rowOne)].Value = item.agent;

							rowOne++;
						}

						break;
					case "st":

						title = "SERVICE";

						switch (service)
						{
							case "TOTAL DUE TODAY":
								var td1 = _db.DisplayServiceTickets(ses.LoginID, ses.Dept, ses.UType, 1);

								if (!string.IsNullOrEmpty(myInput))
								{
									td1 = td1.Where(p => p.TicketNo.Contains(myInput.ToLower()) ||
												 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
												 p.DateReceived.ToString().Contains(myInput.ToLower()) ||
												 p.Status.ToLower().Contains(myInput.ToLower()) ||
												 p.Resolved.ToString().Contains(myInput.ToLower()) ||
												 p.elapsed.ToString().Contains(myInput.ToLower()) ||
												 p.agent.Contains(myInput.ToLower())
									);
								}

								ws.Row(1).Height = 20;
								ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
								ws.Row(1).Style.Font.Bold = true;

								ws.Cells["A1"].Value = "TICKET NO.";
								ws.Cells["B1"].Value = "CUSTOMER NAME";
								ws.Cells["C1"].Value = "LAST UPDATE";
								ws.Cells["D1"].Value = "STATUS";
								ws.Cells["E1"].Value = "DUE DATE";
								ws.Cells["F1"].Value = "ELAPSED";
								ws.Cells["G1"].Value = "AGENT";


								int rowdt = 2;

								foreach (var item in td1)
								{
									ws.Cells[string.Format("A{0}", rowdt)].Value = item.TicketNo;
									ws.Cells[string.Format("B{0}", rowdt)].Value = item.CustomerName;
									ws.Cells[string.Format("C{0}", rowdt)].Value = item.DateReceived.ToString("MM/dd/yyyy hh:mm tt");
									ws.Cells[string.Format("D{0}", rowdt)].Value = item.Status;
									ws.Cells[string.Format("E{0}", rowdt)].Value = item.Resolved.ToString("MM/dd/yyyy");
									ws.Cells[string.Format("F{0}", rowdt)].Value = item.elapsed;
									ws.Cells[string.Format("G{0}", rowdt)].Value = item.agent;

									rowdt++;
								}

								break;
							case "TOTAL DUE WITHIN THE NEXT 7 DAYS":
								var td2 = _db.DisplayServiceTickets(ses.LoginID, ses.Dept, ses.UType, 2);

								if (!string.IsNullOrEmpty(myInput))
								{
									td2 = td2.Where(p => p.TicketNo.Contains(myInput.ToLower()) ||
												 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
												 p.DateReceived.ToString().Contains(myInput.ToLower()) ||
												 p.Status.ToLower().Contains(myInput.ToLower()) ||
												 p.Resolved.ToString().Contains(myInput.ToLower()) ||
												 p.elapsed.ToString().Contains(myInput.ToLower()) ||
												 p.agent.Contains(myInput.ToLower())
									);
								}

								ws.Row(1).Height = 20;
								ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
								ws.Row(1).Style.Font.Bold = true;

								ws.Cells["A1"].Value = "TICKET NO.";
								ws.Cells["B1"].Value = "CUSTOMER NAME";
								ws.Cells["C1"].Value = "LAST UPDATE";
								ws.Cells["D1"].Value = "STATUS";
								ws.Cells["E1"].Value = "DUE DATE";
								ws.Cells["F1"].Value = "ELAPSED";
								ws.Cells["G1"].Value = "AGENT";


								int rowdnex = 2;

								foreach (var item in td2)
								{
									ws.Cells[string.Format("A{0}", rowdnex)].Value = item.TicketNo;
									ws.Cells[string.Format("B{0}", rowdnex)].Value = item.CustomerName;
									ws.Cells[string.Format("C{0}", rowdnex)].Value = item.DateReceived.ToString("MM/dd/yyyy hh:mm tt");
									ws.Cells[string.Format("D{0}", rowdnex)].Value = item.Status;
									ws.Cells[string.Format("E{0}", rowdnex)].Value = item.Resolved.ToString("MM/dd/yyyy");
									ws.Cells[string.Format("F{0}", rowdnex)].Value = item.elapsed;
									ws.Cells[string.Format("G{0}", rowdnex)].Value = item.agent;

									rowdnex++;
								}

								break;

							case "TOTAL PAST DUE":
								var td3 = _db.DisplayServiceTickets(ses.LoginID, ses.Dept, ses.UType, 3);

								if (!string.IsNullOrEmpty(myInput))
								{
									td3 = td3.Where(p => p.TicketNo.Contains(myInput.ToLower()) ||
												 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
												 p.DateReceived.ToString().Contains(myInput.ToLower()) ||
												 p.Status.ToLower().Contains(myInput.ToLower()) ||
												 p.Resolved.ToString().Contains(myInput.ToLower()) ||
												 p.elapsed.ToString().Contains(myInput.ToLower()) ||
												 p.agent.Contains(myInput.ToLower())
									);
								}

								ws.Row(1).Height = 20;
								ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
								ws.Row(1).Style.Font.Bold = true;

								ws.Cells["A1"].Value = "TICKET NO.";
								ws.Cells["B1"].Value = "CUSTOMER NAME";
								ws.Cells["C1"].Value = "LAST UPDATE";
								ws.Cells["D1"].Value = "STATUS";
								ws.Cells["E1"].Value = "DUE DATE";
								ws.Cells["F1"].Value = "ELAPSED";
								ws.Cells["G1"].Value = "AGENT";


								int rowpd = 2;

								foreach (var item in td3)
								{
									ws.Cells[string.Format("A{0}", rowpd)].Value = item.TicketNo;
									ws.Cells[string.Format("B{0}", rowpd)].Value = item.CustomerName;
									ws.Cells[string.Format("C{0}", rowpd)].Value = item.DateReceived.ToString("MM/dd/yyyy hh:mm tt");
									ws.Cells[string.Format("D{0}", rowpd)].Value = item.Status;
									ws.Cells[string.Format("E{0}", rowpd)].Value = item.Resolved.ToString("MM/dd/yyyy");
									ws.Cells[string.Format("F{0}", rowpd)].Value = item.elapsed;
									ws.Cells[string.Format("G{0}", rowpd)].Value = item.agent;

									rowpd++;
								}

								break;

						}

						break;

					case "ec":

						title = "";

						switch (service)
						{
							case "TOTAL ESCALATION DUE FOR 3 DAYS":

								var te3 = _db.View_TransEscalate(1);

								if (!string.IsNullOrEmpty(myInput))
								{
									te3 = te3.Where(p => p.TicketNo.Contains(myInput.ToLower()) ||
												 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
												 p.DateReceived.ToString().Contains(myInput.ToLower()) ||
												 p.Status.ToLower().Contains(myInput.ToLower()) ||
												 p.Resolved.ToString().Contains(myInput.ToLower()) ||
												 p.elapsed.ToString().Contains(myInput.ToLower()) ||
												 p.agent.Contains(myInput.ToLower())
									);
								}

								ws.Row(1).Height = 20;
								ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
								ws.Row(1).Style.Font.Bold = true;

								ws.Cells["A1"].Value = "TICKET NO.";
								ws.Cells["B1"].Value = "CUSTOMER NAME";
								ws.Cells["C1"].Value = "LAST UPDATE";
								ws.Cells["D1"].Value = "STATUS";
								ws.Cells["E1"].Value = "DUE DATE";
								ws.Cells["F1"].Value = "ELAPSED";
								ws.Cells["G1"].Value = "AGENT";


								int rowte = 2;

								foreach (var item in te3)
								{
									ws.Cells[string.Format("A{0}", rowte)].Value = item.TicketNo;
									ws.Cells[string.Format("B{0}", rowte)].Value = item.CustomerName;
									ws.Cells[string.Format("C{0}", rowte)].Value = item.DateReceived.ToString("MM/dd/yyyy hh:mm tt");
									ws.Cells[string.Format("D{0}", rowte)].Value = item.Status;
									ws.Cells[string.Format("E{0}", rowte)].Value = item.Resolved.ToString("MM/dd/yyyy");
									ws.Cells[string.Format("F{0}", rowte)].Value = item.elapsed;
									ws.Cells[string.Format("G{0}", rowte)].Value = item.agent;

									rowte++;
								}

								break;
							case "TOTAL ESCALATION DUE FOR 7 DAYS":
								
								var te7 = _db.View_TransEscalate(2);

								if (!string.IsNullOrEmpty(myInput))
								{
									te7 = te7.Where(p => p.TicketNo.Contains(myInput.ToLower()) ||
												 p.CustomerName.ToLower().Contains(myInput.ToLower()) ||
												 p.DateReceived.ToString().Contains(myInput.ToLower()) ||
												 p.Status.ToLower().Contains(myInput.ToLower()) ||
												 p.Resolved.ToString().Contains(myInput.ToLower()) ||
												 p.elapsed.ToString().Contains(myInput.ToLower()) ||
												 p.agent.Contains(myInput.ToLower())
									);
								}

								ws.Row(1).Height = 20;
								ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
								ws.Row(1).Style.Font.Bold = true;

								ws.Cells["A1"].Value = "TICKET NO.";
								ws.Cells["B1"].Value = "CUSTOMER NAME";
								ws.Cells["C1"].Value = "LAST UPDATE";
								ws.Cells["D1"].Value = "STATUS";
								ws.Cells["E1"].Value = "DUE DATE";
								ws.Cells["F1"].Value = "ELAPSED";
								ws.Cells["G1"].Value = "AGENT";


								int rowtet = 2;

								foreach (var item in te7)
								{
									ws.Cells[string.Format("A{0}", rowtet)].Value = item.TicketNo;
									ws.Cells[string.Format("B{0}", rowtet)].Value = item.CustomerName;
									ws.Cells[string.Format("C{0}", rowtet)].Value = item.DateReceived.ToString("MM/dd/yyyy hh:mm tt");
									ws.Cells[string.Format("D{0}", rowtet)].Value = item.Status;
									ws.Cells[string.Format("E{0}", rowtet)].Value = item.Resolved.ToString("MM/dd/yyyy");
									ws.Cells[string.Format("F{0}", rowtet)].Value = item.elapsed;
									ws.Cells[string.Format("G{0}", rowtet)].Value = item.agent;

									rowtet++;
								}

								break;

						}

						break;

				}


				ws.Cells["A2:G2"].AutoFitColumns();

				package.Save();
			}

			string fileName = "Dashboard " + title + " " + service + " " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		public IActionResult Privacy()
		{
			return View();
		}

		[ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
		public IActionResult Error()
		{
			return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
		}

		
	}
}
