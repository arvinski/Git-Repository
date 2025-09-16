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
	public class AdminController : Controller
	{
		private readonly AgentInterface _db;

		public AdminController(AgentInterface db)
		{
			_db = db;
		}

		public JsonResult GetSearch(string myInput/*, string dt1, string dt2*/)
		{

			HttpContext.Session.SetString("SearchValue", myInput ?? "");
			//HttpContext.Session.SetString("dt1", dt1 ?? "");
			//HttpContext.Session.SetString("dt2", dt2 ?? "");

			return Json(new { success = true });
		}

		public IActionResult Index()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "Index");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}


				ViewBag.Interface = "Admin";
				ViewBag.Page = "Index";

				ViewBag.ListUser = _db.ListUsers();

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				//bool isLog = _db.GetAuditLogs("User Maintenance", wrkstn, "View List of Users");

				bool isLog = _db.GetAuditLogs("User Maintenance", ses.WrkStn, ses.FullName + " view the List of Users");

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}

		[HttpGet]
		public IActionResult NewUser()
		{
			ViewBag.ListGroup = new SelectList(_db.ListGroup(), "GroupID", "GroupDescription");
			ViewBag.ListDept = new SelectList(_db.ListDept(), "Dept", "Department");

			return PartialView("_NewUser");
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public IActionResult NewUser(NewUser usr)
		{
			if (ModelState.IsValid)
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				try
				{
					bool isUser = _db.ChkUser(usr.loginid);

					if (!isUser)
					{
						bool isCreated = _db.CreateUser(usr);

						if (isCreated)
						{
							var ps = _db.GetUserPass(usr.loginid);

							bool sendUser = _db.SendEmailNewUser(usr.loginid, usr.email, usr.fullname, ps.Password);

							if (sendUser)
							{
								//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

								bool isLog = _db.GetAuditLogs("User Maintenance", ses.WrkStn, "Add New User " + usr.fullname);

								TempData["msgOk"] = "A new user was successfully added/created, and credentials have been sent to his/her email.";

								return RedirectToAction("Index");
							}
							else
							{
								TempData["msgSendErr"] = "Email failed to send into the recipient";

								return RedirectToAction("Index");
							}


						}
						else
						{
							TempData["msgErr"] = "Creation of new user was not successfull";

							return RedirectToAction("Index");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing user, please enter new user";
						return RedirectToAction("Index");

					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();
					TempData["msgErr"] = msg;
					return RedirectToAction("Index");
				}
			}

			return PartialView("_NewUser", usr);

		}

		[HttpGet]
		public IActionResult EditUser(string id)
		{
			int _id = Convert.ToInt32(id);

			ViewBag.ListGroup = new SelectList(_db.ListGroup(), "GroupID", "GroupDescription");
			ViewBag.ListDept = new SelectList(_db.ListDept(), "Dept", "Department");

			EditUser edt = _db.ViewUsers(_id);

			return PartialView("_EditUser", edt);
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public IActionResult EditUser(EditUser edt)
		{

			if (ModelState.IsValid)
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				try
				{

					bool isUpdated = _db.UpdateUser(edt);

					if (isUpdated)
					{
						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("User Maintenance", ses.WrkStn, "Edit User " + edt.fullname);

						TempData["msgOk"] = "User successfully updated";

						return RedirectToAction("Index", "Admin");
					}
				}
				catch (Exception)
				{
					return RedirectToAction("Error", "Home");
				}
			}


			return PartialView("_EditUser", edt);
		}

		[HttpGet]
		public JsonResult DelUser(string id)
		{
			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			int _id = Convert.ToInt32(id);

			EditUser del = _db.ViewUsers(_id);

			bool isDeleted = _db.DeleteUser(_id);

			if (isDeleted)
			{

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("User Maintenance", ses.WrkStn, "Delete User " + del.fullname);

				return Json(new { success = true, message = "Delete successful" });
			}
			else
			{
				return Json(new { success = false, message = "Error while Deleting" });
			}
		}

		public IActionResult ExportUsers()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");

			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListUsers();

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.loginid.ToLower().Contains(myInput.ToLower()) ||
								 p.fullname.ToLower().Contains(myInput.ToLower()) ||
								 p.group.ToLower().Contains(myInput.ToLower()) ||
								 p.department.ToLower().Contains(myInput.ToLower()) ||
								 p.email.ToLower().Contains(myInput.ToLower()) ||
								 p.stat.ToLower().Contains(myInput.ToLower()) ||
								 p.expirydate.ToLower().Contains(myInput.ToLower())
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "USER ID";
				ws.Cells["B1"].Value = "USER NAME";
				ws.Cells["C1"].Value = "GROUP";
				ws.Cells["D1"].Value = "DEPARTMENT";
				ws.Cells["E1"].Value = "EMAIL";
				ws.Cells["F1"].Value = "STATUS";
				ws.Cells["G1"].Value = "EXPIRY DATE";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.loginid;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.fullname;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.group;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.department;
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.email;
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.stat;
					ws.Cells[string.Format("G{0}", rowStart)].Value = item.expirydate;

					rowStart++;
				}

				ws.Cells["A:G"].AutoFitColumns();

				package.Save();
			}

			string fileName = "List of Users as of " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		public IActionResult ReasonCodes()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "ReasonCodes");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Maintenance";
				ViewBag.Page = "ReasonCodes";

				//ViewBag.Pg = "Users";

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "View List of Reason Codes");

				ViewBag.ListReasonCodes = _db.ListReasonCodes();

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}

		[HttpGet]
		public IActionResult NewReason()
		{

			ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("RC"), "ID", "Description");

			return PartialView("_NewReason");
		}


		[HttpPost]
		public IActionResult NewReason(NewReason rcd)
		{
			if (ModelState.IsValid)
			{
				try
				{
					var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					bool isRCode = _db.ChkRCode(rcd.RCode);

					if (!isRCode)
					{
						bool isCreated = _db.CreateReasonCode(rcd);

						if (isCreated)
						{
							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Create New Reason Code " + rcd.RCode);

							TempData["msgOk"] = "Reason Code has been successfully inserted.";

							return RedirectToAction("ReasonCodes");
						}
						else
						{
							TempData["msgErr"] = "Creation of new reason code was not successfull";

							return RedirectToAction("ReasonCodes");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing reason code, please enter new reason code";
						return RedirectToAction("ReasonCodes");

					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();
					TempData["msgErr"] = msg;

					return RedirectToAction("ReasonCodes");
				}
			}

			ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("RC"), "ID", "Description");

			return PartialView("_NewReason", rcd);
		}

		[HttpGet]
		public IActionResult EditReason(string id)
		{
			int _id = Convert.ToInt32(id);

			EditReason edt = _db.ViewReasonCode(_id);

			ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("RC"), "ID", "Description");

			return PartialView("_EditReason", edt);
		}

		[HttpPost]
		public IActionResult EditReason(EditReason edt)
		{

			if (ModelState.IsValid)
			{
				try
				{
					var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					bool isUpdated = _db.UpdateReasonCode(edt);

					if (isUpdated)
					{

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Edit Reason Code " + edt.RCode);

						TempData["msgOk"] = "Reason Code has been successfully updated.";

						return RedirectToAction("ReasonCodes");
					}
					else
					{
						TempData["msgErr"] = "Updating of new reason code was not successfull";

						return RedirectToAction("ReasonCodes");
					}

				}
				catch (Exception)
				{
					return RedirectToAction("Error", "Home");
				}
			}

			ViewBag.cmbClassification = new SelectList(_db.Get_Dropdown("RC"), "ID", "Description");

			return PartialView("_EditReason", edt);
		}

		[HttpGet]
		public JsonResult DelReason(string id)
		{
			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			int _id = Convert.ToInt32(id);

			EditReason v = _db.ViewReasonCode(_id);

			bool isRCode = _db.ChkRCodeActivity(v.RCode);

			if (!isRCode)
			{
				bool isDeleted = _db.DeleteReason(_id);

				if (isDeleted)
				{
					//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

					bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Delete Reason Code " + v.RCode);

					return Json(new { success = true, message = "Delete successful" });
				}
				else
				{
					return Json(new { success = false, message = "Error while Deleting" });
				}

			}
			else
			{
				return Json(new { success = false, message = "Reason code exist in the activity!" });
			}

		}

		public IActionResult ExportReason()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");

			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListReasonCodes();

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.RCode.ToLower().Contains(myInput.ToLower()) ||
								 p.Reason.ToLower().Contains(myInput.ToLower()) ||
								 p.Description.ToLower().Contains(myInput.ToLower()) ||
								 p.FirstLetter.ToLower().Contains(myInput.ToLower()) ||
								 p.SecondLetter.ToLower().Contains(myInput.ToLower()) ||
								 p.ThirdLetter.ToLower().Contains(myInput.ToLower()) ||
								 p.Abbreviation.ToLower().Contains(myInput.ToLower()) ||
								 p.Class.ToLower().Contains(myInput.ToLower())
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "REASON CODE";
				ws.Cells["B1"].Value = "REASON";
				ws.Cells["C1"].Value = "DESCRIPTION";
				ws.Cells["D1"].Value = "FIRST LETTER";
				ws.Cells["E1"].Value = "SECOND LETTER";
				ws.Cells["F1"].Value = "THIRD LETTER";
				ws.Cells["G1"].Value = "ABBREVIATION";
				ws.Cells["H1"].Value = "CLASSIFICATION";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.RCode;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.Reason;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.Description;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.FirstLetter;
					ws.Cells[string.Format("E{0}", rowStart)].Value = item.SecondLetter;
					ws.Cells[string.Format("F{0}", rowStart)].Value = item.ThirdLetter;
					ws.Cells[string.Format("G{0}", rowStart)].Value = item.Abbreviation;
					ws.Cells[string.Format("H{0}", rowStart)].Value = item.Class;

					rowStart++;
				}

				ws.Cells["A:H"].AutoFitColumns();

				package.Save();
			}

			string fileName = "List of Reason Codes as of " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		public IActionResult PaymentInstitution()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "PaymentInstitution");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}


				ViewBag.Interface = "Maintenance";
				ViewBag.Page = "PaymentInstitution";

				//ViewBag.Pg = "Users";
				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "View List of Payment Institution");


				ViewBag.ListPaymentInstitutions = _db.ListATM("IT");

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}

		[HttpGet]
		public IActionResult NewPayInst()
		{

			return PartialView("_NewPayInst");
		}

		[HttpPost]
		public IActionResult NewPayInst(NewATM atm)
		{
			if (ModelState.IsValid)
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				try
				{
					bool ischkATM = _db.ChkATM(atm.Description, "IT");

					if (!ischkATM)
					{
						bool isCreated = _db.CreateATM(atm.Description, "IT");

						if (isCreated)
						{
							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Create New Payment Institution " + atm.Description);

							TempData["msgOk"] = "Payment Institution has been successfully inserted.";

							return RedirectToAction("PaymentInstitution");
						}
						else
						{
							TempData["msgErr"] = "Creation of new payment institution was not successfull";

							return RedirectToAction("PaymentInstitution");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing payment institution, please enter new institution name";
						return RedirectToAction("PaymentInstitution");

					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("PaymentInstitution");
				}
			}

			return PartialView("_NewPayInst");
		}

		[HttpGet]
		public IActionResult EditPayInst(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM edt = _db.ViewATM(_id);

			return PartialView("_EditPayInst", edt);
		}

		[HttpPost]
		public IActionResult EditPayInst(EditATM edt)
		{
			if (ModelState.IsValid)
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				try
				{
					bool ischkATM = _db.ChkATM(edt.Description, "IT");

					if (!ischkATM)
					{
						bool isUpdated = _db.UpdateATM(edt);

						if (isUpdated)
						{

							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Edit Payment Institution " + edt.Description);

							TempData["msgOk"] = "Payment Insitution has been successfully updated.";

							return RedirectToAction("PaymentInstitution");
						}
						else
						{
							TempData["msgErr"] = "Updating of new payment institution was not successfull";

							return RedirectToAction("PaymentInstitution");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing payment institution, please enter new institution name";
						return RedirectToAction("PaymentInstitution");
					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("PaymentInstitution");

				}
			}

			return PartialView("_EditPayInst", edt);
		}

		[HttpGet]
		public JsonResult DelPayInst(string id)
		{
			var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			int _id = Convert.ToInt32(id);

			EditATM v = _db.ViewATM(_id);

			bool isExist = _db.ChkPaymentOfActivity(v.Description);

			if (!isExist)
			{
				bool isDeleted = _db.DeleteATM(_id);

				if (isDeleted)
				{
					//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

					bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Delete Payment Institution " + v.Description);

					return Json(new { success = true, message = "Delete successful" });
				}
				else
				{
					return Json(new { success = false, message = "Error while Deleting" });
				}

			}
			else
			{
				return Json(new { success = false, message = "You cannot delete this because it exist in the activity!" });
			}

		}

		public IActionResult ExportPaymentInst()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");

			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListATM("IT");

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.Description.ToLower().Contains(myInput.ToLower()));
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "INSTITUTION NAME";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.Description;

					rowStart++;
				}

				ws.Cells["A:A"].AutoFitColumns();

				package.Save();
			}

			string fileName = "List of Payment Institution as of " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);


		}

		public IActionResult Bancnet()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "Bancnet");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Maintenance";
				ViewBag.Page = "Bancnet";

				//ViewBag.Pg = "Users";

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "View List of Bancnet");

				ViewBag.ListBancnets = _db.ListATM("BN");

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}

		[HttpGet]
		public IActionResult NewBancnet()
		{

			return PartialView("_NewBancnet");
		}

		[HttpPost]
		public IActionResult NewBancnet(NewATM atm)
		{
			if (ModelState.IsValid)
			{
				try
				{
					bool ischkATM = _db.ChkATM(atm.Description, "BN");

					if (!ischkATM)
					{
						bool isCreated = _db.CreateATM(atm.Description, "BN");

						if (isCreated)
						{
							var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Create New Bancnet " + atm.Description);


							TempData["msgOk"] = "Bancnet has been successfully inserted.";

							return RedirectToAction("Bancnet");
						}
						else
						{
							TempData["msgErr"] = "Creation of new Bancnet was not successfull";

							return RedirectToAction("Bancnet");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing Bancnet, please enter new Bancnet name";
						return RedirectToAction("Bancnet");

					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("Bancnet");
				}
			}

			return PartialView("_NewBancnet");
		}

		[HttpGet]
		public IActionResult EditBancnet(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM edt = _db.ViewATM(_id);

			return PartialView("_EditBancnet", edt);
		}

		[HttpPost]
		public IActionResult EditBancnet(EditATM edt)
		{
			if (ModelState.IsValid)
			{
				try
				{
					bool ischkATM = _db.ChkATM(edt.Description, "BN");

					if (!ischkATM)
					{
						bool isUpdated = _db.UpdateATM(edt);

						if (isUpdated)
						{

							var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Edit Bancnet " + edt.Description);

							TempData["msgOk"] = "Bancnet has been successfully updated.";

							return RedirectToAction("Bancnet");
						}
						else
						{
							TempData["msgErr"] = "Updating of new Bancnet was not successfull";

							return RedirectToAction("Bancnet");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing Bancnet, please enter new Bancnet name";
						return RedirectToAction("Bancnet");
					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("Bancnet");

				}
			}

			return PartialView("_EditBancnet", edt);
		}

		[HttpGet]
		public JsonResult DelBancnet(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM v = _db.ViewATM(_id);

			bool isExist = _db.ChkBancnetActivity(v.Description);

			if (!isExist)
			{
				bool isDeleted = _db.DeleteATM(_id);

				if (isDeleted)
				{

					var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

					bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Delete Bancnet " + v.Description);

					return Json(new { success = true, message = "Delete successful" });
				}
				else
				{
					return Json(new { success = false, message = "Error while Deleting" });
				}

			}
			else
			{
				return Json(new { success = false, message = "You cannot delete this because it exist in the activity!" });
			}

		}

		public IActionResult ExportBancnet()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");

			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListATM("BN");

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.Description.ToLower().Contains(myInput.ToLower()));
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "BANCNET NAME";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.Description;

					rowStart++;
				}

				ws.Cells["A:A"].AutoFitColumns();

				package.Save();
			}

			string fileName = "List of Bancnet as of " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);
		}


		public IActionResult ATMUsed()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "ATMUsed");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Maintenance";
				ViewBag.Page = "ATMUsed";

				//ViewBag.Pg = "Users";

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "View List of ATM Used");


				ViewBag.ListATMUsed = _db.ListATM("AU");

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}

		[HttpGet]
		public IActionResult NewATM()
		{

			return PartialView("_NewATM");
		}

		[HttpPost]
		public IActionResult NewATM(NewATM atm)
		{
			if (ModelState.IsValid)
			{
				try
				{
					bool ischkATM = _db.ChkATM(atm.Description, "AU");

					if (!ischkATM)
					{
						bool isCreated = _db.CreateATM(atm.Description, "AU");

						if (isCreated)
						{
							var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Create New ATM " + atm.Description);

							TempData["msgOk"] = "ATM has been successfully inserted.";

							return RedirectToAction("ATMUsed");
						}
						else
						{
							TempData["msgErr"] = "Creation of new ATM was not successfull";

							return RedirectToAction("ATMUsed");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing Bancnet, please enter new ATM name";
						return RedirectToAction("ATMUsed");

					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("ATMUsed");
				}
			}

			return PartialView("_NewATM");
		}

		[HttpGet]
		public IActionResult EditATM(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM edt = _db.ViewATM(_id);

			return PartialView("_EditATM", edt);
		}

		[HttpPost]
		public IActionResult EditATM(EditATM edt)
		{
			if (ModelState.IsValid)
			{
				try
				{
					bool ischkATM = _db.ChkATM(edt.Description, "AU");

					if (!ischkATM)
					{
						bool isUpdated = _db.UpdateATM(edt);

						if (isUpdated)
						{
							var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Edit ATM " + edt.Description);


							TempData["msgOk"] = "ATM has been successfully updated.";

							return RedirectToAction("ATMUsed");
						}
						else
						{
							TempData["msgErr"] = "Updating of new ATM was not successfull";

							return RedirectToAction("ATMUsed");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing ATM, please enter new Bancnet name";
						return RedirectToAction("ATMUsed");
					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("ATMUsed");

				}
			}

			return PartialView("_EditATM", edt);
		}

		[HttpGet]
		public JsonResult DelATM(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM v = _db.ViewATM(_id);

			bool isExist = _db.ChkATMActivity(v.Description);

			if (!isExist)
			{
				bool isDeleted = _db.DeleteATM(_id);

				if (isDeleted)
				{
					var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

					bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Delete ATM " + v.Description);


					return Json(new { success = true, message = "Delete successful" });
				}
				else
				{
					return Json(new { success = false, message = "Error while Deleting" });
				}

			}
			else
			{
				return Json(new { success = false, message = "You cannot delete this because it exist in the activity!" });
			}

		}

		public IActionResult ATMLocation()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "ATMLocation");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Maintenance";
				ViewBag.Page = "ATMLocation";

				//ViewBag.Pg = "Users";

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "View List of ATM Location");


				ViewBag.ListATMLocation = _db.ListATMLocation();

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}

		[HttpGet]
		public IActionResult NewATMLocation()
		{
			ViewBag.cmbATMUsed2 = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");

			return PartialView("_NewATMLocation");
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public IActionResult NewATMLocation(NewATMLoc atm)
		{
			if (ModelState.IsValid)
			{
				try
				{
					//bool ischkATM = _db.ChkATM(atm.Description, "AU");

					//if (!ischkATM)
					//{
					bool isCreated = _db.CreateATMLocation(atm.ATM_ID, atm.ATM_LOCATION);

					if (isCreated)
					{
						var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Create New ATM Location " + atm.ATM_LOCATION);


						TempData["msgOk"] = "ATM Branch Location has been successfully inserted.";

						return RedirectToAction("ATMLocation");
					}
					else
					{
						TempData["msgErr"] = "Creation of new ATM was not successfull";

						return RedirectToAction("ATMLocation");
					}

					//}
					//else
					//{
					//	TempData["msgDup"] = "Existing ATM Location, please enter new ATM Branch Location";
					//	return RedirectToAction("ATMLocation");

					//}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("ATMUsed");
				}
			}

			ViewBag.cmbATMUsed2 = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");

			return PartialView("_NewATMLocation");
		}

		[HttpGet]
		public IActionResult EditATMLocation(string id)
		{
			int _id = Convert.ToInt32(id);

			ViewBag.cmbATMUsed2 = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");

			EditATMLoc edt = _db.ViewATMLocation(_id);

			return PartialView("_EditATMLocation", edt);
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public IActionResult EditATMLocation(EditATMLoc edt)
		{

			if (ModelState.IsValid)
			{
				try
				{

					bool isUpdated = _db.UpdateATMLocation(edt.ATM_LOCATION, edt.ATM_LOC_ID);

					if (isUpdated)
					{
						var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Edit ATM Location " + edt.ATM_LOCATION);

						TempData["msgOk"] = "ATM Location successfully updated";

						return RedirectToAction("ATMLocation", "Admin");
					}
				}
				catch (Exception)
				{
					return RedirectToAction("Error", "Home");
				}
			}

			ViewBag.cmbATMUsed2 = new SelectList(_db.Get_Dropdown2("AU"), "ID", "Description");

			return PartialView("_EditATMLocation", edt);
		}

		[HttpGet]
		public JsonResult DelATMLoc(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATMLoc v = _db.ViewATMLocation(_id);

			//bool isExist = _db.ChkATMActivity(v.Description);

			//if (!isExist)
			//{
			bool isDeleted = _db.DeleteATMLocation(v.ID, v.ATM_LOC_ID);

			if (isDeleted)
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Delete ATM Location " + v.ATM_LOCATION);


				return Json(new { success = true, message = "Delete successful" });
			}
			else
			{
				return Json(new { success = false, message = "Error while Deleting" });
			}

			//}
			//else
			//{
			//	return Json(new { success = false, message = "You cannot delete this because it exist in the activity!" });
			//}

		}

		public IActionResult ExportATM()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");

			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListATM("AU");

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.Description.ToLower().Contains(myInput.ToLower()));
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "ATM NAME";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.Description;

					rowStart++;
				}

				ws.Cells["A:A"].AutoFitColumns();

				package.Save();
			}

			string fileName = "List of ATM Used as of " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);
		}

		public IActionResult ExportATMLocation()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");

			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListATMLocation();

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.ATM.ToLower().Contains(myInput.ToLower()) ||
						p.ATM_LOCATION.ToLower().Contains(myInput.ToLower())
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "ATM NAME";
				ws.Cells["B1"].Value = "ATM LOCATION";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.ATM;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.ATM_LOCATION;

					rowStart++;
				}

				ws.Cells["A:B"].AutoFitColumns();

				package.Save();
			}

			string fileName = "List of ATM Location as of " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);
		}

		public IActionResult SubStatus()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "SubStatus");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Maintenance";
				ViewBag.Page = "SubStatus";

				//ViewBag.Pg = "Users";

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "View List of Sub Status");


				ViewBag.ListSubStatus = _db.ListATM("SS");

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}


		[HttpGet]
		public IActionResult NewSubStatus()
		{

			return PartialView("_NewSubStatus");
		}

		[HttpPost]
		public IActionResult NewSubStatus(NewATM atm)
		{
			if (ModelState.IsValid)
			{
				try
				{
					bool ischkATM = _db.ChkATM(atm.Description, "SS");

					if (!ischkATM)
					{
						bool isCreated = _db.CreateATM(atm.Description, "SS");

						if (isCreated)
						{
							var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Create New Sub Status " + atm.Description);


							TempData["msgOk"] = "Sub Status has been successfully inserted.";

							return RedirectToAction("SubStatus");
						}
						else
						{
							TempData["msgErr"] = "Creation of new Sub Status was not successfull";

							return RedirectToAction("SubStatus");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing Sub Status, please enter new Sub Status name";
						return RedirectToAction("SubStatus");

					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("SubStatus");
				}
			}

			return PartialView("_NewSubStatus");
		}

		[HttpGet]
		public IActionResult EditSubStatus(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM edt = _db.ViewATM(_id);

			return PartialView("_EditSubStatus", edt);
		}

		[HttpPost]
		public IActionResult EditSubStatus(EditATM edt)
		{
			if (ModelState.IsValid)
			{
				try
				{
					//bool ischkATM = _db.ChkATM(edt.Description, "SS");

					//if (!ischkATM)
					//{
					bool isUpdated = _db.UpdateATM(edt);

					if (isUpdated)
					{
						var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Edit Sub Status " + edt.Description);


						TempData["msgOk"] = "Sub Status has been successfully updated.";

						return RedirectToAction("SubStatus");
					}
					else
					{
						TempData["msgErr"] = "Updating of new Sub Status was not successfull";

						return RedirectToAction("SubStatus");
					}

					//}
					//else
					//{
					//	TempData["msgDup"] = "Existing Sub Status, please enter new Bancnet name";
					//	return RedirectToAction("SubStatus");
					//}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("SubStatus");

				}
			}

			return PartialView("_EditSubStatus", edt);
		}

		[HttpGet]
		public JsonResult DelSubStatus(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM v = _db.ViewATM(_id);

			bool isExist = _db.ChkSubStatusActivity(v.Description);

			if (!isExist)
			{
				bool isDeleted = _db.DeleteATM(_id);

				if (isDeleted)
				{
					var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

					bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Delete Sub Status " + v.Description);


					return Json(new { success = true, message = "Delete successful" });
				}
				else
				{
					return Json(new { success = false, message = "Error while Deleting" });
				}

			}
			else
			{
				return Json(new { success = false, message = "You cannot delete this because it exist in the activity!" });
			}

		}

		public IActionResult ExportSubStatus()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");

			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListATM("SS");

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.Description.ToLower().Contains(myInput.ToLower()));
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "SUB STATUS";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.Description;

					rowStart++;
				}

				ws.Cells["A:A"].AutoFitColumns();

				package.Save();
			}

			string fileName = "List of Sub Status as of " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);
		}

		public IActionResult NatureComplain()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "NatureComplain");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Maintenance";
				ViewBag.Page = "NatureComplain";

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "View List of Nature Complaint");



				ViewBag.ListNatureComplain = _db.ListATM("NT");

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}

		[HttpGet]
		public IActionResult NewRemarks()
		{

			return PartialView("_NewRemarks");
		}

		[HttpPost]
		public IActionResult NewRemarks(NewATM atm)
		{
			if (ModelState.IsValid)
			{
				try
				{
					bool ischkATM = _db.ChkATM(atm.Description, "NT");

					if (!ischkATM)
					{
						bool isCreated = _db.CreateATM(atm.Description, "NT");

						if (isCreated)
						{
							var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Create New Nature Complaint " + atm.Description);


							TempData["msgOk"] = "Nature Complaint has been successfully inserted.";

							return RedirectToAction("NatureComplain");
						}
						else
						{
							TempData["msgErr"] = "Creation of new Nature Complaint was not successfull";

							return RedirectToAction("NatureComplain");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing New Complaint, please enter new Nature Complaint name";
						return RedirectToAction("NatureComplain");

					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("NatureComplain");
				}
			}

			return PartialView("_NewRemarks");
		}

		[HttpGet]
		public IActionResult EditRemarks(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM edt = _db.ViewATM(_id);

			return PartialView("_EditRemarks", edt);
		}

		[HttpPost]
		public IActionResult EditRemarks(EditATM edt)
		{
			if (ModelState.IsValid)
			{
				try
				{
					//bool ischkATM = _db.ChkATM(edt.Description, "SS");

					//if (!ischkATM)
					//{
					bool isUpdated = _db.UpdateATM(edt);

					if (isUpdated)
					{
						var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Edit Nature Complaint " + edt.Description);


						TempData["msgOk"] = "Nature Complaint has been successfully updated.";

						return RedirectToAction("NatureComplain");
					}
					else
					{
						TempData["msgErr"] = "Updating of new Nature Complaint was not successfull";

						return RedirectToAction("NatureComplain");
					}

					//}
					//else
					//{
					//	TempData["msgDup"] = "Existing Sub Status, please enter new Bancnet name";
					//	return RedirectToAction("SubStatus");
					//}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("NatureComplain");

				}
			}

			return PartialView("_EditRemarks", edt);
		}

		[HttpGet]
		public JsonResult DelRemarks(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM v = _db.ViewATM(_id);

			bool isExist = _db.ChkRemarksActivity(v.Description);

			if (!isExist)
			{
				bool isDeleted = _db.DeleteATM(_id);

				if (isDeleted)
				{
					var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

					bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Delete Nature Complaint " + v.Description);


					return Json(new { success = true, message = "Delete successful" });
				}
				else
				{
					return Json(new { success = false, message = "Error while Deleting" });
				}

			}
			else
			{
				return Json(new { success = false, message = "You cannot delete this because it exist in the activity!" });
			}

		}

		public IActionResult ExportRemarks()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");

			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListATM("NT");

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.Description.ToLower().Contains(myInput.ToLower()));
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "NATURE COMPLAINT";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.Description;

					rowStart++;
				}

				ws.Cells["A:A"].AutoFitColumns();

				package.Save();
			}

			string fileName = "List of Nature Complaint as of " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);
		}

		public IActionResult ReferredBy()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "ReferredBy");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Maintenance";
				ViewBag.Page = "ReferredBy";

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "View List of Referred By");


				ViewBag.ListReferredBy = _db.ListATM("RB");

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}

		[HttpGet]
		public IActionResult NewReferredBy()
		{

			return PartialView("_NewReferredBy");
		}

		[HttpPost]
		public IActionResult NewReferredBy(NewATM atm)
		{
			if (ModelState.IsValid)
			{
				try
				{
					bool ischkATM = _db.ChkATM(atm.Description, "RB");

					if (!ischkATM)
					{
						bool isCreated = _db.CreateATM(atm.Description, "RB");

						if (isCreated)
						{
							var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Create New Referred By " + atm.Description);

							TempData["msgOk"] = "Referred By has been successfully inserted.";

							return RedirectToAction("ReferredBy");
						}
						else
						{
							TempData["msgErr"] = "Creation of new Referred By was not successfull";

							return RedirectToAction("ReferredBy");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing Referred By, please enter new Referred By name";
						return RedirectToAction("ReferredBy");

					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("ReferredBy");
				}
			}

			return PartialView("_NewReferredBy");
		}

		[HttpGet]
		public IActionResult EditReferredBy(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM edt = _db.ViewATM(_id);

			return PartialView("_EditReferredBy", edt);
		}

		[HttpPost]
		public IActionResult EditReferredBy(EditATM edt)
		{
			if (ModelState.IsValid)
			{
				try
				{
					//bool ischkATM = _db.ChkATM(edt.Description, "SS");

					//if (!ischkATM)
					//{
					bool isUpdated = _db.UpdateATM(edt);

					if (isUpdated)
					{
						var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Edit Referred By " + edt.Description);

						TempData["msgOk"] = "Referred By has been successfully updated.";

						return RedirectToAction("ReferredBy");
					}
					else
					{
						TempData["msgErr"] = "Updating of new New ReferredBy was not successfull";

						return RedirectToAction("ReferredBy");
					}

					//}
					//else
					//{
					//	TempData["msgDup"] = "Existing Sub Status, please enter new Bancnet name";
					//	return RedirectToAction("SubStatus");
					//}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("ReferredBy");

				}
			}

			return PartialView("_EditReferredBy", edt);
		}

		[HttpGet]
		public JsonResult DelReferredBy(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM v = _db.ViewATM(_id);

			bool isExist = _db.ChkReferredByActivity(v.Description);

			if (!isExist)
			{
				bool isDeleted = _db.DeleteATM(_id);

				if (isDeleted)
				{
					var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

					bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Delete Referred By " + v.Description);

					return Json(new { success = true, message = "Delete successful" });
				}
				else
				{
					return Json(new { success = false, message = "Error while Deleting" });
				}

			}
			else
			{
				return Json(new { success = false, message = "You cannot delete this because it exist in the activity!" });
			}

		}

		public IActionResult ExportReferredBy()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");

			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListATM("RB");

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.Description.ToLower().Contains(myInput.ToLower()));
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "REFERRED BY";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.Description;

					rowStart++;
				}

				ws.Cells["A:A"].AutoFitColumns();

				package.Save();
			}

			string fileName = "List of Referred By as of " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);
		}

		public IActionResult UserMenu()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "UserMenu");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Admin";
				ViewBag.Page = "UserMenu";

				//ViewBag.Pg = "Users";

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "View List of User Menu");


				ViewBag.ListUserMenus = _db.ListUserMenu();

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}

		[HttpGet]
		public IActionResult NewMenu()
		{

			return PartialView("_NewMenu");
		}

		[HttpPost]
		public IActionResult NewMenu(NewMenu mnu)
		{
			if (ModelState.IsValid)
			{
				try
				{
					bool ischkMenu = _db.ChkMenu(mnu.View, mnu.Page);

					if (!ischkMenu)
					{
						bool isCreated = _db.CreateMenu(mnu);

						if (isCreated)
						{
							var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Create New Menu " + mnu.Title);

							TempData["msgOk"] = "User Menu has been successfully inserted.";

							return RedirectToAction("UserMenu");
						}
						else
						{
							TempData["msgErr"] = "Creation of new user menu was not successfull";

							return RedirectToAction("UserMenu");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing user menu, please enter new user menu";
						return RedirectToAction("UserMenu");

					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("UserMenu");
				}
			}

			return PartialView("_NewMenu", mnu);
		}

		[HttpGet]
		public IActionResult EditMenu(string id)
		{
			int _id = Convert.ToInt32(id);

			EditMenu edt = _db.ViewMenu(_id);

			return PartialView("_EditMenu", edt);
		}

		[HttpPost]
		public IActionResult EditMenu(EditMenu mnu)
		{
			if (ModelState.IsValid)
			{
				try
				{
					//bool ischkMenu = _db.ChkMenu(mnu.View, mnu.Page);

					//if (!ischkMenu)
					//{
					bool isCreated = _db.UpdateMenu(mnu);

					if (isCreated)
					{

						var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

						//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Edit Menu " + mnu.Title);

						TempData["msgOk"] = "User Menu has been successfully updated.";

						return RedirectToAction("UserMenu");
					}
					else
					{
						TempData["msgErr"] = "Updating of new user menu was not successfull";

						return RedirectToAction("UserMenu");
					}

					//}
					//else
					//{
					//	TempData["msgDup"] = "Existing user menu, please enter new user menu";
					//	return RedirectToAction("UserMenu");

					//}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("UserMenu");

				}
			}

			return PartialView("_EditMenu", mnu);
		}

		[HttpGet]
		public JsonResult DelMenu(string id)
		{
			int _id = Convert.ToInt32(id);

			EditMenu del = _db.ViewMenu(_id);

			bool isDeleted = _db.DeleteMenu(_id);

			if (isDeleted)
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Delete Menu " + del.Title);


				return Json(new { success = true, message = "Delete successful" });
			}
			else
			{
				return Json(new { success = false, message = "Error while Deleting" });
			}
		}

		public IActionResult EndorsedEmail()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "EndorsedEmail");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				ViewBag.Interface = "Maintenance";
				ViewBag.Page = "EndorsedEmail";

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "View List of Endorsed Email");

				ViewBag.ListEndorsed = _db.ListATM("EF");

				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}

		[HttpGet]
		public IActionResult NewEndorsed()
		{

			return PartialView("_NewEndorsed");
		}

		[HttpPost]
		public IActionResult NewEndorsed(NewATM atm)
		{
			if (ModelState.IsValid)
			{
				try
				{
					bool ischkATM = _db.ChkATM(atm.Description, "EF");

					if (!ischkATM)
					{
						bool isCreated = _db.CreateATM(atm.Description, "EF");

						if (isCreated)
						{
							var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Create New Endorsed Email " + atm.Description);

							TempData["msgOk"] = "Endorsed Email has been successfully inserted.";

							return RedirectToAction("EndorsedEmail");
						}
						else
						{
							TempData["msgErr"] = "Creation of new endorsed email was not successfull";

							return RedirectToAction("EndorsedEmail");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing email address, please enter new email address";
						return RedirectToAction("EndorsedEmail");

					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("EndorsedEmail");
				}
			}

			return PartialView("_NewEndorsed");
		}

		[HttpGet]
		public IActionResult EditEndorsed(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM edt = _db.ViewATM(_id);

			return PartialView("_EditEndorsed", edt);
		}

		[HttpPost]
		public IActionResult EditEndorsed(EditATM edt)
		{
			if (ModelState.IsValid)
			{
				try
				{
					bool ischkATM = _db.ChkATM(edt.Description, "EF");

					if (!ischkATM)
					{
						bool isUpdated = _db.UpdateATM(edt);

						if (isUpdated)
						{
							var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

							//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Edit Endorsed Email " + edt.Description);


							TempData["msgOk"] = "Endorsed Email has been successfully updated.";

							return RedirectToAction("EndorsedEmail");
						}
						else
						{
							TempData["msgErr"] = "Updating of new email was not successfull";

							return RedirectToAction("EndorsedEmail");
						}

					}
					else
					{
						TempData["msgDup"] = "Existing email address, please enter new email address";
						return RedirectToAction("EndorsedEmail");
					}
				}
				catch (Exception ex)
				{
					string msg = ex.ToString();

					TempData["msgErr"] = msg;

					return RedirectToAction("EndorsedEmail");

				}
			}

			return PartialView("_EditEndorsed", edt);
		}

		[HttpGet]
		public JsonResult DelEndorsed(string id)
		{
			int _id = Convert.ToInt32(id);

			EditATM v = _db.ViewATM(_id);

			bool isExist = _db.ChkEndorsedActivity(v.Description);

			if (!isExist)
			{
				bool isDeleted = _db.DeleteATM(_id);

				if (isDeleted)
				{
					var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

					bool isLog = _db.GetAuditLogs("Activity Tool", ses.WrkStn, "Delete Endorsed Email " + v.Description);

					return Json(new { success = true, message = "Delete successful" });
				}
				else
				{
					return Json(new { success = false, message = "Error while Deleting" });
				}

			}
			else
			{
				return Json(new { success = false, message = "You cannot delete this because it exist in the activity!" });
			}

		}

		public IActionResult ExportEndorsed()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");

			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListATM("EF");

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.Description.ToLower().Contains(myInput.ToLower()));
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "EMAIL ADDRESS";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.Description;

					rowStart++;
				}

				ws.Cells["A:A"].AutoFitColumns();

				package.Save();
			}

			string fileName = "List of Endorsed Email as of " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);
		}

		[HttpGet]
		public IActionResult UserRights(string id)
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{

				int _id = Convert.ToInt32(id);

				EditUser edt = _db.ViewUsers(_id);

				ViewBag.User = edt.fullname;

				bool isInserted = _db.InsertRights(_id);

				ViewBag.ListUserRights = _db.ListUserRights(_id);

				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("User Maintenance", ses.WrkStn, "View User Rights " + edt.fullname);


				HttpContext.Session.SetString("usrayt", _id.ToString());
				HttpContext.Session.SetString("usrnosi", edt.loginid);

				return View(edt);

			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}


		}

		[HttpPost]
		public IActionResult UserRights(IFormCollection form)
		{

			//var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			//if (ses.UType != "A")
			//{
			//    TempData["msgAuth"] = "You are not authorized to use this feature";
			//    return RedirectToAction("Index", "Home");
			//}

			int uid = Convert.ToInt16(HttpContext.Session.GetString("usrayt"));

			EditUser edt = _db.ViewUsers(uid);

			ViewBag.User = edt.fullname;

			string idd = form["ID"];

			if (idd != null)
			{
				bool isUpdate = _db.UpdateRights(uid, idd);

				if (isUpdate)
				{
					var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					//string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

					bool isLog = _db.GetAuditLogs("User Maintenance", ses.WrkStn, "Update User Rights of " + edt.fullname);


					TempData["msgOk"] = "Rights has been sucessfully updated";
				}

			}

			ViewBag.ListUserRights = _db.ListUserRights(uid);

			return View();
		}

		public IActionResult ExportUserRights()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");

			int uid = Convert.ToInt16(HttpContext.Session.GetString("usrayt"));

			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListUserRights(uid);

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.Title.ToLower().Contains(myInput.ToLower()) 
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "TITLE";
				ws.Cells["B1"].Value = "RIGHTS";

				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.Title;
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.Rights.ToString() == "1" ? "Yes" : "No";

					rowStart++;
				}

				ws.Cells["A:B"].AutoFitColumns();

				package.Save();
			}

			string nosi = HttpContext.Session.GetString("usrnosi");

			string fileName = "User Rights of " + nosi + " " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);
		}

		[HttpGet]
		public IActionResult AuditLogs()
		{

			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				var ses = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

				bool usrights = _db.ChkRights(ses.LoginID, "Admin", "AuditLogs");

				if (!usrights)
				{
					TempData["msgAuth"] = "You are not authorized to use this feature";
					return RedirectToAction("Index", "Home");
				}

				string d1 = DateTime.Now.ToString("dd-MMM-yyyy");
				string d2 = DateTime.Now.ToString("dd-MMM-yyyy");

				HttpContext.Session.SetString("dt1", d1 ?? "");
				HttpContext.Session.SetString("dt2", d2 ?? "");

				ViewBag.ListAuditLog = _db.ListAuditLogs(d1, d2);

				string wrkstn = Request.HttpContext.Connection.RemoteIpAddress + "/" + ses.LoginID.ToString();

				bool isLog = _db.GetAuditLogs("User Maintenance", wrkstn, "View List of Audit Logs");

				return View();

			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}

		[HttpPost]
		public IActionResult AuditLogs(string dt1, string dt2)
		{
			string d1 = dt1;
			string d2 = dt2;

			HttpContext.Session.SetString("dt1", d1 ?? "");
			HttpContext.Session.SetString("dt2", d2 ?? "");

			ViewBag.ListAuditLog = _db.ListAuditLogs(d1, d2);

			return View();
		}

		public IActionResult ExportAuditLogs()
		{
			var stream = new System.IO.MemoryStream();

			var myInput = HttpContext.Session.GetString("SearchValue");
			string dt1 = HttpContext.Session.GetString("dt1");
			string dt2 = HttpContext.Session.GetString("dt2");


			using (ExcelPackage package = new(stream))
			{

				var v = _db.ListAuditLogs(dt1, dt2);

				if (!string.IsNullOrEmpty(myInput))
				{
					v = v.Where(p => p.logdate.ToString().Contains(myInput) ||
						p.description.ToLower().Contains(myInput.ToLower()) ||
						p.workstn.ToLower().Contains(myInput.ToLower()) ||
						p.activity.ToLower().Contains(myInput.ToLower())
					);
				}

				ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");


				ws.Row(1).Height = 20;
				ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				ws.Row(1).Style.Font.Bold = true;

				ws.Cells["A1"].Value = "LOG DATE";
				ws.Cells["B1"].Value = "DESCRIPTION";
				ws.Cells["B1"].Value = "WORK STATION";
				ws.Cells["B1"].Value = "ACTIVITY";


				int rowStart = 2;
				foreach (var item in v)
				{
					ws.Cells[string.Format("A{0}", rowStart)].Value = item.logdate.ToString("MM/dd/yyyy hh:mm tt");
					ws.Cells[string.Format("B{0}", rowStart)].Value = item.description;
					ws.Cells[string.Format("C{0}", rowStart)].Value = item.workstn;
					ws.Cells[string.Format("D{0}", rowStart)].Value = item.activity;

					rowStart++;
				}

				ws.Cells["A:D"].AutoFitColumns();

				package.Save();
			}

			string fileName = "Audit Logs " + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".xlsx";
			string fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			stream.Position = 0;
			return File(stream, fileType, fileName);
		}


	}
}
