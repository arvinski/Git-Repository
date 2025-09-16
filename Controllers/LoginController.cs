using System;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using AgentDesktop.Models;
using AgentDesktop.Services;
using Newtonsoft.Json;

namespace AgentDesktop.Controllers
{
	public class LoginController : Controller
	{
		private readonly AgentInterface _db;
		//private IHttpContextAccessor _ipadd;
		public LoginController(AgentInterface db/*, IHttpContextAccessor ipadd*/)
		{
			_db = db;
			//_ipadd = ipadd;
		}

		public IActionResult Index()
		{
			return View();
		}
	
		[HttpPost]
		[ValidateAntiForgeryToken]
		public IActionResult Index(Login model, string returnUrl)
		{

			if (!ModelState.IsValid)
			{
				return View(model);
			}

			try
			{
				string usrname = model.login_username.ToString();
				string usrpass = Security.Encryptor(model.login_password.Trim()).ToString();

				//var ip = Request.HttpContext.Connection.RemoteIpAddress;

				int cnt = 1;

				int intMaxLoginAttempts = Convert.ToInt16(Connect.ConfigurationManager.AppSetting["AppSecurity:Lock"]);

				bool uClosed = _db.UserClosed(Convert.ToInt16(Connect.ConfigurationManager.AppSetting["AppSecurity:Inactive"]), Convert.ToInt16(Connect.ConfigurationManager.AppSetting["AppSecurity:Unlock"]));

				bool tskUser = _db.ChkUser(usrname);

				if (tskUser)
				{
					bool tskPass = _db.ChkPass(usrname, usrpass);
					if (tskPass)
					{

						var tsk = _db.GetLoginID(usrname, usrpass);

						if (tsk != null)
						{
							string v = (tsk.Login.ToString() ?? "0");

							switch (tsk.chPass)
							{
								case "I":
									TempData["msgErr"] = "User ID is Inactive. Please contact administrator.";
									break;
								case "L":
									TempData["msgErr"] = "Your ID is locked, press UNLOCK to activate your account.";
									break;
								case "F":
								case "A":

									if (v == "0")
									{
										var ses = new SessionModel()
										{
											ID = tsk.ID,
											FullName = tsk.FullName,
											LoginID = tsk.LoginID,
											UType = tsk.GroupID,
											Dept = tsk.Dept,
											Stat = tsk.chPass,
											WrkStn = _db.GetIPAdd() + "/" + tsk.LoginID.ToString()
										};

										HttpContext.Session.SetString("SesObj", JsonConvert.SerializeObject(ses));

										HttpContext.Session.SetString("FullName", tsk.FullName);

										if (tsk.chPass == "F")
										{
											TempData["msgpass"] = "Please change your password immediately";
											//return RedirectToAction("PassChange", "Login");
										}

										if (tsk.ExpiryDate >= 1 && tsk.ExpiryDate <= 5)
										{
											TempData["msgExpire"] = "Your password will expire in " + tsk.ExpiryDate.ToString() + " days. Please change your password immediately.";
										}
										else if (tsk.ExpiryDate <= 0)
										{
											TempData["msgpass"] = "You Pasword has expired. Please change your password immediately";
										}

										bool LogIn = _db.LogUserOut(tsk.ID, "O");

										HttpContext.Session.SetString("Logged_IN", "Y");

										string wrkstn = _db.GetIPAdd() + "/" + tsk.LoginID.ToString();

										bool isLog = _db.GetAuditLogs("User Login", wrkstn, tsk.FullName + " has been successfully login to the system");

										if (!string.IsNullOrEmpty(returnUrl))
										{
											return LocalRedirect(returnUrl);
										}

										return RedirectToAction("Index", "Home");

									}
									else
									{
										HttpContext.Session.SetString("Unlak", model.login_username);
										TempData["msglag"] = "meron";
									}

									break;

							}
						}
					}
					else
					{
						var s = TempData["LoginCount"] ?? "0";

						if (Convert.ToInt16(s) >= intMaxLoginAttempts)
						{
							TempData.Remove("LoginCount");

							bool isLocked = _db.LockUser(usrname);

							if (isLocked)
							{
								TempData["msgErr"] = "User " + usrname + " are now locked in the system";
							}

						}
						else
						{

							TempData["LoginCount"] = Convert.ToInt32(TempData["LoginCount"]) + 1;

							cnt = intMaxLoginAttempts - Convert.ToInt32(TempData["LoginCount"]);

							TempData.Keep("LoginCount");

							//TempData["msginc"] = "You have " + cnt.ToString() + " more tries before your account will be lock into the system!";

							TempData["msgErr"] = "Invalid Username or Password";
						}

					}
				}
				else
				{
					TempData["msgErr"] = "Invalid User Name / Password";
				}

			}
			catch (Exception ex)
			{
				string err = ex.ToString();

				TempData["msgErr"] = err;
			}

			return View(model);
		}

		[HttpGet]
		public IActionResult ForgotPass()
		{
			return PartialView("_ForgotPass");

		}


		[HttpPost]
		public IActionResult ForgotPass(ForgotPass usr)
		{
			if (ModelState.IsValid)
			{
				var vusr = _db.GetEmail(usr.Username);

				if (vusr != null)
				{
					string pass = randompassword.Generate(8, 10);

					bool sendemail = _db.SendEmailForgot(vusr.Email, pass, vusr.FullName);

					if (sendemail)
					{	
						bool sndusr = _db.UpdateStatus(vusr.LoginID, Security.Encryptor(pass).ToString());

						if (sndusr)
						{
							
							string wrkstn = _db.GetIPAdd() + " / " + vusr.LoginID.ToString();

							bool isLog = _db.GetAuditLogs("Forgot Password", wrkstn, vusr.FullName + " has been successfully received temporary password from the system");


							//bool isLog = _db.InsertLog("Forgot Password", ip, "User " + vusr.usrId + " has requested for temporary password thru email");

							TempData["msgok"] = "Password has been reset. Please check your registered email.";
							return RedirectToAction("Index", "Login");
						}
					}
					else
					{
						TempData["msginfo"] = "Email sending was not successfully please change smtp client";
						return RedirectToAction("Index", "Login");
					}
				}
				else
				{
					TempData["msgusr"] = "Invalid username. Please enter registered user name.";
					return RedirectToAction("Index", "Login");
				}
			}

			return PartialView("_ForgotPass");

		}

		public IActionResult PassChange()
		{
			if (HttpContext.Session.GetString("Logged_IN") == "Y")
			{
				return View();
			}
			else
			{
				TempData["msginfo"] = "Please Login first before moving to this page";

				return RedirectToAction("Index", "Login");
			}
		}


		[HttpPost]
		public IActionResult PassChange(ResetPassword rest)
		{
			try
			{
				if (ModelState.IsValid)
				{
					var sesUser = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

					//Validate Old Password
					bool oldPass = _db.ChkPass(sesUser.LoginID, Security.Encryptor(rest.OldPassword).ToString());

					if (oldPass)
					{
						//Validate Temporary Password
						bool chktmp = _db.ChkTmpPass(sesUser.LoginID, Security.Encryptor(rest.Password.ToString()));

						if (chktmp)
						{
							//LOOK UP to table PASSLIST SOUNDEX
							bool chkp = _db.ChkPassList(rest.Password.ToString());
							if (chkp)
							{
								//LOOK UP to table PASSLIST SEQUENTIAL
								bool chkplst = _db.ChkPassListSeq(rest.Password.ToString());
								if (chkplst)
								{
									//LOOK UP if you use your Name or username
									bool chknam = _db.ChkPassName(sesUser.LoginID, rest.Password.ToString());

									if (chknam)
									{
										int pr = Convert.ToInt16(Connect.ConfigurationManager.AppSetting["AppSecurity:PassReuse"]);

										CountPass chkcnt = _db.ChkPassCount(sesUser.LoginID, Security.Encryptor(rest.Password).ToString());

										if (chkcnt.oldpass == 0)
										{
											bool sndusr = _db.UpdatePass(sesUser.LoginID, Security.Encryptor(rest.Password).ToString(), sesUser.UType);

											if (sndusr)
											{
												bool isLog = _db.GetAuditLogs("Change Password", sesUser.WrkStn, sesUser.FullName + " has been successfully changed his/her password");

												TempData["msgChange"] = "You will be logged out, please login again";

												return RedirectToAction("Logoff", "Login", new { id = 1 });
											}
											else
											{
												TempData["msgprocerr"] = "Server is busy, please try again later.";
											}
										}
										else
										{
											if (chkcnt.count_pass > pr)
											{
												bool sndusr = _db.UpdatePass(sesUser.LoginID, Security.Encryptor(rest.Password).ToString(), sesUser.UType);

												if (sndusr)
												{
													bool isLog = _db.GetAuditLogs("Change Password", sesUser.WrkStn, sesUser.FullName + " has been successfully changed his/her password");

													TempData["msgChange"] = "You will be logged out, please login again";


													return RedirectToAction("Logoff", "Login", new { id = 1 });
												}
												else
												{
													TempData["msgprocerr"] = "Server is busy, please try again later.";
												}
											}
											else
											{
												TempData["msgError"] = "Sorry you entered an old password, please enter new password";
											}
										}

									}
									else
									{
										TempData["msgError"] = "Invalid password, your name are not allowed to become the password.";
									}

								}
								else
								{
									TempData["msgError"] = "Invalid password, password does not comply with Baseline Security.";
								}
							}
							else
							{
								TempData["msgError"] = "Invalid password, password does not comply with Baseline Security.";
							}
						}
						else
						{
							TempData["msgError"] = "You have entered an Old Password. Please enter new password";
						}
					}
					else
					{
						TempData["msgError"] = "This is not your old password.";
					}

					return View(rest);
				}
			}
			catch (Exception ex)
			{

				string msg = ex.ToString();

				TempData["msgErr"] = msg;
			}

			return View(rest);
		}

		[HttpGet]
		public IActionResult UnlockUser()
		{

			ForgotPass usr = new ForgotPass();

			string unlak = HttpContext.Session.GetString("Unlak");

			if (unlak == string.Empty || unlak == null)
			{
				usr.Username = string.Empty;
			}
			else
			{
				usr.Username = unlak.ToString().ToLower();
			}

			return PartialView("_UnlockUser", usr);

		}

		[HttpPost]
		public IActionResult UnlockUser(ForgotPass usr)
		{
			if (ModelState.IsValid)
			{
				var vusr = _db.GetEmail(usr.Username);

				if (vusr != null)
				{

					string baseUrl = $"{this.Request.Scheme}://{this.Request.Host}{this.Request.PathBase}" + "/Login/Unlocking/" + vusr.ID;

					bool sendUlock = _db.SendEmailUnlock(vusr.FullName, vusr.Email, baseUrl.Replace("http", "https"), vusr.LoginID);

					if (sendUlock)
					{
						string wrkstn = _db.GetIPAdd() + " / " + vusr.LoginID.ToString();

						bool isLog = _db.GetAuditLogs("User Unlock", wrkstn, vusr.FullName + " has been successfully received email that will unlock him in the system");

						TempData["msgok"] = "Please check your email for verification.";
						return RedirectToAction("Index", "Login");
					}
					else
					{
						TempData["msgErr"] = "Sending email verification failed, please try again later.";
						return RedirectToAction("Index", "Login");

					}

					//bool ulock = _db.UserUnlock(usr.Username);

					//if (ulock)
					//{
					//	//string wrkstn = _db.GetIPAdd() + " / " + usr.Username.ToString();

					//	//bool isLog = _db.GetAuditLogs("Request Unlock", wrkstn, vusr.FIRSTNAME.ToString() + " " + vusr.LASTNAME.ToString() + " has been successfully unlocked from the system");

					//	TempData["msgok"] = "User successfully unlocked please proceed to login.";

					//	return RedirectToAction("Index", "Login");
					//}
					//else
					//{
					//	TempData["msgErr"] = "Error on processing please try again later";
					//	return RedirectToAction("Index", "Login");
					//}
				}
				else
				{
					TempData["msgErr"] = "Invalid username. Please enter registered user name.";
					return RedirectToAction("Index", "Login");
				}
			}

			return PartialView("_UnlockUser", usr);

		}

		[HttpGet]
		public IActionResult Unlocking(int id)
		{
			var vusr = _db.ViewUsers(id);

			bool ulock = _db.UserUnlock(vusr.loginid);

			if (ulock)
			{

				TempData["msgok"] = "User successfully unlocked. Please proceed to login.";

				return RedirectToAction("Index", "Login");
			}
			else
			{
				TempData["msgErr"] = "Error on processing please try again later";
				return RedirectToAction("Index", "Login");
			}


		}


		public IActionResult Logoff(int? id)
		{

			var sesUser = JsonConvert.DeserializeObject<SessionModel>(HttpContext.Session.GetString("SesObj"));

			bool uClosed = _db.UserClosed(Convert.ToInt16(Connect.ConfigurationManager.AppSetting["AppSecurity:Inactive"]), Convert.ToInt16(Connect.ConfigurationManager.AppSetting["AppSecurity:Unlock"]));

			bool isLogOut = _db.LogUserOut(sesUser.ID, "X");

			if (isLogOut)
			{
				//string wrkstn = _ipadd.HttpContext.Connection.RemoteIpAddress.ToString() + " / " + sesUser.LoginID.ToString();

				//bool isLog = _db.GetAuditLogs("User Logout", wrkstn, sesUser.FullName + " successfully logout in the system");

				bool isLog = _db.GetAuditLogs("User Logout", sesUser.WrkStn, sesUser.FullName + " successfully logout in the system");

				HttpContext.Session.Clear();
			}

			if (id != null)
			{
				TempData["msgok"] = "Successfully changed your password, please log-in to continue";
			}
			else
			{
				TempData["msgok"] = "Successfully log-out in the system.";
			}


			return RedirectToAction("Index", "Login");
		}

	}
}
