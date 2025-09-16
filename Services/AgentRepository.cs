using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using AgentDesktop.Models;
using AgentDesktop.Services;
using System.Net.Mail;
using Microsoft.Data.SqlClient;
using Microsoft.AspNetCore.Http.Extensions;
using System.Security.Policy;
using Microsoft.AspNetCore.Http;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Net;
using System.Net.Sockets;

namespace AgentDesktop.Services
{
	public class AgentRepository : AgentInterface
	{
        //private IDbConnection db1 = ConnectionManager.GetConnection("B");

        private IDbConnection db = new SqlConnection(new SqlConnectionStringBuilder()
		{
			DataSource = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppDBConnection:SQLNM"].ToString()),
			InitialCatalog = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppDBConnection:DBNM"].ToString()),
			PersistSecurityInfo = true,
			UserID = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppDBConnection:LGID"].ToString()),
			Password = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppDBConnection:PAWD"].ToString()),
			MultipleActiveResultSets = true
        }.ConnectionString);

        private IConfiguration _configuration;
        private IWebHostEnvironment _environment;
        private IHttpContextAccessor _httpContextAccessor;

        public AgentRepository(IConfiguration configuration, IWebHostEnvironment environment, IHttpContextAccessor httpContextAccessor)
        {
            _configuration = configuration;
            _environment = environment;
            _httpContextAccessor = httpContextAccessor;
        }

        #region Login
        public bool ChkUser(string _usr)
		{
			int cnt = this.db.Query<chkLogin>("Select LoginID username, [Password] from UserName where loginid = '" + _usr + "'", new { }).Count();

			if (cnt == 0)
			{
				return false;
			}
			else
			{
				return true;
			}
		}

		public EmailForgot GetEmail(string _usr)
		{
			return db.Query<EmailForgot>("Select ID, FullName, LoginID, Email from UserName where loginid = '" + _usr + "'", new { }).FirstOrDefault();
		}

        public bool SendEmailForgot(string _email, string _pass, string _fname)
        {
            try
            {
                string mailserver = Connect.ConfigurationManager.AppSetting["AppEmail:MailServer"];
                string mailserverport = Connect.ConfigurationManager.AppSetting["AppEmail:Port"];
                string mailuser = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppEmail:EUser"]);
                string mailpass = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppEmail:EPass"]);

                System.Net.NetworkCredential credentials = new(mailuser, mailpass);
                SmtpClient client = new SmtpClient(mailserver, Convert.ToInt16(mailserverport));


                client.UseDefaultCredentials = false;

                client.Credentials = credentials;

                client.EnableSsl = false;

                System.Collections.ArrayList emailArray = new System.Collections.ArrayList();

                MailMessage message = new MailMessage();
                MailAddress toAddress = new MailAddress(_email);

                message.From = new MailAddress("CUSTOMER.MANAGEMENT@sterlingbankasia.com", "Customer Management");
                message.To.Add(toAddress);
                message.Subject = "Forgot Password in the New AgentDesktop System";
                message.IsBodyHtml = true;

                string mbody = "<p>Hi " + _fname + ", <br/><br/> " +
                                "Your password has been reset. <br/><br/> " +
                                "Please log into your account with the following temporary password: <br/> " +
                                "<h3>" + _pass + "</h3><br/><br/>" +
                                "After logging in with this temporary password, please visit your account " +
                                "settings page and change your password to something secure that you'll remember. <br/><br/>" +
                                "Thank you and best regards. <br/><br/><br/> " +
                                "Customer Support <br/><br/><br/>" +
                                "This e-mail was sent using the new AgentDesktop please do not reply.</p>";


                message.Body = mbody;

                client.Send(message);

                return true;
            }
            catch (Exception ex)
            {
                string x = ex.Message;
                return false;
            }

        }

        public bool UpdateStatus(string _userid, string _pass)
        {
            try
            {

                string sqlQuery = "sp_UpdateStatus";

                db.Execute(sqlQuery, new
                {
                    UserID = _userid,
                    Pass = _pass
                }, commandTimeout: 0, commandType: CommandType.StoredProcedure);


                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ChkPass(string _usr, string _pass)
        {
            int cnt = this.db.Query<chkLogin>("select loginid username, Password from UserName WHERE LOWER(loginid) = '" + _usr.ToLower() + "' and password = '" + _pass + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public LoginPass GetLoginID(string usrName, string usrPass)
        {
            return this.db.Query<LoginPass>("select ID, FullName, LoginID, Active, GroupID, Email, Dept, isnull([Login], 0)[Login], chPass, Password, " +
                "convert(int, datediff(day, GETDATE(), isnull(ExpiryDate, getdate()))) ExpiryDate " +
                "from UserName WHERE LOWER(LoginID) = '" + usrName.ToLower() + "' and Password = '" + usrPass + "'", new { }).FirstOrDefault();
        }

        public bool LockUser(string _username)
        {
            try
            {
                string sqlQuery = "update username set chPass = 'L' where Lower(LoginID) = '" + _username.ToLower() + "'";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ChkTmpPass(string _userid, string _pass)
        {
            int cnt = this.db.Query<PassList>("select [password] from UserMembership where LoginID = '" + _userid + "' and [password] = '" + _pass + "' ", new { }).Count();

            if (cnt != 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool ChkPassList(string _pass)
        {
            int cnt = this.db.Query<PassList>("select pass from PASSLIST where soundex(pass) = SOUNDEX('" + _pass + "')", new { }).Count();

            if (cnt != 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool ChkPassListSeq(string _pass)
        {
            int cnt = this.db.Query<PassList>("sp_CountPassList", new { pass = _pass }, commandTimeout: 0, commandType: CommandType.StoredProcedure).Count();

            if (cnt != 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }


        public bool ChkPassName(string _id, string _pass)
        {

            int cnt = this.db.Query<PassList>("sp_CountPassUser", new { loginid = _id, pass = _pass }, commandTimeout: 0, commandType: CommandType.StoredProcedure).Count();


            //int cnt = this.db.Query<PassList>("SELECT pass FROM ( " +
            //    "select FullName pass from UserName where LoginID = '" + _id + "' " +
            //    "union all " +
            //    "select LoginID from UserName where LoginID = '" + _id + "' " +
            //    "union all " +
            //    "select Email from UserName where LoginID = '" + _id + "' " +
            //    ") a where soundex(a.pass) like SOUNDEX('" + _pass + "')", new { }).Count();

            if (cnt != 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public CountPass ChkPassCount(string _id, string _pass)
        {

            string sqlQuery = "sp_CountPass";

            return this.db.Query<CountPass>(sqlQuery, new { UserID = _id, pass = _pass }, commandTimeout: 0, commandType: CommandType.StoredProcedure).FirstOrDefault();

        }


        public bool UpdatePass(string _username, string _pass, int _usrtyp)
        {
            try
            {
                string sqlQuery = "sp_UpdatePassword";

                db.Execute(sqlQuery, new
                {
                    UserId = _username,
                    Pass = _pass,
                    UserType = _usrtyp
                }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool UserClosed(int _inactive, int _unlock)
        {
            try
            {

                //string sqlQuery = "UPDATE sba_Employees set LOGGED = 'N' where USERID in ( " +
                //	"select a.UserId from sba_users a inner join sba_Employees b on a.UserId = b.USERID " +
                //	"where b.LOGGED = 'Y' AND DATEDIFF(mi, LastActivityDate, getdate()) >= " + _unlock + ")";
                //db.Execute(sqlQuery, new { });

                string sqlQuery = "sp_Inactive";

                db.Execute(sqlQuery, new
                {
                    Inactive = _inactive,
                    Unlock = _unlock
                }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool LogUserOut(int _userid, string _mod)
        {
            try
            {

                //string sqlQuery = "UPDATE sba_Users SET LastActivityDate = getdate() WHERE UserName = '" + _loginid + "'";
                //db.Execute(sqlQuery, new { });

                string sqlQuery = "sp_LogINOUT";

                db.Execute(sqlQuery, new
                {
                    ID = _userid,
                    MODE = _mod
                }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool Activity(string _loginid)
        {
            try
            {

                string sqlQuery = "UPDATE username SET lastactivity = getdate(), login = 0 WHERE loginid = '" + _loginid + "'";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool SendEmailUnlock(string _name, string _email, string _url, string _usrname)
        {
            try
            {

                string mailserver = Connect.ConfigurationManager.AppSetting["AppEmail:MailServer"];
                string mailserverport = Connect.ConfigurationManager.AppSetting["AppEmail:Port"];
                string mailuser = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppEmail:EUser"]);
                string mailpass = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppEmail:EPass"]);

                System.Net.NetworkCredential credentials = new(mailuser, mailpass);
                SmtpClient client = new SmtpClient(mailserver, Convert.ToInt16(mailserverport));

                client.UseDefaultCredentials = false;

                client.Credentials = credentials;

                client.EnableSsl = false;

                System.Collections.ArrayList emailArray = new System.Collections.ArrayList();

                var callbackUrl = _url;

                MailMessage message = new MailMessage();
                MailAddress toAddress = new MailAddress(_email);

                string mailto = toAddress.ToString();

                message.From = new MailAddress("CUSTOMER.MANAGEMENT@sterlingbankasia.com", "Customer Management");

                message.To.Add(toAddress);


                message.Subject = "User Unlock Verification for the new AgentDesktop";
                message.IsBodyHtml = true;

                string mbody = "<p>Hi " + _name + ", <br/><br/> " +
                                "Your userid " + _usrname + " has been unlocked, please click <a href=\"" + callbackUrl + "\">here</a> to login the new AgentDesktop System.<br/><br/> " +
                                "Thank you and best regards. <br/><br/><br/> " +
                                "Customer Support <br/><br/><br/>" +
                                "This e-mail was sent using the new AgentDesktop System please do not reply.</p>";

                message.Body = mbody;

                client.Send(message);

                return true;
            }
            catch (Exception ex)
            {
                string x = ex.Message;
                return false;
            }

        }

        public bool UserUnlock(string _username)
        {
            try
            {
                string sqlQuery = "update UserName set chPass = 'A', login = 0, lastactivity = getdate() where Lower(loginid) = '" + _username.ToLower() + "'";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public LoginPass GetAgent(int id)
        {
            return this.db.Query<LoginPass>("select ID, LoginID from UserName WHERE id = " + id + "", new { }).FirstOrDefault();
        }

        #endregion

        #region Admin
        public IEnumerable<ViewUsers> ListUsers()
        {
            return this.db.Query<ViewUsers>("select a.id, loginid, fullname, groupdescription [group], email, department, " +
                "case when chPass = 'A' then 'Active' when chPass = 'L' then 'Locked' when chPass = 'F' then 'Forgot Password' else 'Inactive' end stat, " +
                "convert(varchar, isnull(ExpiryDate, getdate()), 1) expirydate " + 
                "from UserName a left join GroupName b on a.groupid = b.groupid left join dept c on a.dept = c.dept", new { });
        }

        public List<GroupName> ListGroup()
        {
            return this.db.Query<GroupName>("select GroupID, GroupDescription from GroupName order by GroupID", new { }).ToList();
        }


        public List<Departments> ListDept()
        {
            return this.db.Query<Departments>("select Dept, Department from Dept order by did", new { }).ToList();
        }

        public bool CreateUser(NewUser usr)
        {
            try
            {
                string expdt = usr.group == 1 ? DateTime.Now.AddDays(180).ToString("dd-MMM-yyyy") : DateTime.Now.AddDays(90).ToString("dd-MMM-yyyy");

                string randompass = Security.Encryptor(randompassword.Generate(8, 10));

                string USRPASS = randompass;//Connect.ConfigurationManager.AppSetting["AppSecurity:DefaultPass"];

                string sqlQuery = "INSERT INTO UserName (FullName,LoginID,GroupID,Email,Dept,chPass,Password,ExpiryDate) " +
                    "VALUES('" + usr.fullname + "', '" + usr.loginid + "', " + usr.group + ", '" + usr.email + "', '" + usr.department + "', 'F', '" + USRPASS + "', Convert(date, '" + expdt + "'))";

                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception ex)
            {
                string x = ex.ToString();
                return false;
            }
        }

        public UserPass GetUserPass(string _usr)
        {
            return db.Query<UserPass>("Select LoginID, Password from UserName where loginid = '" + _usr + "'", new { }).FirstOrDefault();
        }

        public bool SendEmailNewUser(string _usrid, string _email, string _fname, string _pass)
        {
            try
            {
                string mailserver = Connect.ConfigurationManager.AppSetting["AppEmail:MailServer"];
                string mailserverport = Connect.ConfigurationManager.AppSetting["AppEmail:Port"];
                string mailuser = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppEmail:EUser"]);
                string mailpass = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppEmail:EPass"]);

                System.Net.NetworkCredential credentials = new(mailuser, mailpass);
                SmtpClient client = new SmtpClient(mailserver, Convert.ToInt16(mailserverport));


                client.UseDefaultCredentials = false;

                client.Credentials = credentials;

                client.EnableSsl = false;

                System.Collections.ArrayList emailArray = new System.Collections.ArrayList();

                MailMessage message = new MailMessage();
                MailAddress toAddress = new MailAddress(_email);

                message.From = new MailAddress("CUSTOMER.MANAGEMENT@sterlingbankasia.com", "Customer Management");
                message.To.Add(toAddress);
                message.Subject = "New User in the New AgentDesktop System";
                message.IsBodyHtml = true;

                string mbody = "<p>Hi " + _fname + ", <br/><br/> " +
                                "You are now registered in the new AgentDesktop System. <br/><br/> " +
                                "Please log-in with the following credentials: <br/> " +
                                "<h3>User Name :" + _usrid + "</h3>" +
                                "<h3>Password : " + Security.DeCryptor(_pass) + "</h3><br/><br/>" +
                                "After logging in with this temporary password, please change your password immediately.<br/><br/>" +
                                "Thank you and best regards. <br/><br/><br/> " +
                                "Customer Support <br/><br/><br/>" +
                                "This e-mail was sent using the new AgentDesktop please do not reply.</p>";


                message.Body = mbody;

                client.Send(message);

                return true;
            }
            catch (Exception ex)
            {
                string x = ex.Message;
                return false;
            }

        }

        public EditUser ViewUsers(int _id)
        {
            string sql = "select id, fullName, loginid, groupid [group], email, dept department, chpass from UserName where id = " + _id + "";

            return this.db.Query<EditUser>(sql, new { }).FirstOrDefault();
        }

        public bool UpdateUser(EditUser edt)
        {
            try
            {
                string sqlQuery = "UPDATE UserName SET fullname = '" + edt.fullname + "', groupid = " + edt.group + ", email = '" + edt.email + "', dept = '" + edt.department + "', chpass = '" + edt.chpass + "' " +
                    "WHERE ID = '" + edt.id + "'";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool DeleteUser(int _id)
        {
            try
            {
                string sqlQuery = "DELETE FROM UserName WHERE ID = " + _id + "";
                db.Execute(sqlQuery, new { });


                string sqlQuery2 = "delete from UserRights where LoginID = " + _id + "";
                db.Execute(sqlQuery2, new { });

                string sqlQuery3 = "delete from UserMembership where LoginID in ( " +
                    "select LoginID from UserName where id = " + _id + ")";
                db.Execute(sqlQuery3, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public IEnumerable<ReasonCode> ListReasonCodes()
        {
            return this.db.Query<ReasonCode>("select RNID, RCode, Reason, Description, FirstLetter, SecondLetter, ThirdLetter, Abbreviation, Class from reasonnew", new { });
        }
        public bool ChkRCode(string _rcd)
        {
            int cnt = this.db.Query<chkLogin>("select * from ReasonNew where rcode = '" + _rcd + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool CreateReasonCode(NewReason rcd)
        {
            try
            {

                string sqlQuery = "INSERT INTO ReasonNew (RCode,Reason,Description,FirstLetter,SecondLetter,ThirdLetter,Abbreviation,Class) " +
                    "VALUES('" + rcd.RCode + "','" + rcd.Reason + "','" + rcd.Description + "', '" + rcd.FirstLetter + "','" + rcd.SecondLetter + "','" + rcd.ThirdLetter + "','" + rcd.Abbreviation + "', '" + rcd.Class + "')";

                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public EditReason ViewReasonCode(int _id)
        {
            string sql = "select RNID, RCode,Reason,Description,FirstLetter,SecondLetter,ThirdLetter,Abbreviation, Class from ReasonNew where rnid = " + _id + "";

            return this.db.Query<EditReason>(sql, new { }).FirstOrDefault();
        }

        public bool UpdateReasonCode(EditReason edt)
        {
            try
            {
                string sqlQuery = "UPDATE ReasonNew SET RCode = '" + edt.RCode + "',Reason = '" + edt.Reason + "',Description = '" + edt.Description + "',FirstLetter = '" + edt.FirstLetter + "',SecondLetter = '" + edt.SecondLetter + "',ThirdLetter = '" + edt.ThirdLetter + "',Abbreviation = '" + edt.Abbreviation + "',Class = '" + edt.Class + "' WHERE RNID = " + edt.RNID + "";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ChkRCodeActivity(string _rcd)
        {
            int cnt = this.db.Query<CallReport>("select * from CallReport where ReasonCode = '" + _rcd + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool DeleteReason(int _id)
        {
            try
            {
                string sqlQuery = "DELETE FROM ReasonNew WHERE RNID = '" + _id + "'";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public IEnumerable<ATM> ListATM(string _typ)
        {
            return this.db.Query<ATM>("select ID, Description, DesType from ATM where DesType = '" + _typ + "'", new { });
        }

        public bool CreateATM(string _desc, string _typ)
        {
            try
            {

                string sqlQuery = "INSERT INTO ATM (Description, DesType) " +
                    "VALUES('" + _desc + "','" + _typ + "')";

                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ChkATM(string _desc, string _typ)
        {
            int cnt = this.db.Query<ATM>("select * from ATM where Description = '" + _desc + "' and DesType = '" + _typ + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public EditATM ViewATM(int _id)
        {
            string sql = "select ID, Description from ATM where id = " + _id + "";

            return this.db.Query<EditATM>(sql, new { }).FirstOrDefault();
        }

        public bool UpdateATM(EditATM edt)
        {
            try
            {
                string sqlQuery = "UPDATE ATM SET Description = '" + edt.Description + "' WHERE ID = " + edt.ID + "";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ChkPaymentOfActivity(string _desc)
        {
            int cnt = this.db.Query<CallReport>("select * from CallReport where PaymentOf = '" + _desc + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool DeleteATM(int _id)
        {
            try
            {
                string sqlQuery = "DELETE FROM ATM WHERE ID = '" + _id + "'";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool DeleteATMLocation(int _id, int _loc_id)
        {
            try
            {
                string sqlQuery = "delete from ATM_LOCATION where ATM_LOC_ID = " + _loc_id + "";
                db.Execute(sqlQuery, new { });

                string sqlQuery2 = "delete from ATM_GROUP_LOC where id = " + _id + "";
                db.Execute(sqlQuery2, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ChkBancnetActivity(string _desc)
        {
            int cnt = this.db.Query<CallReport>("select * from CallReport where BancnetUsed = '" + _desc + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool ChkSubStatusActivity(string _desc)
        {
            int cnt = this.db.Query<CallReport>("select * from CallReport where SubStatus = '" + _desc + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool ChkRemarksActivity(string _desc)
        {
            int cnt = this.db.Query<CallReport>("select * from CallReport where Remarks = '" + _desc + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool ChkReferredByActivity(string _desc)
        {
            int cnt = this.db.Query<CallReport>("select * from CallReport where ReferedBy = '" + _desc + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }


        public bool ChkATMActivity(string _desc)
        {
            int cnt = this.db.Query<CallReport>("select * from CallReport where ATMUsed = '" + _desc + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public IEnumerable<ATM_GROUP_LOC> ListATMLocation()
        {
            return this.db.Query<ATM_GROUP_LOC>("select b.ID, a.DESCRIPTION ATM, c.ATM_LOCATION  from ATM a left join ATM_GROUP_LOC b on a.id = b.atm_id " +
                "inner join ATM_LOCATION c on b.atm_loc_id = c.ATM_LOC_ID where a.DesType = 'AU'", new { });
        }

        public bool CreateATMLocation(int _id, string _desc)
        {
            try
            {

                string sqlQuery = "sp_InsertATMLocation";

                db.Execute(sqlQuery, new
                {
                    ATM_ID = _id,
                    ATM_LOC = _desc
                }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

                return true;
            }
            catch (Exception ex)
            {
                string msg = ex.ToString();
                return false;
            }
        }

        public EditATMLoc ViewATMLocation(int _id)
        {
            string sql = "select a.ID, a.ATM_ID, c.ATM_LOC_ID, c.ATM_LOCATION from ATM_GROUP_LOC a left join ATM b on a.ATM_ID = b.ID " +
                "left join ATM_LOCATION c on a.ATM_LOC_ID = c.ATM_LOC_ID where a.id = " + _id + "";

            return this.db.Query<EditATMLoc>(sql, new { }).FirstOrDefault();
        }

        public bool UpdateATMLocation(string _atmloc, int _locid)
        {
            try
            {
                string sqlQuery = "UPDATE ATM_LOCATION SET ATM_LOCATION = '" + _atmloc + "' WHERE ATM_LOC_ID = " + _locid + "";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public IEnumerable<UserMenu> ListUserMenu()
        {
            return this.db.Query<UserMenu>("select ID, [View], [Page], Title from UserMenu", new { });
        }

        public bool ChkMenu(string _view, string _page)
        {
            int cnt = this.db.Query<UserMenu>("select * from UserMenu where [View] = '" + _view + "' and [Page] = '" + _page + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool CreateMenu(NewMenu mnu)
        {
            try
            {

                string sqlQuery = "INSERT INTO UserMenu ([View], [Page], Title) " +
                    "VALUES('" + mnu.View + "','" + mnu.Page + "', '" + mnu.Title + "')";

                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public EditMenu ViewMenu(int _id)
        {
            string sql = "select ID, [View], [Page], Title from UserMenu where id = " + _id + "";

            return this.db.Query<EditMenu>(sql, new { }).FirstOrDefault();
        }

        public bool UpdateMenu(EditMenu edt)
        {
            try
            {
                string sqlQuery = "UPDATE UserMenu SET [Page] = '" + edt.Page.Trim() + "', Title = '" + edt.Title.Trim() + "' WHERE ID = " + edt.ID + "";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool DeleteMenu(int _id)
        {
            try
            {
                string sqlQuery = "DELETE FROM USERMENU WHERE ID = '" + _id + "'";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ChkEndorsedActivity(string _desc)
        {
            int cnt = this.db.Query<CallReport>("select * from CallReport where EndorsedFrom = '" + _desc + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public IEnumerable<UserRight> ListUserRights(int id)
        {
            return this.db.Query<UserRight>("select a.ID, title, Rights from UserRightsNew a inner join UserMenu b on a.MenuID = b.ID " +
                "where a.USRID = @USRID", new { USRID = id });
        }

        public bool InsertRights(int _id)
        {
            try
            {

                db.Execute("sp_InsertRightsNew", new { USRID = _id }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool UpdateRights(int _id, string _idd)
        {
            try
            {

                db.Execute("sp_UpdateRightsNew", new { USRID = _id, IDD = _idd }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ChkRights(string _uid, string _vw, string _pg)
        {
            int cnt = this.db.Query<UserRight>("select a.ID, b.Title, a.Rights from UserRightsNew a inner join UserMenu b on a.MenuID = b.ID " +
                "inner join UserName c on a.USRID = c.ID where c.LoginID = '" + _uid + "' and[View] = '" + _vw + "' and[Page] = '" + _pg + "' and Rights = 1", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool GetAuditLogs(string _description, string _workstn, string _activity)
        {
            try
            {
                string sqlQuery = "GetAuditLogs";

                db.Execute(sqlQuery, new { description = _description, workstn = _workstn, activity = _activity }, commandTimeout: 0, commandType: CommandType.StoredProcedure);


                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public IEnumerable<Audit_Logs> ListAuditLogs(string _dt1, string _dt2)
        {
            return this.db.Query<Audit_Logs>("select * from Audit_Logs where convert(date, logdate) between @Date1 and @Date2", new { Date1 = _dt1, Date2 = _dt2 });
        }

        #endregion

        #region Home
        public bool Populate_Dashboard(string _id)
        {
            try
            {

                db.Execute("sp_DashBoard2", new { LoginID = _id }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public IEnumerable<ServiceTicket> ListServiceTicket(string _loginid, string _dep)
        {

			string sql;
			//if (_dep == "M")
			//{
                sql = "select TID, description, cnt from ServiceTicket where LoginID = '" + _loginid + "'";

   //         }
   //         else
			//{
   //             sql = "select TID, description, cnt from ServiceTicketATM where LoginID = '" + _loginid + "'";
   //         }

            return this.db.Query<ServiceTicket>(sql, new { });
        }

        public IEnumerable<ServiceTicket> ListServiceStatus(string _loginid, string _dep)
        {

            string sql;
            //if (_dep == "M")
            //{
                sql = "select sid tid, status description, " +
                "case when status = 'Commitment' then isnull((select stscnt from ServiceStatus where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
                "     when status = 'Endorsed' then isnull((select stscnt from ServiceStatus where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
                "     when status = 'Escalated' then isnull((select stscnt from ServiceStatus where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
                "	 when status = 'Resolved' then isnull((select stscnt from ServiceStatus where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
                "	 when status = 'Closed' then isnull((select stscnt from ServiceStatus where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
                "	 when status = 'Cancelled' then isnull((select stscnt from ServiceStatus where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
                "else 0 end cnt from Status a";

            //}
            //else
            //{
            //    sql = "select sid tid, status description, " +
            //        "case when status = 'Commitment' then isnull((select stscnt from ServiceStatusATM where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
            //        "     when status = 'Endorsed' then isnull((select stscnt from ServiceStatusATM where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
            //        "     when status = 'Escalated' then isnull((select stscnt from ServiceStatusATM where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
            //        "	 when status = 'Resolved' then isnull((select stscnt from ServiceStatusATM where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
            //        "	 when status = 'Closed' then isnull((select stscnt from ServiceStatusATM where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
            //        "	 when status = 'Cancelled' then isnull((select stscnt from ServiceStatusATM where LoginID = '" + _loginid + "' and status = a.Status), 0) " +
            //        "else 0 end cnt from Status a";
            //}

            return this.db.Query<ServiceTicket>(sql, new { });
        }

        public IEnumerable<Escalation> ListEscalation()
		{
            return this.db.Query<Escalation>("sp_Escalation", new { }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<ServiceReason> ListServiceReason(string _loginid, string _dep)
        {

            string sql;
            //if (_dep == "M")
            //{
                sql = "select a.rnid, a.Description, a.inbound, a.outbound, b.Description title, b.RCode from ServiceReason a left join reasonnew b on substring(a.Description, 1, 3) = b.rcode " +
                    "where LoginID = '" + _loginid + "'";
            //}
            //else
            //{
            //    sql = "select a.rnid, a.Description, a.inbound, a.outbound, b.Description title, b.RCode from ServiceReasonATM a left join reasonnew b on substring(a.Description, 1, 3) = b.rcode " +
            //        "where LoginID = '" + _loginid + "'";
            //}

            return this.db.Query<ServiceReason>(sql, new { });
        }

        public IEnumerable<ServiceChannel> ListServiceChannel(string _loginid, string _dep)
        {

            string sql;
            //if (_dep == "M")
            //{
                sql = "select a.cid, a.Description, inbound, outbound, case " +
                    "	when b.Description = 'Telephone' then 'si si-call-in fa-2x' " +
                    "   when b.Description = 'Email' then 'si si-envelope fa-2x' " +
                    "   when b.Description = 'Fax' then 'fa fa-fax fa-2x' " +
                    "   when b.Description = 'SMS' then 'si si-speech fa-2x' " +
                    "   when b.Description = 'Voicemail' then 'si si-volume-2 fa-2x' " +
                    "   when b.Description = 'Mail' then 'si si-envelope fa-2x' " +
                    "else 'si si-users fa-2x' end icon " +
                    "from ServiceCommChannel a left join ComChannel b on a.Description = b.Description where LoginID = '" + _loginid + "'";
            //}
            //else
            //{
            //    sql = "select a.cid, a.Description, inbound, outbound, case " +
            //        "	when b.Description = 'Telephone' then 'si si-call-in fa-2x' " +
            //        "   when b.Description = 'Email' then 'si si-envelope fa-2x' " +
            //        "   when b.Description = 'Fax' then 'fa fa-fax fa-2x' " +
            //        "   when b.Description = 'SMS' then 'si si-speech fa-2x' " +
            //        "   when b.Description = 'Voicemail' then 'si si-volume-2 fa-2x' " +
            //        "   when b.Description = 'Mail' then 'si si-envelope fa-2x' " +
            //        "else 'si si-users fa-2x' end icon " +
            //        "from ServiceCommChannelATM a left join ComChannel b on a.Description = b.Description where LoginID = '" + _loginid + "'";
            //}

            return this.db.Query<ServiceChannel>(sql, new { });
        }

        public IEnumerable<DisplayServiceTicket> DisplayServiceTickets(string _loginid, string _dept, int _utype, int _cond )
        {
            return this.db.Query<DisplayServiceTicket>("sp_DBServiceTicket_New", new { LoginID = _loginid, Dept = _dept, Grp = _utype, Cond = _cond }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

#nullable enable
        public IEnumerable<DisplayServiceTicketDated> View_TransStatus(string _LoginID, string _Stat, int _Mod, string? _Date1, string? _Date2)
        {
            return this.db.Query<DisplayServiceTicketDated>("sp_DBServiceStatus", new { LoginID = _LoginID, Stat = _Stat, Mod = _Mod, Date1 = _Date1, Date2 = _Date2 }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

        }

        public IEnumerable<DisplayServiceTicketDated> View_TransChannel(string _LoginID, string _Stat, int _Mod, string? _Date1, string? _Date2)
        {
            return this.db.Query<DisplayServiceTicketDated>("sp_DBServiceChannel", new { LoginID = _LoginID, Channel = _Stat, Mod = _Mod, Date1 = _Date1, Date2 = _Date2 }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

        }

        public IEnumerable<DisplayServiceTicketDated> View_TransReason(string _LoginID, string _Stat, int _Mod, string? _Date1, string? _Date2)
        {
            return this.db.Query<DisplayServiceTicketDated>("sp_DBServiceReason", new { LoginID = _LoginID, RCode = _Stat, Mod = _Mod, Date1 = _Date1, Date2 = _Date2 }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

        }

        public IEnumerable<DisplayServiceTicket> View_TransEscalate(int _cond)
        {
            return this.db.Query<DisplayServiceTicket>("sp_DBEscalation", new { Cond = _cond }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        #endregion

        #region Activity
        public IEnumerable<CMActivity> View_CMActivity(ActivitySearch _sch)
        {

            //string _rcode = _rno.Replace(",", "','") ;

            string dt1 = (_sch.Date1 == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : _sch.Date1;

            string dt2 = (_sch.Date2 == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : _sch.Date2;

            string yr = DateTime.Now.Year.ToString();


            string strSQL;

            strSQL = "select RID, TicketNo, InOutBound, CommChannel, DateReceived, abbreviation, CustomerName, " +
                     "CardNo, Destination, TransType, ATMTrans, case when isnumeric(Location) = 0 then Location else dbo.Get_ATMLocation(ATMUsed, Location) end Location,  " +
                     "case when isnumeric(ATMUsed) = 0 then ATMUsed else dbo.Get_ATMUsed(ATMUsed) end ATMUsed,  PaymentOf, TerminalUsed, BillerName, " +
                     "Merchant, Inter, BancnetUsed, Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, " +
                     "PLStatus, CardBlocked, EndoredTo, Activity, Status, Agent, created_by, Last_Update, EndorsedFrom, Resolved, Remarks, " +
                     "ReferedBy, ContactPerson, ReasonCode, Tagging, ResolvedDate, CallTime, LocalNum, CardPresent, SubStatus, Currency, brnnam Branch from ( " +
                     "select a.rid, TicketNo, InOutBound, CommChannel, " +
                     "dbo.Get_DateReceived(a.TicketNo) DateReceived, " +
                     "a.reasoncode + ' - ' + abbreviation abbreviation, CustomerName, " +
                     "CardNo, Destination, TransType, ATMTrans, Location, ATMUsed, PaymentOf, TerminalUsed, BillerName, Merchant, Inter, BancnetUsed, " +
                     "Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, PLStatus, CardBlocked, EndoredTo, Activity = rtrim(Activity), " +
                     "Status, Agent, EndorsedFrom, resolved, Remarks, ReferedBy, ContactPerson, a.reasoncode, " +
                     "a.Tagging, created_by, last_update, " +
                     "dbo.Get_ResovedDate(a.TicketNo) ResolvedDate, CallTime, LocalNum, CardPresent, SubStatus, Currency, brnnam " +
                     "from CallReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
                     "left join Branch c on a.branch = c.brncd " +
                     "where type = 'Y' ";

            if (_sch.DateSearch == false)
            {
                strSQL += "and convert(date, datereceived) between convert(date, '" + dt1 + "') and convert(date, '" + dt2 + "') ";
            }
            else
            {
                strSQL += "and year(DateReceived) between '2019' and  '" + yr + "' ";
            }

            if (_sch.CustomerName != null)
			{
                strSQL += "and CustomerName like '%" + _sch.CustomerName.Trim() + "%' ";

			}

            if (_sch.CardNo != null)
            {
                strSQL += "and CardNo like '%" + _sch.CardNo.Trim() + "%' ";
            }

            if (_sch.CommType != "0")
            {
                if (_sch.CommType != null)
                {
                    strSQL += "and InOutBound = '" + _sch.CommType + "' ";
                }
            }

            if (_sch.CommChannel != "0")
            {
                if (_sch.CommChannel != null)
                {
                    strSQL += "and CommChannel = '" + _sch.CommChannel + "' ";
                }
            }

            if (_sch.ReasonCode != null)
            {
                strSQL += "and ReasonCode = '" + _sch.ReasonCode + "' ";
            }

            if (_sch.Status != "0")
            {
                if (_sch.Status != null)
                {
                    strSQL += "and Status = '" + _sch.Status + "' ";
                }
            }

            if (_sch.Agent != null)
            {
                strSQL += "and Created_By = '" + _sch.Agent + "' ";
            }

            if (_sch.TicketNo != null)
            {
                strSQL += "and TicketNo like '%" + _sch.TicketNo.Trim() + "%' ";
            }

            if (_sch.ReferredBy != null)
            {
                strSQL += "and ReferedBy like '%" + _sch.ReferredBy.Trim() + "%' ";
            }

            if (_sch.ContactPerson != null)
            {
                strSQL += "and ContactPerson like '%" + _sch.ContactPerson.Trim() + "%' ";
            }

            strSQL += ") x order by x.DateReceived, TicketNo";

            return this.db.Query<CMActivity>(strSQL, new { }, commandTimeout: 0);

            //return this.db.Query<CMActivity>("sp_CMActivity", new { Date1 = _dt1, Date2 = _dt2, Mod = _mod }, commandType: CommandType.StoredProcedure);

        }

        public IEnumerable<CMActivity> View_CMActivityHist(ActivitySearch _sch)
        {
            //string _rcode = _rno.Replace(",", "','");

            string dt1 = (_sch.Date1 == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : _sch.Date1;

            string dt2 = (_sch.Date2 == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : _sch.Date2;

            string yr = DateTime.Now.Year.ToString();

            string strSQL;

            strSQL = "select RID, TicketNo, InOutBound, CommChannel, DateReceived, abbreviation, CustomerName, " +
                     "CardNo, Destination, TransType, ATMTrans, Location, ATMUsed, PaymentOf, TerminalUsed, BillerName, " +
                     "Merchant, Inter, BancnetUsed, Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, " +
                     "PLStatus, CardBlocked, EndoredTo, Activity, Status, Agent, created_by, Last_Update, EndorsedFrom, Resolved, Remarks, " +
                     "ReferedBy, ContactPerson, ReasonCode, Tagging, ResolvedDate, CallTime, LocalNum, CardPresent, SubStatus, Currency, brnnam Branch from ( " +
                     "select a.rid, TicketNo, InOutBound, CommChannel, DateReceived, " +
                     "a.reasoncode + ' - ' + abbreviation abbreviation, CustomerName, " +
                     "CardNo, Destination, TransType, ATMTrans, Location, ATMUsed, PaymentOf, TerminalUsed, BillerName, Merchant, Inter, BancnetUsed, " +
                     "Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, PLStatus, CardBlocked, EndoredTo, Activity = rtrim(Activity), " +
                     "Status, Agent, EndorsedFrom, resolved, Remarks, ReferedBy, ContactPerson, a.reasoncode, " +
                     "a.Tagging, created_by, last_update, substatus, currency, c.brnnam, " +
                     "(select isnull(DateReceived, ' ') from CallReport where ticketno = a.TicketNo and [TYPE] = 'Y' and [Status] in ('Resolved', 'Closed', 'Cancelled')) ResolvedDate, CallTime, LocalNum, CardPresent " +
                     "from CallReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
                     "left join Branch c on a.branch = c.brncd " +
                     "where type in ('Y', 'N') ";


            if (_sch.DateSearch == false)
            {
                strSQL += "and convert(date, datereceived) between convert(date, '" + dt1 + "') and convert(date, '" + dt2 + "') ";
            }
            else
            {
                strSQL += "and year(DateReceived) between '2019' and  '" + yr + "' ";
            }

            if (_sch.CustomerName != null)
            {
                strSQL += "and CustomerName like '%" + _sch.CustomerName.Trim() + "%' ";

            }

            if (_sch.CardNo != null)
            {
                strSQL += "and CardNo like '%" + _sch.CardNo.Trim() + "%' ";
            }

            if (_sch.CommType != "0")
            {
                if (_sch.CommType != null)
                {
                    strSQL += "and InOutBound = '" + _sch.CommType + "' ";
                }
            }

            if (_sch.CommChannel != "0")
            {
                if (_sch.CommChannel != null)
                {
                    strSQL += "and CommChannel = '" + _sch.CommChannel + "' ";
                }
            }

            if (_sch.ReasonCode != null)
            {
                strSQL += "and ReasonCode = '" + _sch.ReasonCode + "' ";
            }

            if (_sch.Status != "0")
            {
                if (_sch.Status != null)
                {
                    strSQL += "and Status = '" + _sch.Status + "' ";
                }
            }

            if (_sch.Agent != null)
            {
                strSQL += "and Created_By = '" + _sch.Agent + "' ";
            }

            if (_sch.TicketNo != null)
            {
                strSQL += "and TicketNo like '%" + _sch.TicketNo.Trim() + "%' ";
            }

            if (_sch.ReferredBy != null)
            {
                strSQL += "and ReferedBy like '%" + _sch.ReferredBy.Trim() + "%' ";
            }

            if (_sch.ContactPerson != null)
            {
                strSQL += "and ContactPerson like '%" + _sch.ContactPerson.Trim() + "%' ";
            }

            strSQL += ") x order by x.DateReceived, TicketNo";

            return this.db.Query<CMActivity>(strSQL, new { }, commandTimeout: 0);

            //return this.db.Query<CMActivity>("sp_CMActivity", new { Date1 = _dt1, Date2 = _dt2, Mod = _mod }, commandType: CommandType.StoredProcedure);

        }

        public IEnumerable<ListActivity> ListActivity(string _ticketno, string _dep)
        {

            string sql;

            if (_dep == "M")
            {
                sql = "select a.RID,a.TicketNo,a.InOutBound,a.CommChannel,a.DateReceived,a.ReasonCode + ' - ' + b.Abbreviation Abbreviation, " +
                    "a.CustomerName,a.CardNo,a.Destination,a.TransType,a.ATMTrans,a.Location,a.ATMUsed,a.PaymentOf,a.TerminalUsed,a.BillerName, " +
                    "a.Merchant,a.Inter,a.BancnetUsed,a.Online,a.Website,a.RemitFrom,a.RemitConcern,a.Amount,a.DateTrans,a.PLStatus,a.CardBlocked, " +
                    "a.EndoredTo,a.Activity,a.Status,a.Agent,a.created_by,a.last_update,a.EndorsedFrom,a.Resolved,a.Remarks,a.ReferedBy,a.ContactPerson,a.Tagging,a.CallTime,a.LocalNum, a.CardPresent " +
                    "from CallReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
                    "where a.TicketNo = @ticketno order by DateReceived";
            }
            else
            {
                sql = "select a.RID,a.TicketNo,a.InOutBound,a.CommChannel,a.DateReceived,Abbreviation = a.ReasonCode + ' - ' + b.Abbreviation, " +
                    "a.CustomerName,a.CardNo,a.Destination,a.TransType,a.ATMTrans,a.Location,a.ATMUsed,a.PaymentOf,a.TerminalUsed,a.BillerName, " +
                    "a.Merchant,a.Inter,a.BancnetUsed,a.Online,a.Website,a.RemitFrom,a.RemitConcern,a.Amount,a.DateTrans,a.PLStatus, " +
                    "a.CardBlocked,a.EndoredTo, TRIM(a.Activity) Activity,a.Status,a.Agent,a.Resolved,a.Remarks,a.ReferedBy, " +
                    "a.ContactPerson,a.Tagging,a.CardPresent " +
                    "from ATMReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
                    "where a.TicketNo = @ticketno order by DateReceived";
            }

            return this.db.Query<ListActivity>(sql, new { ticketno = _ticketno }, commandTimeout: 0);
        }

        public CallReport GetTicketNo(int _rid)
        {
            string sql = "select RID, CommChannel, DateReceived, ReasonCode, NameVerified, CustomerName,CardNo,Activity,Status, " +
                "Agent, TicketNo, Type, LastName, FirstName, MiddleName, InOutBound, TransType, Location, ATMUsed, PaymentOf, " +
                "TerminalUsed, Merchant, Inter, BancnetUsed, Online, Website, Amount, DateTrans, PLStatus, CardBlocked, EndoredTo, " +
                "Destination, ATMTrans, RemitFrom, RemitConcern, EndorsedFrom, Aging, Resolved, BillerName, Remarks, ReferedBy, " +
                "ContactPerson, Tagging, CallTime, LocalNum, Created_By, Last_Update, client_email, CardPresent " +
                "from CallReport where rid = @rid";

#pragma warning disable CS8603 // Possible null reference return.
            return db.Query<CallReport>(sql, new { rid = _rid }).FirstOrDefault();
#pragma warning restore CS8603 // Possible null reference return.
        }
#nullable disable
        public List<Combo1> ListReasonCode()
        {
            return this.db.Query<Combo1>("select rcode ID, rcode + ' - ' + Abbreviation Description from ReasonNew order by rcode", new { }).ToList();
        }

        public List<Combo1> Get_Dropdown(string typ)
        {
            return this.db.Query<Combo1>("select Description ID, Description from ATM where DesType = '" + typ + "' order by Description", new { }).ToList();
        }

        public List<Combo2> Get_Dropdown2(string typ)
        {
            return this.db.Query<Combo2>("select ID, Description from ATM where DesType = '" + typ + "' order by Description", new { }).ToList();
        }

        public List<Combo1> Get_Dropdown3()
        {
            return this.db.Query<Combo1>("select brncd ID, brncd + ' - ' + BRNNAM Description from Branch order by Description", new { }).ToList();
        }

        public List<Combo2> Get_Dropdown4()
        {
            return this.db.Query<Combo2>("SELECT ATM_LOC_ID ID, ATM_LOCATION Description FROM ATM_LOCATION ORDER BY ATM_LOCATION", new { }).ToList();
        }

        public IEnumerable<CMActivity> View_Ticket(string _ticketno)
        {

            string strSQL = "select a.RID,a.TicketNo,a.InOutBound,a.CommChannel,a.DateReceived,a.ReasonCode + ' - ' + b.Abbreviation Abbreviation, " +
                                    "a.CustomerName,a.CardNo,a.Destination,a.TransType,a.ATMTrans,a.Location,a.ATMUsed,a.PaymentOf,a.TerminalUsed,a.BillerName, " +
                                    "a.Merchant,a.Inter,a.BancnetUsed,a.Online,a.Website,a.RemitFrom,a.RemitConcern,a.Amount,a.DateTrans,a.PLStatus,a.CardBlocked, " +
                                    "a.EndoredTo,a.Activity,a.Status,a.Agent,a.created_by,a.last_update,a.EndorsedFrom,a.Resolved,a.Remarks,a.ReferedBy,a.ContactPerson,a.Tagging,a.CallTime,a.LocalNum, a.CardPresent " +
                                    "from CallReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
                                    "where a.TicketNo = '" + _ticketno + "' order by DateReceived";

            return this.db.Query<CMActivity>(strSQL, new {  }, commandTimeout: 0);
        }

        public IEnumerable<SearchICBA> Search_ICBA(string _clientname)
        {
            string client = "%" + _clientname + "%";

            string strSQL = "SELECT distinct c.descr,a.crline,b.lname,b.fsname,b.mdname,b.lsname,a.cifkey,a.brncd|| a.modcd|| LPAD (a.acno, 6, '0')|| LPAD (a.chkdgt, 2, '0') AS acctno, a.acsts, " +
                "a.brncd|| a.modcd|| LPAD (a.acno, 6, '0')|| LPAD (a.chkdgt, 2, '0') ||','|| b.lsname||','||b.fsname||','||b.mdname client_name " +
                "FROM sbafinal01.sa01mast a INNER JOIN sbafinal01.cf01cif b ON a.cifkey = b.cifkey " +
                "INNER JOIN sbafinal01.cf99prtyp c ON a.crline = c.crline AND a.prline = c.prline AND a.prtyp = c.prtyp and a.fccd = c.fccd " +
                "where upper(b.lname) like upper('" + client + "') " +
                "union all " +
                "SELECT distinct descr,a.crline,b.lname,b.fsname,b.mdname,b.lsname,a.cifkey,a.brncd || a.modcd || LPAD(a.acno, 6, '0') || LPAD(a.chkdgt, 2, '0') AS acctno, a.fdsts, " +
                "a.brncd|| a.modcd|| LPAD (a.acno, 6, '0')|| LPAD (a.chkdgt, 2, '0') ||','|| b.lsname||','||b.fsname||','||b.mdname client_name " +
                "FROM sbafinal01.fd01mast a INNER JOIN sbafinal01.cf01cif b ON a.cifkey = b.cifkey INNER JOIN sbafinal01.fd06dtrn c ON a.brncd = c.brncd AND a.modcd = c.modcd AND a.acno = c.acno AND a.chkdgt = c.chkdgt " +
                "INNER JOIN sbafinal01.cf99prtyp d ON c.crline = d.crline AND c.prline = d.prline " +
                "where upper(b.lname) like upper('" + client + "') " +
                "union all " +
                "SELECT distinct catdesc,a.crline,b.lname,b.fsname,b.mdname,b.lsname,a.cifkey,a.brncd || a.modcd || LPAD(a.acno, 6, '0') || LPAD(a.chkdgt, 2, '0') AS acctno, a.acsts, " +
                "a.brncd|| a.modcd|| LPAD (a.acno, 6, '0')|| LPAD (a.chkdgt, 2, '0') ||','|| b.lsname||','||b.fsname||','||b.mdname client_name " +
                "FROM sbafinal01.ln01mast a INNER JOIN sbafinal01.cf01cif b ON a.cifkey = b.cifkey " +
                "INNER JOIN sbafinal01.cf99prtyp c ON a.crline = c.crline AND a.prline = c.prline AND a.prtyp = c.prtyp and a.fccd = c.fccd " +
                "where upper(b.lname) like upper('" + client + "')";


            IDbConnection conn = ConnectionManager.GetConnection("I");

            var result = conn.Query<SearchICBA>(strSQL, new { }, commandTimeout: 0);

            ConnectionManager.CloseConnection(conn);

            return result;

        }

        public IEnumerable<SearchBW> Search_BW(string _clientname)
        {
            string client = "%" + _clientname + "%";

            string strSQL = "select * from ( " +
                "select a.card_number, decode(a.emboss_line_1, null, e.last_name || ',' || e.first_name || ',' || e.birth_name, a.emboss_line_1) full_name, " +
                "e.last_name,e.first_name,case when e.birth_name = 'N/A' THEN '' when e.birth_name = 'NA' THEN '' when e.birth_name = 'NONE' THEN '' ELSE e.birth_name end mid_name, " +
                "c.card_status, b.note_text, b.card_expiry_date, d.acct_number, " +
                "(select cc.current_balance - bb.pending_auths " +
                "from bw3.cas_client_account aa inner join bw3.cas_online_account bb on aa.limit_number = bb.limit_number and aa.institution_number = bb.institution_number " +
                "inner join bw3.cas_cycle_book_balance cc on aa.acct_number = cc.acct_number and aa.institution_number = cc.institution_number and cc.processing_status = '004' " +
                "where aa.acct_number = d.acct_number) balance, " +
                "a.card_number||','|| e.last_name || ',' || e.first_name ||','|| case when e.birth_name = 'N/A' THEN '' when e.birth_name = 'NA' THEN '' when e.birth_name = 'NONE' THEN '' ELSE e.birth_name end CLIENT_NAME " +
                "FROM bw3.svc_client_cards a inner join bw3.svc_card_information b on a.card_number = b.card_number " +
                "LEFT JOIN bw3.bwt_card_status c on b.CARD_STATUS = c.index_field and b.institution_number = c.institution_number " +
                "INNER JOIN bw3.CAS_CLIENT_ACCOUNT d ON d.client_number = a.CLIENT_NUMBER AND d.group_number = a.GROUP_NUMBER and d.institution_number = a.institution_number " +
                "INNER JOIN bw3.CIS_CLIENT_DETAILS e ON e.CLIENT_NUMBER = d.CLIENT_NUMBER and e.institution_number = d.institution_number " +
                "WHERE SUBSTR(a.CARD_NUMBER, 1, 6) NOT IN('461753', '604885', '486392') " +
                ") where upper(FULL_NAME) LIKE upper('" + client + "')";


            IDbConnection conn = ConnectionManager.GetConnection("B");

            var result = conn.Query<SearchBW>(strSQL, new { }, commandTimeout: 0);

            ConnectionManager.CloseConnection(conn);

            return result;

        }

        public List<Combo3> Get_Dropdown_Location(int _id)
        {
            return this.db.Query<Combo3>("select convert(nvarchar, b.ATM_LOC_ID) Value, b.ATM_LOCATION Text from ATM_GROUP_LOC a inner join ATM_LOCATION b on a.ATM_LOC_ID = b.ATM_LOC_ID " +
                "where ATM_ID = " + _id + "", new { }).ToList();
        }

        public IEnumerable<ATMLogs> ListATMLogs()
        {
            return this.db.Query<ATMLogs>("sp_ATMLogsNew", new { }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<ATMOps> ListATMOps(string _dt1, string _dt2)
        {
            return this.db.Query<ATMOps>("sp_ATMActivity", new { Date1 = _dt1, Date2 = _dt2 }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<SearchCardNo> Search_Card(string _cardno)
        {
            string cardno = "%" + _cardno + "%";

            string strSQL = "select TicketNo, DateReceived, isnull(CustomerName,'') CustomerName, isnull(LastName,'') LastName, isnull(FirstName, '') FirstName, isnull(MiddleName, '') MiddleName, CardNo, Status, " +
                "CardNo+','+isnull(LastName,'')+','+isnull(FirstName,'')+','+isnull(MiddleName,'') CLIENT_NAME " +
                "from CallReport where CardNo like '" + cardno + "' and type = 'Y'";


            var result = db.Query<SearchCardNo>(strSQL, new { });

            return result;

        }

        public bool CreateNewActivity(string CommChannel,string ReasonCode,string CustomerName,string CardNo,
            string Activity,string Status,string Agent,string TicketNo,string LastName,string FirstName,string MiddleName,string InOutBound,
            string TransType,string Location,string ATMUsed,string PaymentOf,string TerminalUsed,string Merchant,string Inter,string BancnetUsed,string Online,
            string Website,double Amount,string DateTrans,string PLStatus,string CardBlocked,string EndoredTo,string Destination,string ATMTrans,
            string RemitFrom,string RemitConcern,string EndorsedFrom,string Resolved,string BillerName,string Remarks,string ReferedBy,
            string ContactPerson,string Tagging,string CallTime,string LocalNum,string Created_By,string Last_Update,string CardPresent,string SubStatus,string Currency, string Branch
         )
        {
            try
            {

                string ol = (Online == "true") ? "Y" : "N";
                string cb = (CardBlocked == "true") ? "Yes" : "No";

                string sqlQuery = "";

                if (DateTrans == null)
                {
                    if (Resolved == null)
                    {
                        sqlQuery = "INSERT INTO CallReport (CommChannel,DateReceived,ReasonCode,CustomerName,CardNo,Activity,Status,Agent, " +
                            "TicketNo,Type,LastName,FirstName,MiddleName,InOutBound,TransType,Location,ATMUsed,PaymentOf,TerminalUsed,Merchant,Inter, " +
                            "BancnetUsed,Online,Website,Amount,DateTrans,PLStatus,CardBlocked,EndoredTo,Destination,ATMTrans,RemitFrom,RemitConcern, " +
                            "EndorsedFrom,Resolved,BillerName,Remarks,ReferedBy,ContactPerson,Tagging,CallTime,LocalNum,Created_By,Last_Update, " +
                            "CardPresent,SubStatus,Currency, Branch) " +
                            "VALUES('" + CommChannel + "', GETDATE(), '" + ReasonCode + "', '" + CustomerName + "', '" + CardNo + "', '" + Activity.Replace("'", "''") + "', '" + Status + "', '" + Agent + "', '" + TicketNo + "' " +
                            ", 'Y', '" + LastName + "', '" + FirstName + "', '" + MiddleName + "', '" + InOutBound + "', '" + TransType + "', '" + Location + "', '" + ATMUsed + "', '" + PaymentOf + "', '" + TerminalUsed + "', '" + Merchant + "', '" + Inter + "' " +
                            ", '" + BancnetUsed + "', '" + ol + "', '" + Website + "', " + Amount + ", null, '" + PLStatus + "', '" + cb + "', '" + EndoredTo + "', '" + Destination + "', '" + ATMTrans + "', '" + RemitFrom + "', '" + RemitConcern + "' " +
                            ", '" + EndorsedFrom + "', null, '" + BillerName + "', '" + Remarks + "', '" + ReferedBy + "', '" + ContactPerson + "', '" + Tagging + "', '" + CallTime + "', '" + LocalNum + "', '" + Created_By + "', '" + Last_Update + "' " +
                            ", '" + CardPresent + "', '" + SubStatus + "', '" + Currency + "', '" + Branch + "')";

                    }
                    else
					{
                        sqlQuery = "INSERT INTO CallReport (CommChannel,DateReceived,ReasonCode,CustomerName,CardNo,Activity,Status,Agent, " +
                            "TicketNo,Type,LastName,FirstName,MiddleName,InOutBound,TransType,Location,ATMUsed,PaymentOf,TerminalUsed,Merchant,Inter, " +
                            "BancnetUsed,Online,Website,Amount,DateTrans,PLStatus,CardBlocked,EndoredTo,Destination,ATMTrans,RemitFrom,RemitConcern, " +
                            "EndorsedFrom,Resolved,BillerName,Remarks,ReferedBy,ContactPerson,Tagging,CallTime,LocalNum,Created_By,Last_Update, " +
                            "CardPresent,SubStatus,Currency,Branch) " +
                            "VALUES('" + CommChannel + "', GETDATE(), '" + ReasonCode + "', '" + CustomerName + "', '" + CardNo + "', '" + Activity.Replace("'", "''") + "', '" + Status + "', '" + Agent + "', '" + TicketNo + "' " +
                            ", 'Y', '" + LastName + "', '" + FirstName + "', '" + MiddleName + "', '" + InOutBound + "', '" + TransType + "', '" + Location + "', '" + ATMUsed + "', '" + PaymentOf + "', '" + TerminalUsed + "', '" + Merchant + "', '" + Inter + "' " +
                            ", '" + BancnetUsed + "', '" + ol + "', '" + Website + "', " + Amount + ", null, '" + PLStatus + "', '" + cb + "', '" + EndoredTo + "', '" + Destination + "', '" + ATMTrans + "', '" + RemitFrom + "', '" + RemitConcern + "' " +
                            ", '" + EndorsedFrom + "', '" + Resolved + "', '" + BillerName + "', '" + Remarks + "', '" + ReferedBy + "', '" + ContactPerson + "', '" + Tagging + "', '" + CallTime + "', '" + LocalNum + "', '" + Created_By + "', '" + Last_Update + "' " +
                            ", '" + CardPresent + "', '" + SubStatus + "', '" + Currency + "', '" + Branch + "')";

                    }

                }
                else
				{
                    if (Resolved == null)
					{
                        sqlQuery = "INSERT INTO CallReport (CommChannel,DateReceived,ReasonCode,CustomerName,CardNo,Activity,Status,Agent, " +
                            "TicketNo,Type,LastName,FirstName,MiddleName,InOutBound,TransType,Location,ATMUsed,PaymentOf,TerminalUsed,Merchant,Inter, " +
                            "BancnetUsed,Online,Website,Amount,DateTrans,PLStatus,CardBlocked,EndoredTo,Destination,ATMTrans,RemitFrom,RemitConcern, " +
                            "EndorsedFrom,Resolved,BillerName,Remarks,ReferedBy,ContactPerson,Tagging,CallTime,LocalNum,Created_By,Last_Update, " +
                            "CardPresent,SubStatus,Currency,Branch) " +
                            "VALUES('" + CommChannel + "', GETDATE(), '" + ReasonCode + "', '" + CustomerName + "', '" + CardNo + "', '" + Activity.Replace("'", "''") + "', '" + Status + "', '" + Agent + "', '" + TicketNo + "' " +
                            ", 'Y', '" + LastName + "', '" + FirstName + "', '" + MiddleName + "', '" + InOutBound + "', '" + TransType + "', '" + Location + "', '" + ATMUsed + "', '" + PaymentOf + "', '" + TerminalUsed + "', '" + Merchant + "', '" + Inter + "' " +
                            ", '" + BancnetUsed + "', '" + ol + "', '" + Website + "', " + Amount + ", '" + DateTrans + "', '" + PLStatus + "', '" + cb + "', '" + EndoredTo + "', '" + Destination + "', '" + ATMTrans + "', '" + RemitFrom + "', '" + RemitConcern + "' " +
                            ", '" + EndorsedFrom + "', null, '" + BillerName + "', '" + Remarks + "', '" + ReferedBy + "', '" + ContactPerson + "', '" + Tagging + "', '" + CallTime + "', '" + LocalNum + "', '" + Created_By + "', '" + Last_Update + "' " +
                            ", '" + CardPresent + "', '" + SubStatus + "','" + Currency + "', '" + Branch + "')";
                    }
                    else
					{
                        sqlQuery = "INSERT INTO CallReport (CommChannel,DateReceived,ReasonCode,CustomerName,CardNo,Activity,Status,Agent, " +
                            "TicketNo,Type,LastName,FirstName,MiddleName,InOutBound,TransType,Location,ATMUsed,PaymentOf,TerminalUsed,Merchant,Inter, " +
                            "BancnetUsed,Online,Website,Amount,DateTrans,PLStatus,CardBlocked,EndoredTo,Destination,ATMTrans,RemitFrom,RemitConcern, " +
                            "EndorsedFrom,Resolved,BillerName,Remarks,ReferedBy,ContactPerson,Tagging,CallTime,LocalNum,Created_By,Last_Update, " +
                            "CardPresent,SubStatus,Currency,Branch) " +
                            "VALUES('" + CommChannel + "', GETDATE(), '" + ReasonCode + "', '" + CustomerName + "', '" + CardNo + "', '" + Activity.Replace("'", "''") + "', '" + Status + "', '" + Agent + "', '" + TicketNo + "' " +
                            ", 'Y', '" + LastName + "', '" + FirstName + "', '" + MiddleName + "', '" + InOutBound + "', '" + TransType + "', '" + Location + "', '" + ATMUsed + "', '" + PaymentOf + "', '" + TerminalUsed + "', '" + Merchant + "', '" + Inter + "' " +
                            ", '" + BancnetUsed + "', '" + ol + "', '" + Website + "', " + Amount + ", '" + DateTrans + "', '" + PLStatus + "', '" + cb + "', '" + EndoredTo + "', '" + Destination + "', '" + ATMTrans + "', '" + RemitFrom + "', '" + RemitConcern + "' " +
                            ", '" + EndorsedFrom + "', '" + Resolved + "', '" + BillerName + "', '" + Remarks + "', '" + ReferedBy + "', '" + ContactPerson + "', '" + Tagging + "', '" + CallTime + "', '" + LocalNum + "', '" + Created_By + "', '" + Last_Update + "' " +
                            ", '" + CardPresent + "', '" + SubStatus + "','" + Currency + "', '" + Branch + "')";
                    }

                }

                db.Execute(sqlQuery, new { }, commandTimeout: 0);

                return true;
            }
            catch (Exception ex)
            {
                string x = ex.Message;
                return false;
            }
        }

        public Show_Ticket DisplayTickets(string _rcode)
        {
            return this.db.Query<Show_Ticket>("sp_TicketNo", new { RCode = _rcode }, commandTimeout: 0, commandType: CommandType.StoredProcedure).FirstOrDefault();
        }

        public bool UploadInputFile(string file, string ticketno)
        {
            try
            {
                string sqlQuery = "Insert into file_location (ticketno, downloadfile) " +
                    "Values ('" + ticketno + "', '" + file + "')";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public EditCallReport ViewCMEntry(int _id)
        {
            string sql = "select RID,TicketNo,CommChannel,ReasonCode,NameVerified,CustomerName,CardNo,Activity,Status,Agent,LastName,FirstName,MiddleName,InOutBound,TransType, " +
                "Location,ATMUsed,PaymentOf,TerminalUsed,Merchant,Inter,BancnetUsed,case when Online = 'N' then 0 else 1 end Online,Website,Amount,case when convert(date, DateTrans) = '1900-01-01' then null else FORMAT(DateTrans,'dd-MMM-yyyy') end DateTrans, " +
                "PLStatus,case when CardBlocked = 'No' then 0 else 1 end CardBlocked,EndoredTo,Destination, " +
                "ATMTrans,RemitFrom,RemitConcern,EndorsedFrom,case when convert(date, Resolved) = '1900-01-01' then null else FORMAT(Resolved,'dd-MMM-yyyy') end Resolved, " +
                "BillerName,Remarks,ReferedBy,ContactPerson,Tagging,CallTime,LocalNum,Created_By,Last_Update, " +
                "CardPresent,SubStatus, Currency, Branch from CallReport where rid = " + _id + "";

            return this.db.Query<EditCallReport>(sql, new { }).FirstOrDefault();
        }

        public IEnumerable<CMActivity> View_TicketNo(string _ticketno)
        {

            string strSQL;

            //strSQL = "select RID, TicketNo, InOutBound, CommChannel, DateReceived, abbreviation, CustomerName, " +
            //         "CardNo, Destination, TransType, ATMTrans, case when isnumeric(Location) = 0 then Location else dbo.Get_ATMLocation(ATMUsed, Location) end Location,  " +
            //         "case when isnumeric(ATMUsed) = 0 then ATMUsed else dbo.Get_ATMUsed(ATMUsed) end ATMUsed,  PaymentOf, TerminalUsed, BillerName, " +
            //         "Merchant, Inter, BancnetUsed, Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, " +
            //         "PLStatus, CardBlocked, EndoredTo, Activity, Status, Agent, created_by, Last_Update, EndorsedFrom, Resolved, Remarks, " +
            //         "ReferedBy, ContactPerson, ReasonCode, Tagging, ResolvedDate, CallTime, LocalNum, CardPresent, SubStatus, Currency, Branch from ( " +
            //         "select a.rid, TicketNo, InOutBound, CommChannel, " +
            //         "DateReceived, " +
            //         "a.reasoncode + ' - ' + abbreviation abbreviation, CustomerName, " +
            //         "CardNo, Destination, TransType, ATMTrans, Location, ATMUsed, PaymentOf, TerminalUsed, BillerName, Merchant, Inter, BancnetUsed, " +
            //         "Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, PLStatus, CardBlocked, EndoredTo, Activity = rtrim(Activity), " +
            //         "Status, Agent, EndorsedFrom, resolved, Remarks, ReferedBy, ContactPerson, a.reasoncode, " +
            //         "a.Tagging, created_by, last_update, " +
            //         "(select isnull(DateReceived, '') from CallReport where ticketno = a.ticketno and [TYPE] = 'Y' and [Status] in ('Resolved', 'Closed', 'Cancelled')) ResolvedDate, " +
            //         "CallTime, LocalNum, CardPresent, SubStatus, Currency, brnnam Branch " +
            //         "from CallReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
            //         "left join Branch c on a.branch = c.brncd " +
            //         "where ticketno = '" + _ticketno + "') x order by x.DateReceived, TicketNo ";

            strSQL = "select RID, TicketNo, InOutBound, CommChannel, DateReceived, abbreviation, CustomerName, CardNo, Destination, TransType, " +
                "ATMTrans, case when isnumeric(Location) = 0 then Location else dbo.Get_ATMLocation(ATMUsed, Location) end Location, " +
                "case when isnumeric(ATMUsed) = 0 then ATMUsed else dbo.Get_ATMUsed(ATMUsed) end ATMUsed, PaymentOf, TerminalUsed, BillerName, " +
                "Merchant, Inter, BancnetUsed, Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, " +
                "PLStatus, CardBlocked, EndoredTo, Activity, Status, Agent, created_by, Last_Update, EndorsedFrom, Resolved, Remarks, " +
                "ReferedBy, ContactPerson, ReasonCode, Tagging, ResolvedDate, CallTime, LocalNum, CardPresent, SubStatus, Currency, Branch from ( " +
                "select a.rid, TicketNo, InOutBound, CommChannel, DateReceived, a.reasoncode + ' - ' + abbreviation abbreviation, CustomerName, " +
                "CardNo, Destination, TransType, ATMTrans, Location, ATMUsed, PaymentOf, TerminalUsed, BillerName, Merchant, Inter, BancnetUsed, " +
                "Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, PLStatus, CardBlocked, EndoredTo, Activity = rtrim(Activity), " +
                "Status, Agent, EndorsedFrom, resolved, Remarks, ReferedBy, ContactPerson, a.reasoncode, a.Tagging, created_by, last_update, " +
                "(select isnull(DateReceived, '') from CallReport where ticketno = a.ticketno and[TYPE] = 'Y' and[Status] in ('Resolved', 'Closed', 'Cancelled')) ResolvedDate, " +
                "CallTime, LocalNum, CardPresent, SubStatus, Currency, brnnam Branch " +
                "from CallReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
                "left join Branch c on a.branch = c.brncd where ticketno = '" + _ticketno + "') x  order by x.DateReceived, TicketNo ";
                //"union all " +
                //"select RID, TicketNo, InOutBound, CommChannel, DateReceived, abbreviation, CustomerName, " +
                //"CardNo, Destination, TransType, ATMTrans, case when isnumeric(Location) = 0 then Location else dbo.Get_ATMLocation(ATMUsed, Location) end Location, " +
                //"case when isnumeric(ATMUsed) = 0 then ATMUsed else dbo.Get_ATMUsed(ATMUsed) end ATMUsed, PaymentOf, TerminalUsed, BillerName, " +
                //"Merchant, Inter, BancnetUsed, Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, " +
                //"PLStatus, CardBlocked, EndoredTo, Activity, Status, Agent, created_by, Last_Update, EndorsedFrom, Resolved, Remarks, " +
                //"ReferedBy, ContactPerson, ReasonCode, Tagging, ResolvedDate, CallTime, LocalNum, CardPresent, SubStatus, Currency, Branch from ( " +
                //"select a.rid, TicketNo, InOutBound, CommChannel, DateReceived, a.reasoncode + ' - ' + abbreviation abbreviation, CustomerName, " +
                //"CardNo, Destination, TransType, ATMTrans, Location, ATMUsed, PaymentOf, TerminalUsed, BillerName, Merchant, Inter, BancnetUsed, " +
                //"Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, PLStatus, CardBlocked, EndoredTo, Activity = rtrim(Activity), " +
                //"Status, Agent, '' EndorsedFrom, resolved, Remarks, ReferedBy, ContactPerson, a.reasoncode, " +
                //"a.Tagging, created_by, last_update, " +
                //"(select isnull(DateReceived, '') from CallReport where ticketno = a.ticketno and[TYPE] = 'Y' and[Status] in ('Resolved', 'Closed', 'Cancelled')) ResolvedDate, " +
                //"'' CallTime, '' LocalNum, CardPresent, SubStatus, Currency, brnnam Branch " +
                //"from ATMReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
                //"left join Branch c on a.branch = c.brncd " +
                //"where ticketno = '" + _ticketno + "') x order by x.DateReceived, TicketNo ";

            return this.db.Query<CMActivity>(strSQL, new { }, commandTimeout: 0);

        }

        public bool SendEndorsedMail(string _tgi, string _irmt, string _emailto, string _ticketno, string _remarks, string _lastname, string _firstname, string _midname, string _cardno, string _activity, string _fullname, string _referredby, string _contactperson, List<string> attachments)
        {
            try
            {

                string mailserver = Connect.ConfigurationManager.AppSetting["AppEmail:MailServer"];
                string mailserverport = Connect.ConfigurationManager.AppSetting["AppEmail:Port"];
                string mailuser = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppEmail:EUser"]);
                string mailpass = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppEmail:EPass"]);

                System.Net.NetworkCredential credentials = new(mailuser, mailpass);
                SmtpClient client = new SmtpClient(mailserver, Convert.ToInt16(mailserverport));

                client.UseDefaultCredentials = false;

                client.Credentials = credentials;

                client.EnableSsl = false;

                System.Collections.ArrayList emailArray = new System.Collections.ArrayList();

                MailMessage message = new MailMessage();
                MailAddress toAddress = new MailAddress(_emailto);

                string mailto = toAddress.ToString();

                string cardno = _cardno.Substring(0, 6) + "******" + _cardno.Substring(_cardno.Length - 4, 4);

                if (_irmt == "Y")
                {
					//message.From = new MailAddress("aalivara@sterlingbankasia.com", "IRemit Services");
					//message.CC.Add(new MailAddress("aalivara@sterlingbankasia.com", "IRemit Services"));

					message.From = new MailAddress("iremit.services@sterlingbankasia.com", "IRemit Services");
					message.CC.Add(new MailAddress("iremit.services@sterlingbankasia.com", "IRemit Services"));
				}
                else
				{
					//message.From = new MailAddress("aalivara@sterlingbankasia.com", "CUSTOMER MANAGEMENT");
					//message.CC.Add(new MailAddress("aalivara@sterlingbankasia.com", "CUSTOMER MANAGEMENT"));

					message.From = new MailAddress("customer.management@sterlingbankasia.com", "CUSTOMER MANAGEMENT");
					message.CC.Add(new MailAddress("customer.management@sterlingbankasia.com", "CUSTOMER MANAGEMENT"));

				}


                message.To.Add(toAddress);

                string flw = ((_tgi == "1") ? " - (FOLLOW UP)" : string.Empty);

                message.Subject = "Service Ticket No. " +
                                   _ticketno + " " +
                                   _lastname + ", " +
                                   _firstname + ", " +
                                   _midname + flw;

                message.IsBodyHtml = true;

				string mbody = "<p>This is an auto generated email message from Customer Service with a corresponding Service Ticket Number " + _ticketno + ".</p> " +
                    "<p> You are receiving this because the subject concerns your area of expertise. Details are below.</p> " +
                    "<hr/> " +
                    "<p> " + _remarks + " </p> " +
                    "<hr/> " +
                    "<p> Message:</p> " +
                    "<p> " + _lastname + ", " + _firstname + ", " + _midname + "</p> " +
                    "<p> " + cardno + " </p> " +
                    "<p> " + _activity + " </p> " +
                    "<hr/> " +
                    "<p> We would appreciate your prompt feedback. Please reply to all with the same subject heading.</p> " +
                    "<p> &nbsp;</p> " +
                    "<p> Thank you.</p> " +
                    "<p> &nbsp;</p> " +
                    "<p> CUSTOMER MANAGEMENT </p> " +
                    "<p> Endorser: " + _fullname + "</p> " +
                    "<p> Referred By: " + _referredby + " </p> " +
                    "<p> Contact Person: " + _contactperson + " </p>";

                message.Body = mbody;

                if (attachments != null)
                {
                    foreach (var attachment in attachments)
                    {
                        message.Attachments.Add(new Attachment(attachment));
                    }
                }

                client.Send(message);

                return true;
            }
            catch (Exception ex)
            {
                string x = ex.Message;
                return false;
            }

        }

        public bool SendEndorsedMailATM(string _irmt, string _ticketno, string _remarks, string _lastname, string _firstname, string _midname, string _cardno, string _activity, string _fullname, string _referredby, string _contactperson, List<string> attachments)
        {
            try
            {

                string mailserver = Connect.ConfigurationManager.AppSetting["AppEmail:MailServer"];
                string mailserverport = Connect.ConfigurationManager.AppSetting["AppEmail:Port"];
                string mailuser = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppEmail:EUser"]);
                string mailpass = Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppEmail:EPass"]);

                System.Net.NetworkCredential credentials = new(mailuser, mailpass);
                SmtpClient client = new SmtpClient(mailserver, Convert.ToInt16(mailserverport));

                client.UseDefaultCredentials = false;

                client.Credentials = credentials;

                client.EnableSsl = false;

                System.Collections.ArrayList emailArray = new System.Collections.ArrayList();

                string cardno = _cardno.Substring(0, 6) + "******" + _cardno.Substring(_cardno.Length - 4, 4);

                MailMessage message = new MailMessage();
                //MailAddress toAddress = new MailAddress(_emailto);

                if (_irmt == "Y")
                {
					//message.From = new MailAddress("aalivara@sterlingbankasia.com", "ATM Operations");
					//message.To.Add(new MailAddress("aalivara@sterlingbankasia.com", "IRemit Services"));
					//message.CC.Add(new MailAddress("aalivara@sterlingbankasia.com", "ATM Operations"));

					message.From = new MailAddress("atm.ops@sterlingbankasia.com", "ATM Operations");
					message.To.Add(new MailAddress("iremit.services@sterlingbankasia.com", "IRemit Services"));
					message.CC.Add(new MailAddress("atm.ops@sterlingbankasia.com", "ATM Operations"));



				}
                else
                {
					message.From = new MailAddress("atm.ops@sterlingbankasia.com", "ATM Operations");
					message.To.Add(new MailAddress("customer.management@sterlingbankasia.com", "CUSTOMER MANAGEMENT"));
					message.CC.Add(new MailAddress("atm.ops@sterlingbankasia.com", "ATM Operations"));

					//message.From = new MailAddress("aalivara@sterlingbankasia.com", "ATM Operations");
					//message.To.Add(new MailAddress("aalivara@sterlingbankasia.com", "CUSTOMER MANAGEMENT"));
					//message.CC.Add(new MailAddress("aalivara@sterlingbankasia.com", "ATM Operations"));

				}


                message.Subject = "Service Ticket No. " +
                                   _ticketno + " " +
                                   _lastname + ", " +
                                   _firstname + ", " +
                                   _midname;

                message.IsBodyHtml = true;

                string mbody = "<p>This is an auto generated email message from ATM Operations with a corresponding Service Ticket Number " + _ticketno + ".</p> " +
                    "<p> You are receiving this because the subject concerns your area of expertise. Details are below.</p> " +
                    "<hr/> " +
                    "<p> " + _remarks + " </p> " +
                    "<hr/> " +
                    "<p> Message:</p> " +
                    "<p> " + _lastname + ", " + _firstname + ", " + _midname + "</p> " +
                    "<p> " + cardno + " </p> " +
                    "<p> " + _activity + " </p> " +
                    "<hr/> " +
                    "<p> We would appreciate your prompt feedback. Please reply to all with the same subject heading.</p> " +
                    "<p> &nbsp;</p> " +
                    "<p> Thank you.</p> " +
                    "<p> &nbsp;</p> " +
                    "<p> ATM OPERATIONS </p> " +
                    "<p> Endorser: " + _fullname + "</p> " +
                    "<p> Referred By: " + _referredby + " </p> " +
                    "<p> Contact Person: " + _contactperson + " </p>";

                message.Body = mbody;

                if (attachments != null)
                {
                    foreach (var attachment in attachments)
                    {
                        message.Attachments.Add(new Attachment(attachment));
                    }
                }

                client.Send(message);

                return true;
            }
            catch (Exception ex)
            {
                string x = ex.Message;
                return false;
            }

        }


        public bool ChkIremitCode(string _rcd)
        {
            int cnt = this.db.Query<IRemitCode>("Select RCode from iremitcodes where rcode = '" + _rcd + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool CreateNewATMActivity(string CommChannel, string ReasonCode, string CustomerName, string CardNo,
            string Activity, string Status, string Agent, string TicketNo, string LastName, string FirstName, string MiddleName, string InOutBound,
            string TransType, string Location, string ATMUsed, string PaymentOf, string TerminalUsed, string Merchant, string Inter, string BancnetUsed, string Online,
            string Website, double Amount, string DateTrans, string PLStatus, string CardBlocked, string EndoredTo, string Destination, string ATMTrans,
            string RemitFrom, string RemitConcern, string EndorsedFrom, string Resolved, string BillerName, string Remarks, string ReferedBy,
            string ContactPerson, string Tagging, string CallTime, string LocalNum, string Created_By, string Last_Update, string CardPresent, string SubStatus, string Currency, string Branch
         )
        {
            try
            {

                string ol = (Online == "true") ? "Y" : "N";
                string cb = (CardBlocked == "true") ? "Yes" : "No";

                string sqlQuery = "";

                if (DateTrans == null)
                {
                    if (Resolved == null)
                    {
                        sqlQuery = "insert into ATMReport (TicketNo,CommChannel,DateReceived,ReasonCode,CustomerName,CardNo,Activity,Status,Agent,Type,LastName,FirstName,MiddleName,InOutBound, " +
                            "TransType,Location,ATMUsed,PaymentOf,TerminalUsed,BillerName,Merchant,Inter,BancnetUsed,Online,Website,Amount,DateTrans,PLStatus,CardBlocked, " +
                            "EndoredTo,Destination,ATMTrans,RemitFrom,RemitConcern,Resolved,Remarks,ReferedBy,ContactPerson,Tagging,CardPresent,SubStatus,Currency,Branch,Created_By,Last_Update) " +
                            "values('" + TicketNo + "', '" + CommChannel + "', getdate(), '" + ReasonCode + "', '" + CustomerName + "', '" + CardNo + "', '" + Activity.Replace("'", "''") + "', '" + Status + "', '" + Agent + "', 'Y', " +
                            "'" + LastName + "', '" + FirstName + "', '" + MiddleName + "', '" + InOutBound + "', '" + TransType + "', '" + Location + "', '" + ATMUsed + "', '" + PaymentOf + "', " +
                            "'" + TerminalUsed + "', '" + BillerName + "', '" + Merchant + "', '" + Inter + "', '" + BancnetUsed + "', '" + ol + "', '" + Website + "', " + Amount + ", " +
                            "NULL, '" + PLStatus + "', '" + cb + "', '" + EndoredTo + "', '" + Destination + "', '" + ATMTrans + "', '" + RemitFrom + "', '" + RemitConcern + "', " +
                            "NULL, '" + Remarks + "', '" + ReferedBy + "', '" + ContactPerson + "', '" + Tagging + "', '" + CardPresent + "', '" + SubStatus + "', '" + Currency + "', '" + Branch + "', '" + Created_By + "', '" + Last_Update + "')";
                    }
                    else
                    {
                        sqlQuery = "insert into ATMReport (TicketNo,CommChannel,DateReceived,ReasonCode,CustomerName,CardNo,Activity,Status,Agent,Type,LastName,FirstName,MiddleName,InOutBound, " +
                            "TransType,Location,ATMUsed,PaymentOf,TerminalUsed,BillerName,Merchant,Inter,BancnetUsed,Online,Website,Amount,DateTrans,PLStatus,CardBlocked, " +
                            "EndoredTo,Destination,ATMTrans,RemitFrom,RemitConcern,Resolved,Remarks,ReferedBy,ContactPerson,Tagging,CardPresent,SubStatus,Currency,Branch,Created_By,Last_Update) " +
                            "values('" + TicketNo + "', '" + CommChannel + "', getdate(), '" + ReasonCode + "', '" + CustomerName + "', '" + CardNo + "', '" + Activity.Replace("'", "''") + "', '" + Status + "', '" + Agent + "', 'Y', " +
                            "'" + LastName + "', '" + FirstName + "', '" + MiddleName + "', '" + InOutBound + "', '" + TransType + "', '" + Location + "', '" + ATMUsed + "', '" + PaymentOf + "', " +
                            "'" + TerminalUsed + "', '" + BillerName + "', '" + Merchant + "', '" + Inter + "', '" + BancnetUsed + "', '" + ol + "', '" + Website + "', " + Amount + ", " +
                            "NULL, '" + PLStatus + "', '" + cb + "', '" + EndoredTo + "', '" + Destination + "', '" + ATMTrans + "', '" + RemitFrom + "', '" + RemitConcern + "', " +
                            "'" + Resolved + "', '" + Remarks + "', '" + ReferedBy + "', '" + ContactPerson + "', '" + Tagging + "', '" + CardPresent + "', '" + SubStatus + "', '" + Currency + "', '" + Branch + "', '" + Created_By + "', '" + Last_Update + "')";

                    }

                }
                else
                {
                    if (Resolved == null)
                    {
                        sqlQuery = "insert into ATMReport (TicketNo,CommChannel,DateReceived,ReasonCode,CustomerName,CardNo,Activity,Status,Agent,Type,LastName,FirstName,MiddleName,InOutBound, " +
                           "TransType,Location,ATMUsed,PaymentOf,TerminalUsed,BillerName,Merchant,Inter,BancnetUsed,Online,Website,Amount,DateTrans,PLStatus,CardBlocked, " +
                           "EndoredTo,Destination,ATMTrans,RemitFrom,RemitConcern,Resolved,Remarks,ReferedBy,ContactPerson,Tagging,CardPresent,SubStatus,Currency,Branch,Created_By,Last_Update) " +
                           "values('" + TicketNo + "', '" + CommChannel + "', getdate(), '" + ReasonCode + "', '" + CustomerName + "', '" + CardNo + "', '" + Activity.Replace("'", "''") + "', '" + Status + "', '" + Agent + "', 'Y', " +
                           "'" + LastName + "', '" + FirstName + "', '" + MiddleName + "', '" + InOutBound + "', '" + TransType + "', '" + Location + "', '" + ATMUsed + "', '" + PaymentOf + "', " +
                           "'" + TerminalUsed + "', '" + BillerName + "', '" + Merchant + "', '" + Inter + "', '" + BancnetUsed + "', '" + ol + "', '" + Website + "', " + Amount + ", " +
                           "'" + DateTrans + "', '" + PLStatus + "', '" + cb + "', '" + EndoredTo + "', '" + Destination + "', '" + ATMTrans + "', '" + RemitFrom + "', '" + RemitConcern + "', " +
                           "NULL, '" + Remarks + "', '" + ReferedBy + "', '" + ContactPerson + "', '" + Tagging + "', '" + CardPresent + "', '" + SubStatus + "', '" + Currency + "', '" + Branch + "', '" + Created_By + "', '" + Last_Update + "')";
                    }
                    else
                    {
                        sqlQuery = "insert into ATMReport (TicketNo,CommChannel,DateReceived,ReasonCode,CustomerName,CardNo,Activity,Status,Agent,Type,LastName,FirstName,MiddleName,InOutBound, " +
                           "TransType,Location,ATMUsed,PaymentOf,TerminalUsed,BillerName,Merchant,Inter,BancnetUsed,Online,Website,Amount,DateTrans,PLStatus,CardBlocked, " +
                           "EndoredTo,Destination,ATMTrans,RemitFrom,RemitConcern,Resolved,Remarks,ReferedBy,ContactPerson,Tagging,CardPresent,SubStatus,Currency,Branch,Created_By,Last_Update) " +
                           "values('" + TicketNo + "', '" + CommChannel + "', getdate(), '" + ReasonCode + "', '" + CustomerName + "', '" + CardNo + "', '" + Activity.Replace("'", "''") + "', '" + Status + "', '" + Agent + "', 'Y', " +
                           "'" + LastName + "', '" + FirstName + "', '" + MiddleName + "', '" + InOutBound + "', '" + TransType + "', '" + Location + "', '" + ATMUsed + "', '" + PaymentOf + "', " +
                           "'" + TerminalUsed + "', '" + BillerName + "', '" + Merchant + "', '" + Inter + "', '" + BancnetUsed + "', '" + ol + "', '" + Website + "', " + Amount + ", " +
                           "'" + DateTrans + "', '" + PLStatus + "', '" + cb + "', '" + EndoredTo + "', '" + Destination + "', '" + ATMTrans + "', '" + RemitFrom + "', '" + RemitConcern + "', " +
                           "'" + Resolved + "', '" + Remarks + "', '" + ReferedBy + "', '" + ContactPerson + "', '" + Tagging + "', '" + CardPresent + "', '" + SubStatus + "', '" + Currency + "', '" + Branch + "', '" + Created_By + "', '" + Last_Update + "')";
                    }

                }

                db.Execute(sqlQuery, new { }, commandTimeout: 0);

                return true;
            }
            catch (Exception ex)
            {
                string x = ex.Message;
                return false;
            }
        }

        public bool UpdateType(int _id, string _typ)
        {
            try
            {
                if (_typ == "C")
				{
                    string sqlQuery = "UPDATE CallReport SET Type = 'N' WHERE RID = " + _id + "";
                    db.Execute(sqlQuery, new { });
                }
                else
				{
                    string sqlQuery = "UPDATE ATMReport SET Type = 'N' WHERE RID = " + _id + "";
                    db.Execute(sqlQuery, new { });
                }


                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool DeleteCMActivity(string _ticketno)
        {
            try
            {
                string sqlQuery = "DELETE FROM CallReport WHERE TicketNo = '" + _ticketno + "'";
                db.Execute(sqlQuery, new { });

                string sqlQuery2 = "DELETE FROM ATMReport WHERE TicketNo = '" + _ticketno + "'";
                db.Execute(sqlQuery2, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ChkActivityInATM(string _ticketno)
        {
            int cnt = this.db.Query<ATMReport>("select * from CallReport where TicketNo = '" + _ticketno + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool EditNewActivity(EditCallReport edt)
        {
            try
            {

				//string ol = (edt.Online == true) ? "Y" : "N";
				//string cb = (edt.CardBlocked == true) ? "Yes" : "No";
    //            int amt = (edt.Amount == null) ? 0 : Convert.ToInt32(edt.Amount);

				string sqlQuery = "";

                sqlQuery = "update CallReport set CustomerName = '" + edt.CustomerName + "', CardNo = '" + edt.CardNo + "', Activity = '" + edt.Activity.Replace("'", "''") + "', Status = '" + edt.Status + "', " +
                    "LastName = '" + edt.LastName + "', FirstName = '" + edt.FirstName + "', MiddleName = '" + edt.MiddleName + "', InOutBound = '" + edt.InOutBound + "', " +
                    "Tagging = '" + edt.Tagging + "', CallTime = '" + edt.CallTime + "', LocalNum = '" + edt.LocalNum + "' " +
                    "where RID = " + edt.RID + "";

     //           if (edt.DateTrans == null || edt.DateTrans == string.Empty)
     //           {
     //               if (edt.Resolved == null || edt.Resolved == string.Empty)
     //               {

					//	sqlQuery = "update CallReport set CommChannel = '" + edt.CommChannel + "', DateReceived = GETDATE(), ReasonCode = '" + edt.ReasonCode + "', " +
					//		"CustomerName = '" + edt.CustomerName + "', CardNo = '" + edt.CardNo + "', Activity = '" + edt.Activity.Replace("'", "''") + "', Status = '" + edt.Status + "', " +
					//		"Agent = '" + edt.Agent + "', TicketNo = '" + edt.TicketNo + "', Type = 'Y', LastName = '" + edt.LastName + "', FirstName = '" + edt.FirstName + "', MiddleName = '" + edt.MiddleName + "', " +
					//		"InOutBound = '" + edt.InOutBound + "', TransType = '" + edt.TransType + "', Location = '" + edt.Location + "', ATMUsed = '" + edt.ATMUsed + "', PaymentOf = '" + edt.PaymentOf + "', " +
					//		"TerminalUsed = '" + edt.TerminalUsed + "', Merchant = '" + edt.Merchant + "', Inter = '" + edt.Inter + "',  BancnetUsed = '" + edt.BancnetUsed + "', Online = '" + ol + "', " +
					//		"Website = '" + edt.Website + "', Amount = " + amt + ", DateTrans = NULL, PLStatus = '" + edt.PLStatus + "', CardBlocked = '" + cb + "', EndoredTo = '" + edt.EndoredTo + "', " +
					//		"Destination = '" + edt.Destination + "', ATMTrans = '" + edt.ATMTrans + "', RemitFrom = '" + edt.RemitFrom + "', RemitConcern = '" + edt.RemitConcern + "', EndorsedFrom = '" + edt.EndorsedFrom + "', " +
					//		"Resolved = NULL, BillerName = '" + edt.BillerName + "', Remarks = '" + edt.Remarks + "', ReferedBy = '" + edt.ReferedBy + "', ContactPerson = '" + edt.ContactPerson + "', " +
					//		"Tagging = '" + edt.Tagging + "', CallTime = '" + edt.CallTime + "', LocalNum = '" + edt.LocalNum + "', Created_By = '" + edt.Created_By + "', Last_Update = '" + edt.Last_Update + "', " +
					//		"CardPresent = '" + edt.CardPresent + "', SubStatus = '" + edt.SubStatus + "', Currency = '" + edt.Currency + "', Branch = '" + edt.Branch + "' where RID = " + edt.RID + "";
					//}
     //               else
     //               {

     //                   sqlQuery = "update CallReport set CommChannel = '" + edt.CommChannel + "', DateReceived = GETDATE(), ReasonCode = '" + edt.ReasonCode + "', " +
     //                       "CustomerName = '" + edt.CustomerName + "', CardNo = '" + edt.CardNo + "', Activity = '" + edt.Activity.Replace("'", "''") + "', Status = '" + edt.Status + "', " +
     //                       "Agent = '" + edt.Agent + "', TicketNo = '" + edt.TicketNo + "', Type = 'Y', LastName = '" + edt.LastName + "', FirstName = '" + edt.FirstName + "', MiddleName = '" + edt.MiddleName + "', " +
     //                       "InOutBound = '" + edt.InOutBound + "', TransType = '" + edt.TransType + "', Location = '" + edt.Location + "', ATMUsed = '" + edt.ATMUsed + "', PaymentOf = '" + edt.PaymentOf + "', " +
     //                       "TerminalUsed = '" + edt.TerminalUsed + "', Merchant = '" + edt.Merchant + "', Inter = '" + edt.Inter + "',  BancnetUsed = '" + edt.BancnetUsed + "', Online = '" + ol + "', " +
     //                       "Website = '" + edt.Website + "', Amount = " + amt + ", DateTrans = NULL, PLStatus = '" + edt.PLStatus + "', CardBlocked = '" + cb + "', EndoredTo = '" + edt.EndoredTo + "', " +
     //                       "Destination = '" + edt.Destination + "', ATMTrans = '" + edt.ATMTrans + "', RemitFrom = '" + edt.RemitFrom + "', RemitConcern = '" + edt.RemitConcern + "', EndorsedFrom = '" + edt.EndorsedFrom + "', " +
     //                       "Resolved = '" + edt.Resolved + "', BillerName = '" + edt.BillerName + "', Remarks = '" + edt.Remarks + "', ReferedBy = '" + edt.ReferedBy + "', ContactPerson = '" + edt.ContactPerson + "', " +
     //                       "Tagging = '" + edt.Tagging + "', CallTime = '" + edt.CallTime + "', LocalNum = '" + edt.LocalNum + "', Created_By = '" + edt.Created_By + "', Last_Update = '" + edt.Last_Update + "', " +
     //                       "CardPresent = '" + edt.CardPresent + "', SubStatus = '" + edt.SubStatus + "', Currency = '" + edt.Currency + "', Branch = '" + edt.Branch + "' where RID = " + edt.RID + "";

     //               }

     //           }
                //else
                //{
                //    if (edt.Resolved == null || edt.Resolved == string.Empty)
                //    {
                //        sqlQuery = "update CallReport set CommChannel = '" + edt.CommChannel + "', DateReceived = GETDATE(), ReasonCode = '" + edt.ReasonCode + "', " +
                //            "CustomerName = '" + edt.CustomerName + "', CardNo = '" + edt.CardNo + "', Activity = '" + edt.Activity.Replace("'", "''") + "', Status = '" + edt.Status + "', " +
                //            "Agent = '" + edt.Agent + "', TicketNo = '" + edt.TicketNo + "', Type = 'Y', LastName = '" + edt.LastName + "', FirstName = '" + edt.FirstName + "', MiddleName = '" + edt.MiddleName + "', " +
                //            "InOutBound = '" + edt.InOutBound + "', TransType = '" + edt.TransType + "', Location = '" + edt.Location + "', ATMUsed = '" + edt.ATMUsed + "', PaymentOf = '" + edt.PaymentOf + "', " +
                //            "TerminalUsed = '" + edt.TerminalUsed + "', Merchant = '" + edt.Merchant + "', Inter = '" + edt.Inter + "',  BancnetUsed = '" + edt.BancnetUsed + "', Online = '" + ol + "', " +
                //            "Website = '" + edt.Website + "', Amount = " + amt + ", DateTrans = '" + edt.DateTrans + "', PLStatus = '" + edt.PLStatus + "', CardBlocked = '" + cb + "', EndoredTo = '" + edt.EndoredTo + "', " +
                //            "Destination = '" + edt.Destination + "', ATMTrans = '" + edt.ATMTrans + "', RemitFrom = '" + edt.RemitFrom + "', RemitConcern = '" + edt.RemitConcern + "', EndorsedFrom = '" + edt.EndorsedFrom + "', " +
                //            "Resolved = NULL, BillerName = '" + edt.BillerName + "', Remarks = '" + edt.Remarks + "', ReferedBy = '" + edt.ReferedBy + "', ContactPerson = '" + edt.ContactPerson + "', " +
                //            "Tagging = '" + edt.Tagging + "', CallTime = '" + edt.CallTime + "', LocalNum = '" + edt.LocalNum + "', Created_By = '" + edt.Created_By + "', Last_Update = '" + edt.Last_Update + "', " +
                //            "CardPresent = '" + edt.CardPresent + "', SubStatus = '" + edt.SubStatus + "', Currency = '" + edt.Currency + "', Branch = '" + edt.Branch + "' where RID = " + edt.RID + "";

                //    }
                //    else
                //    {

                //        sqlQuery = "update CallReport set CommChannel = '" + edt.CommChannel + "', DateReceived = GETDATE(), ReasonCode = '" + edt.ReasonCode + "', " +
                //            "CustomerName = '" + edt.CustomerName + "', CardNo = '" + edt.CardNo + "', Activity = '" + edt.Activity.Replace("'", "''") + "', Status = '" + edt.Status + "', " +
                //            "Agent = '" + edt.Agent + "', TicketNo = '" + edt.TicketNo + "', Type = 'Y', LastName = '" + edt.LastName + "', FirstName = '" + edt.FirstName + "', MiddleName = '" + edt.MiddleName + "', " +
                //            "InOutBound = '" + edt.InOutBound + "', TransType = '" + edt.TransType + "', Location = '" + edt.Location + "', ATMUsed = '" + edt.ATMUsed + "', PaymentOf = '" + edt.PaymentOf + "', " +
                //            "TerminalUsed = '" + edt.TerminalUsed + "', Merchant = '" + edt.Merchant + "', Inter = '" + edt.Inter + "',  BancnetUsed = '" + edt.BancnetUsed + "', Online = '" + ol + "', " +
                //            "Website = '" + edt.Website + "', Amount = " + amt + ", DateTrans = '" + edt.DateTrans + "', PLStatus = '" + edt.PLStatus + "', CardBlocked = '" + cb + "', EndoredTo = '" + edt.EndoredTo + "', " +
                //            "Destination = '" + edt.Destination + "', ATMTrans = '" + edt.ATMTrans + "', RemitFrom = '" + edt.RemitFrom + "', RemitConcern = '" + edt.RemitConcern + "', EndorsedFrom = '" + edt.EndorsedFrom + "', " +
                //            "Resolved = '" + edt.Resolved + "', BillerName = '" + edt.BillerName + "', Remarks = '" + edt.Remarks + "', ReferedBy = '" + edt.ReferedBy + "', ContactPerson = '" + edt.ContactPerson + "', " +
                //            "Tagging = '" + edt.Tagging + "', CallTime = '" + edt.CallTime + "', LocalNum = '" + edt.LocalNum + "', Created_By = '" + edt.Created_By + "', Last_Update = '" + edt.Last_Update + "', " +
                //            "CardPresent = '" + edt.CardPresent + "', SubStatus = '" + edt.SubStatus + "', Currency = '" + edt.Currency + "', Branch = '" + edt.Branch + "' where RID = " + edt.RID + "";

                //    }

                //}

                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception ex)
            {
                string x = ex.Message;
                return false;
            }
        }

        public IEnumerable<ATMOps> View_ATMActivity(ActivitySearch _sch)
        {

            string dt1 = (_sch.Date1 == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : _sch.Date1;

            string dt2 = (_sch.Date2 == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : _sch.Date2;

            string yr = DateTime.Now.Year.ToString();

            string strSQL;

            strSQL = "select RID, TicketNo, InOutBound, CommChannel, DateReceived, abbreviation, CustomerName, " +
                "CardNo, Destination, TransType, ATMTrans, case when isnumeric(Location) = 0 then Location else dbo.Get_ATMLocation(ATMUsed, Location) end Location, " +
                "case when isnumeric(ATMUsed) = 0 then ATMUsed else dbo.Get_ATMUsed(ATMUsed) end ATMUsed, PaymentOf, TerminalUsed, BillerName, " +
                "Merchant, Inter, BancnetUsed, Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, " +
                "PLStatus, CardBlocked, EndoredTo, Activity, Status, Agent, created_by, Last_Update, Resolved, Remarks, " +
                "ReferedBy, ContactPerson, ReasonCode, Tagging, ResolvedDate, CardPresent, SubStatus, Currency, Branch from( " +
                "select a.rid, TicketNo, InOutBound, CommChannel, dbo.Get_DateReceived(a.TicketNo) DateReceived, " +
                "a.reasoncode + ' - ' + abbreviation abbreviation, CustomerName, CardNo, Destination, TransType, ATMTrans, Location, " +
                "ATMUsed, PaymentOf, TerminalUsed, BillerName, Merchant, Inter, BancnetUsed, Online, Website, RemitFrom, RemitConcern, " +
                "Amount, DateTrans, PLStatus, CardBlocked, EndoredTo, Activity = rtrim(Activity), Status, Agent, resolved, Remarks, ReferedBy, " +
                "ContactPerson, a.reasoncode, a.Tagging, created_by, last_update, dbo.Get_ResovedDate(a.TicketNo) ResolvedDate, CardPresent, " +
                "SubStatus, Currency, brnnam Branch from ATMReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
                "left join Branch c on a.branch = c.brncd " +
                "where type = 'Y' ";

            if (_sch.DateSearch == false)
            {
                strSQL += "and convert(date, datereceived) between convert(date, '" + dt1 + "') and convert(date, '" + dt2 + "') ";
            }
            else
			{
                strSQL += "and year(DateReceived) between '2019' and '" + yr + "' ";
            }
            

            if (_sch.CustomerName != null)
            {
                strSQL += "and CustomerName like '%" + _sch.CustomerName.Trim() + "%' ";

            }

            if (_sch.CardNo != null)
            {
                strSQL += "and CardNo like '%" + _sch.CardNo.Trim() + "%' ";
            }

            if (_sch.CommType != "0")
            {
                if (_sch.CommType != null)
                {
                    strSQL += "and InOutBound = '" + _sch.CommType + "' ";
                }
            }

            if (_sch.CommChannel != "0")
            {
                if (_sch.CommChannel != null)
                {
                    strSQL += "and CommChannel = '" + _sch.CommChannel + "' ";
                }
            }

            if (_sch.ReasonCode != null)
            {
                strSQL += "and ReasonCode = '" + _sch.ReasonCode + "' ";
            }

            if (_sch.Status != "0")
            {
                if (_sch.Status != null)
                {
                    strSQL += "and Status = '" + _sch.Status + "' ";
                }
            }

            if (_sch.Agent != null)
            {
                strSQL += "and Created_By = '" + _sch.Agent + "' ";
            }

            if (_sch.TicketNo != null)
            {
                strSQL += "and TicketNo like '%" + _sch.TicketNo.Trim() + "%' ";
            }

            if (_sch.ReferredBy != null)
            {
                strSQL += "and ReferedBy like '%" + _sch.ReferredBy.Trim() + "%' ";
            }

            if (_sch.ContactPerson != null)
            {
                strSQL += "and ContactPerson like '%" + _sch.ContactPerson.Trim() + "%' ";
            }

            strSQL += ") x order by x.DateReceived, TicketNo";

            return this.db.Query<ATMOps>(strSQL, new { }, commandTimeout: 0);


        }

        public bool ChkATMUsed(string _atmused)
        {
            int cnt = this.db.Query<ATMUsed>("select distinct ATM_ID from ATM_GROUP_LOC where CONVERT(varchar, ATM_ID) = '" + _atmused + "'", new { }).Count();

            if (cnt == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public EditATMReport ViewATMEntry(int _id)
        {
            string sql = "select RID,TicketNo,CommChannel,ReasonCode,CustomerName,CardNo,Activity,Status,Agent,LastName,FirstName,MiddleName,InOutBound,TransType, " +
                "Location,ATMUsed,PaymentOf,TerminalUsed,Merchant,Inter,BancnetUsed,case when Online = 'N' then 0 else 1 end Online,Website,Amount,case when convert(date, DateTrans) = '1900-01-01' then null else FORMAT(DateTrans,'dd-MMM-yyyy') end DateTrans, " +
                "PLStatus,case when CardBlocked = 'No' then 0 else 1 end CardBlocked,EndoredTo,Destination, " +
                "ATMTrans,RemitFrom,RemitConcern, " +
                "BillerName,Remarks,ReferedBy,ContactPerson,Tagging, " +
                "CardPresent,SubStatus,Currency,Branch,NameVerified,Resolved,Created_By from ATMReport where rid = " + _id + "";

            return this.db.Query<EditATMReport>(sql, new { }, commandTimeout: 0 ).FirstOrDefault();
        }

        public IEnumerable<ATMOps> View_TicketATM(string _ticketno)
        {

            string strSQL;

            //strSQL = "select RID, TicketNo, InOutBound, CommChannel, DateReceived, abbreviation, CustomerName, " +
            //    "CardNo, Destination, TransType, ATMTrans, case when isnumeric(Location) = 0 then Location else dbo.Get_ATMLocation(ATMUsed, Location) end Location, " +
            //    "case when isnumeric(ATMUsed) = 0 then ATMUsed else dbo.Get_ATMUsed(ATMUsed) end ATMUsed, PaymentOf, TerminalUsed, BillerName, " +
            //    "Merchant, Inter, BancnetUsed, Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, " +
            //    "PLStatus, CardBlocked, EndoredTo, Activity, Status, Agent, created_by, Last_Update, Resolved, Remarks, " +
            //    "ReferedBy, ContactPerson, ReasonCode, Tagging, ResolvedDate, CardPresent, SubStatus, Currency, Branch from( " +
            //    "select a.rid, TicketNo, InOutBound, CommChannel, dbo.Get_DateReceived(a.TicketNo) DateReceived, a.reasoncode + ' - ' + abbreviation abbreviation, CustomerName, " +
            //    "CardNo, Destination, TransType, ATMTrans, Location, ATMUsed, PaymentOf, TerminalUsed, BillerName, Merchant, Inter, BancnetUsed, " +
            //    "Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, PLStatus, CardBlocked, EndoredTo, Activity = rtrim(Activity), " +
            //    "Status, Agent, resolved, Remarks, ReferedBy, ContactPerson, a.reasoncode, a.Tagging, created_by, last_update, " +
            //    "dbo.Get_ResovedDate(a.TicketNo) ResolvedDate, CardPresent, SubStatus, Currency, brnnam Branch " +
            //    "from ATMReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
            //    "left join Branch c on a.branch = c.brncd " +
            //    "where type = 'Y' and TicketNo = '" + _ticketno + "' " +
            //    ") x order by x.DateReceived, TicketNo";

            strSQL = "SELECT DISTINCT * FROM (" +
                "select TicketNo, InOutBound, CommChannel, DATEADD(MILLISECOND, -DATEPART(MILLISECOND, DateReceived), DateReceived) DateReceived, abbreviation, CustomerName, CardNo, Destination, " +
                "TransType, ATMTrans, case when isnumeric(Location) = 0 then Location else dbo.Get_ATMLocation(ATMUsed, Location) end Location, " +
                "case when isnumeric(ATMUsed) = 0 then ATMUsed else dbo.Get_ATMUsed(ATMUsed) end ATMUsed, PaymentOf, TerminalUsed, BillerName, " +
                "Merchant, Inter, BancnetUsed, Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, " +
                "PLStatus, CardBlocked, EndoredTo, Activity, Status, Agent, Created_By, Last_Update, Resolved, Remarks, " +
                "ReferedBy, ContactPerson, ReasonCode, Tagging, ResolvedDate, CardPresent, SubStatus, Currency, Branch from ( " +
                "    select a.rid, TicketNo, InOutBound, CommChannel, DateReceived, a.reasoncode + ' - ' + abbreviation abbreviation, CustomerName, " +
                "    CardNo, Destination, TransType, ATMTrans, Location, ATMUsed, PaymentOf, TerminalUsed, BillerName, Merchant, Inter, BancnetUsed, " +
                "    Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, PLStatus, CardBlocked, EndoredTo, Activity = rtrim(Activity), " +
                "    Status, Agent, resolved, Remarks, ReferedBy, ContactPerson, a.reasoncode, a.Tagging, created_by, last_update, " +
                "    (select isnull(DateReceived, ' ') from CallReport where ticketno = a.TicketNo and[TYPE] = 'Y' and[Status] in ('Resolved', 'Closed', 'Cancelled')) ResolvedDate, " +
                "	 CardPresent, SubStatus, Currency, brnnam Branch " +
                "    from ATMReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
                "    left join Branch c on a.branch = c.brncd " +
                "    where TicketNo = '" + _ticketno + "' " +
                ") x " +
                "union all " +
                "select TicketNo, InOutBound, CommChannel, DATEADD(MILLISECOND, -DATEPART(MILLISECOND, DateReceived), DateReceived) DateReceived, abbreviation, CustomerName,  CardNo, Destination, TransType, " +
                "ATMTrans, case when isnumeric(Location) = 0 then Location else dbo.Get_ATMLocation(ATMUsed, Location) end Location, " +
                "case when isnumeric(ATMUsed) = 0 then ATMUsed else dbo.Get_ATMUsed(ATMUsed) end ATMUsed, PaymentOf, TerminalUsed, BillerName, " +
                "Merchant, Inter, BancnetUsed, Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, " +
                "PLStatus, CardBlocked, EndoredTo, Activity, Status, Agent, created_by, Last_Update, Resolved, Remarks, " +
                "ReferedBy, ContactPerson, ReasonCode, Tagging, ResolvedDate, CardPresent, SubStatus, Currency, Branch from ( " +
                "select a.rid, TicketNo, InOutBound, CommChannel, dbo.Get_DateReceived(a.TicketNo) DateReceived, a.reasoncode + ' - ' + abbreviation abbreviation, CustomerName, " +
                "CardNo, Destination, TransType, ATMTrans, Location, ATMUsed, PaymentOf, TerminalUsed, BillerName, Merchant, Inter, BancnetUsed, " +
                "Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, PLStatus, CardBlocked, EndoredTo, Activity = rtrim(Activity), " +
                "Status, Agent, resolved, Remarks, ReferedBy, ContactPerson, a.reasoncode, a.Tagging, created_by, last_update, " +
                "(select isnull(DateReceived, ' ') from CallReport where ticketno = a.TicketNo and[TYPE] = 'Y' and[Status] in ('Resolved', 'Closed', 'Cancelled')) ResolvedDate,  " +
                "CardPresent, SubStatus, Currency, brnnam Branch " +
                "from CallReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
                "left join Branch c on a.branch = c.brncd " +
                "where TicketNo = '" + _ticketno + "' " +
                ") x ) y order by y.DateReceived, y.TicketNo ";
                


            return this.db.Query<ATMOps>(strSQL, new { }, commandTimeout: 0);


        }

        public bool CreateATMActivity(string CommChannel, string ReasonCode, string CustomerName, string CardNo,
            string Activity, string Status, string Agent, string TicketNo, string LastName, string FirstName, string MiddleName, string InOutBound,
            string TransType, string Location, string ATMUsed, string PaymentOf, string TerminalUsed, string Merchant, string Inter, string BancnetUsed, string Online,
            string Website, double Amount, string DateTrans, string PLStatus, string CardBlocked, string EndoredTo, string Destination, string ATMTrans,
            string RemitFrom, string RemitConcern, string BillerName, string Remarks, string ReferedBy,
            string ContactPerson, string Tagging, string CardPresent, string SubStatus, string Currency, string Branch, string NameVerified, string Resolved, string Last_Update, string Created_By
         )
        {
            try
            {

                string ol = (Online == "true") ? "Y" : "N";
                string cb = (CardBlocked == "true") ? "Yes" : "No";

                string sqlQuery = "";

                if (DateTrans == null)
                {
                    sqlQuery = "INSERT INTO ATMReport (CommChannel,DateReceived,ReasonCode,CustomerName,CardNo,Activity,Status,Agent, " +
                        "TicketNo,Type,LastName,FirstName,MiddleName,InOutBound,TransType,Location,ATMUsed,PaymentOf,TerminalUsed,Merchant,Inter, " +
                        "BancnetUsed,Online,Website,Amount,DateTrans,PLStatus,CardBlocked,EndoredTo,Destination,ATMTrans,RemitFrom,RemitConcern, " +
                        "BillerName,Remarks,ReferedBy,ContactPerson,Tagging, " +
                        "CardPresent,SubStatus, Currency,Branch,NameVerified,Resolved,Last_Update,Created_By) " +
                        "VALUES('" + CommChannel + "', GETDATE(), '" + ReasonCode + "', '" + CustomerName + "', '" + CardNo + "', '" + Activity.Replace("'", "''") + "', '" + Status + "', '" + Agent + "', '" + TicketNo + "' " +
                        ", 'Y', '" + LastName + "', '" + FirstName + "', '" + MiddleName + "', '" + InOutBound + "', '" + TransType + "', '" + Location + "', '" + ATMUsed + "', '" + PaymentOf + "', '" + TerminalUsed + "', '" + Merchant + "', '" + Inter + "' " +
                        ", '" + BancnetUsed + "', '" + ol + "', '" + Website + "', " + Amount + ", null, '" + PLStatus + "', '" + cb + "', '" + EndoredTo + "', '" + Destination + "', '" + ATMTrans + "', '" + RemitFrom + "', '" + RemitConcern + "' " +
                        ", '" + BillerName + "', '" + Remarks + "', '" + ReferedBy + "', '" + ContactPerson + "', '" + Tagging + "' " +
                        ", '" + CardPresent + "', '" + SubStatus + "', '" + Currency + "', '" + Branch + "', '" + NameVerified + "', Convert(date, '" + Resolved + "'), '" + Last_Update + "', '" + Created_By + "')";

                }
                else
                {
                    sqlQuery = "INSERT INTO ATMReport (CommChannel,DateReceived,ReasonCode,CustomerName,CardNo,Activity,Status,Agent, " +
                        "TicketNo,Type,LastName,FirstName,MiddleName,InOutBound,TransType,Location,ATMUsed,PaymentOf,TerminalUsed,Merchant,Inter, " +
                        "BancnetUsed,Online,Website,Amount,DateTrans,PLStatus,CardBlocked,EndoredTo,Destination,ATMTrans,RemitFrom,RemitConcern, " +
                        "BillerName,Remarks,ReferedBy,ContactPerson,Tagging, " +
                        "CardPresent,SubStatus,Currency,Branch,NameVerified,Resolved,Last_Update,Created_By) " +
                        "VALUES('" + CommChannel + "', GETDATE(), '" + ReasonCode + "', '" + CustomerName + "', '" + CardNo + "', '" + Activity.Replace("'", "''") + "', '" + Status + "', '" + Agent + "', '" + TicketNo + "' " +
                        ", 'Y', '" + LastName + "', '" + FirstName + "', '" + MiddleName + "', '" + InOutBound + "', '" + TransType + "', '" + Location + "', '" + ATMUsed + "', '" + PaymentOf + "', '" + TerminalUsed + "', '" + Merchant + "', '" + Inter + "' " +
                        ", '" + BancnetUsed + "', '" + ol + "', '" + Website + "', " + Amount + ", '" + DateTrans + "', '" + PLStatus + "', '" + cb + "', '" + EndoredTo + "', '" + Destination + "', '" + ATMTrans + "', '" + RemitFrom + "', '" + RemitConcern + "' " +
                        ", '" + BillerName + "', '" + Remarks + "', '" + ReferedBy + "', '" + ContactPerson + "', '" + Tagging + "' " +
                        ", '" + CardPresent + "', '" + SubStatus + "', '" + Currency + "', '" + Branch + "', '" + NameVerified + "', Convert(date, '" + Resolved + "'), '" + Last_Update + "', '" + Created_By + "')";
                }

                db.Execute(sqlQuery, new { }, commandTimeout: 0);

                return true;
            }
            catch (Exception ex)
            {
                string x = ex.Message;
                return false;
            }
        }

        public IEnumerable<ATMOps> ListATMOpsResolved(string _dt1, string _dt2)
        {
            return this.db.Query<ATMOps>("sp_ATMReportResolvedNew", new { Date1 = _dt1, Date2 = _dt2 }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<ATMOps> ListATMOpsUnResolved(string _dt1, string _dt2)
        {
            return this.db.Query<ATMOps>("sp_ATMReportUnResolvedNew", new { Date1 = _dt1, Date2 = _dt2 }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<CMActivity> View_ReasonCodes(string _dt1, string _dt2, string _rno)
        {

            string strSQL;

            strSQL = "select RID, TicketNo, InOutBound, CommChannel, DateReceived, abbreviation, CustomerName, " +
                     "CardNo, Destination, TransType, ATMTrans, Location, ATMUsed, PaymentOf, TerminalUsed, BillerName, " +
                     "Merchant, Inter, BancnetUsed, Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, " +
                     "PLStatus, CardBlocked, EndoredTo, Activity, Status, Agent, created_by, Last_Update, EndorsedFrom, Resolved, Remarks, " +
                     "ReferedBy, ContactPerson, ReasonCode, Tagging, ResolvedDate, CallTime, LocalNum, CardPresent from ( " +
                     "select a.rid, TicketNo, InOutBound, CommChannel, " +
                     "(select DateReceived from CallReport where RID in (select MIN(rid) from CallReport where ticketno = a.ticketno)) DateReceived, " +
                     "a.reasoncode + ' - ' + abbreviation abbreviation, CustomerName, " +
                     "CardNo, Destination, TransType, ATMTrans, Location, ATMUsed, PaymentOf, TerminalUsed, BillerName, Merchant, Inter, BancnetUsed, " +
                     "Online, Website, RemitFrom, RemitConcern, Amount, DateTrans, PLStatus, CardBlocked, EndoredTo, Activity = rtrim(Activity), " +
                     "Status, Agent, EndorsedFrom, resolved, Remarks, ReferedBy, ContactPerson, a.reasoncode, " +
                     "a.Tagging, created_by, last_update, " +
                     "(select isnull(DateReceived, ' ') from CallReport where ticketno = a.TicketNo and [TYPE] = 'Y' and [Status] in ('Resolved', 'Closed', 'Cancelled')) ResolvedDate, CallTime, LocalNum, CardPresent " +
                     "from CallReport a inner join ReasonNew b on a.reasoncode = b.rcode " +
                     "where type = 'Y' ";

            if (_rno == "X")
            {
                strSQL += "and convert(date, datereceived) between '" + _dt1 + "' and '" + _dt1 + "' ";
            }
            else
            {
                strSQL += "and ReasonCode in (" + _rno + ") ";
            }

            strSQL += ") x order by x.DateReceived, TicketNo";

            return this.db.Query<CMActivity>(strSQL, new { }, commandTimeout: 0);

            //return this.db.Query<CMActivity>("sp_CMActivity", new { Date1 = _dt1, Date2 = _dt2, Mod = _mod }, commandType: CommandType.StoredProcedure);

        }

        public MaxRID Get_RID(string _ticketno, string _dept)
        {
            return this.db.Query<MaxRID>("select dbo.Get_MaxRID('" + _ticketno + "', '" + _dept + "') as RID", new { }).FirstOrDefault();
        }

        public bool UpdateLocation(string _ticketno)
        {
            try
            {
                string sqlQuery = "UPDATE CallReport SET Location = b.Location, ReferedBy = b.ReferedBy from CallReport a inner join ( " +
                    "select a.TicketNo, Location, ReferedBy from CallReport a inner join( " +
                    "select min(RID) RID, TicketNo from CallReport where TicketNo = '" + _ticketno + "' " +
                    "group by TicketNo) b on a.rid = b.rid) b on a.TicketNo = b.TicketNo";
                db.Execute(sqlQuery, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region Report

        public IEnumerable<Combo1> ListAgent()
        {
            return this.db.Query<Combo1>("select LoginID ID, LoginID Description from UserName where dept in ('A', 'M') order by id", new { });
        }

        public IEnumerable<Summary> View_ReportSummary(string StartDate, string EndDate, string Bound, string Agent, string Type, string Comm)
        {
            string dt1 = (StartDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : StartDate;

            string dt2 = (EndDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : EndDate;

            return this.db.Query<Summary>("sp_ReportSummary", new { StartDate = dt1, EndDate = dt2, Bound = Bound, Agent = Agent, Type = Type, Comm = Comm }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<ReportActivity> View_ActivityLogs(string StartDate, string EndDate, string comm_channel, string comm_type, string agent, string dept)
        {
            string dt1 = (StartDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : StartDate;

            string dt2 = (EndDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : EndDate;

            return this.db.Query<ReportActivity>("sp_ReportActivityNew", new { StartDate = dt1, EndDate = dt2, ComChan = comm_channel, Bound = comm_type, Agent = agent, Type = dept }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<ReasonComm> View_ReasonComm(string StartDate, string EndDate, string agent, string dept)
        {
            string dt1 = (StartDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : StartDate;

            string dt2 = (EndDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : EndDate;

            return this.db.Query<ReasonComm>("sp_ReportMRComm", new { StartDate = dt1, EndDate = dt2, Agent = agent, Type = dept }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<ReasonStat> View_ReasonStat(string StartDate, string EndDate, string agent, string dept)
        {
            string dt1 = (StartDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : StartDate;

            string dt2 = (EndDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : EndDate;

            return this.db.Query<ReasonStat>("sp_ReportMRStat", new { StartDate = dt1, EndDate = dt2, Agent = agent, Type = dept }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<Aging> View_Aging(string StartDate, string EndDate, string RCode, string Type, string Branch)
        {
            string dt1 = (StartDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : StartDate;

            string dt2 = (EndDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : EndDate;

            return this.db.Query<Aging>("sp_Aging3", new { StartDate = dt1, EndDate = dt2, RCode = RCode, Type = Type, Branch = Branch }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<AgingDetail> View_AgingDetail(string StartDate, string EndDate, string RCode, string Type, string Branch)
        {
            string dt1 = (StartDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : StartDate;

            string dt2 = (EndDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : EndDate;

            return this.db.Query<AgingDetail>("sp_AgingTAT", new { StartDate = dt1, EndDate = dt2, RCode = RCode, Type = Type, Branch = Branch }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<Aging> View_AgingResolved(string StartDate, string EndDate, string RCode, string Type)
        {
            string dt1 = (StartDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : StartDate;

            string dt2 = (EndDate == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : EndDate;

            return this.db.Query<Aging>("sp_Aging3_Resolved", new { StartDate = dt1, EndDate = dt2, RCode = RCode, Type = Type }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<ReportIRC> List_ReportIRC(string dt1, string dt2, string typ)
        {
            return this.db.Query<ReportIRC>("sp_IRCReasonCodes", new { Date1 = dt1, Date2 = dt2, Type = typ }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<CMActivity> List_ReportIRCDetail(string dt1, string dt2, string typ)
        {
            return this.db.Query<CMActivity>("sp_IRCReasonCodesDetails", new { Date1 = dt1, Date2 = dt2, Type = typ }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

        public IEnumerable<ReportIRC> List_ReportIRCATM(string dt1, string dt2, string typ)
        {
            string det1 = (dt1 == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : dt1;

            string det2 = (dt2 == null) ? DateTime.Now.ToString("dd-MMM-yyyy") : dt2;

            return this.db.Query<ReportIRC>("sp_IRCReasonCodesATM", new { Date1 = det1, Date2 = det2, Type = typ }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }


        #endregion

        #region Inquiry
        public IEnumerable<InquiryICBA> ListInquiryICBA(string _fil, string _cardno, string _fname, string _lname, string _fullname)
        {
            return this.db.Query<InquiryICBA>("sp_ICBAInq", new { Fil = _fil, CardNumber = _cardno, FirstName = _fname, LastName = _lname, FullName = _fullname }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }


        public IEnumerable<MasterDetailFile> GetLoanDetails(string _acctno, string _cifkey)
        {
            db.Open();
            var objDetails = SqlMapper.QueryMultiple(db, "GetLoanDetails", new { accountno = _acctno, cifkey = _cifkey }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

            MasterDetailFile ObjMaster = new MasterDetailFile();

            //Assigning each Multiple tables data to specific single model class  
            ObjMaster.vwCustomerInfo = objDetails.Read<displayCustomer>().ToList();
            ObjMaster.vwAccountInfo = objDetails.Read<displayAccountInfo>().ToList();
            ObjMaster.vwAccountCollateral = objDetails.Read<displayAccountCollateral>().ToList();
            ObjMaster.vwOSBalance = objDetails.Read<displayOSBalance>().ToList();
            ObjMaster.vwCollection = objDetails.Read<displayCollection>().ToList();
            ObjMaster.vwPostDated = objDetails.Read<displayPostDated>().ToList();
            ObjMaster.vwOverdue = objDetails.Read<displayOverdue>().ToList();
            ObjMaster.vwRepayment = objDetails.Read<displayRepayment>().ToList();
            ObjMaster.vwRepaymentTable = objDetails.Read<displayRepaymentTable>().ToList();
            ObjMaster.vwRepaymentDetail = objDetails.Read<displayRepaymentDetail>().ToList();
            ObjMaster.vwLoanHistory = objDetails.Read<displayLoanHist>().ToList();
            ObjMaster.vwLoanSummary = objDetails.Read<displayLoanSummary>().ToList();
            ObjMaster.vwLoanInitRate = objDetails.Read<displayLNInitialRate>().ToList();

            List<MasterDetailFile> CustomerObj = new List<MasterDetailFile>();
            //Add list of records into MasterDetails list  
            CustomerObj.Add(ObjMaster);
            db.Close();

            return CustomerObj;

        }

        public IEnumerable<MasterCASAFile> GetCASADetails(string _acctno, string _cifkey)
        {
            db.Open();
            var objDetails = SqlMapper.QueryMultiple(db, "GetCASADetails", new { accountno = _acctno, cifkey = _cifkey }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

            MasterCASAFile ObjMaster = new MasterCASAFile();

            //Assigning each Multiple tables data to specific single model class  
            ObjMaster.vwCustomerInfo = objDetails.Read<displayCustomer>().ToList();
            ObjMaster.vwTotalBalance = objDetails.Read<dispCASATotal>().ToList();
            ObjMaster.vwCASAHistory = objDetails.Read<dispCASAHistory>().ToList();

            List<MasterCASAFile> CustomerObj = new List<MasterCASAFile>();
            //Add list of records into MasterDetails list  
            CustomerObj.Add(ObjMaster);
            db.Close();

            return CustomerObj;

        }

        public IEnumerable<MasterFDFile> GetFDDetails(string _acctno, string _cifkey)
        {
            db.Open();
            var objDetails = SqlMapper.QueryMultiple(db, "GetFixedDetails", new { accountno = _acctno, cifkey = _cifkey }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

            MasterFDFile ObjMaster = new MasterFDFile();

            //Assigning each Multiple tables data to specific single model class  
            ObjMaster.vwCustomerInfo = objDetails.Read<displayCustomer>().ToList();
            ObjMaster.vwCertInfo = objDetails.Read<dispFDCertInfo>().ToList();
            ObjMaster.vwFDHistory = objDetails.Read<dispFDHistory>().ToList();
            ObjMaster.vwFDSummary = objDetails.Read<dispFDSummary>().ToList();

            List<MasterFDFile> CustomerObj = new List<MasterFDFile>();
            //Add list of records into MasterDetails list  
            CustomerObj.Add(ObjMaster);
            db.Close();

            return CustomerObj;

        }

        //BANKWORKS
        public IEnumerable<InquiryBW> ListInquiryBW(string _fil, string _cardno, string _fname, string _lname)
        {
            return this.db.Query<InquiryBW>("sp_BWPrepaid", new { Fil = _fil, CardNumber = _cardno, FirstName = _fname, LastName = _lname }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }


        public IEnumerable<MasterBWFile> GetBWDetails(string _acctno, string _cardno, string _st, string _et)
        {
            db.Open();
            var objDetails = SqlMapper.QueryMultiple(db, "GetBWDetails", new { accountno = _acctno, cardno = _cardno, start = _st, end = _et }, commandTimeout: 0, commandType: CommandType.StoredProcedure);

            MasterBWFile ObjMaster = new MasterBWFile();

            //Assigning each Multiple tables data to specific single model class  
            ObjMaster.BWClientInfo = objDetails.Read<bwClienInfo>().ToList();
            ObjMaster.BWAddress = objDetails.Read<bwAddress>().ToList();
            ObjMaster.BWAccountDetails = objDetails.Read<bwAccountDetails>().ToList();
            ObjMaster.BWHistory = objDetails.Read<bwHistory>().ToList();
            ObjMaster.BWCardUsage = objDetails.Read<bwCardUsage>().ToList();
            ObjMaster.BWOnlineAuth = objDetails.Read<bwAuthorization>().ToList();
            ObjMaster.BWContactInfo = objDetails.Read<bwContactInfo>().ToList();
            ObjMaster.BWStatus = objDetails.Read<bwStatus>().ToList();
            ObjMaster.BWBalance = objDetails.Read<bwBalance>().ToList();

            List<MasterBWFile> CustomerObj = new List<MasterBWFile>();
            //Add list of records into MasterDetails list  
            CustomerObj.Add(ObjMaster);
            db.Close();

            return CustomerObj;

        }

        public IEnumerable<bwAuthorization> Get_BWOnlineAuth(string _cardno, string _start, string _end)
        {

            string strSQL = "select a.transaction_date, a.time_transaction, a.merchant_name, b.client_country merchant_country,  decode(a.dr_cr_indicator, '002', 'Debit', 'Credit') dr_cr_indicator, c.transaction_type, a.account_amount_gr, d.swift_code acct_currency, " +
                             "e.response, a.tran_amount_gr, d.swift_code tran_currency, a.rate_fx_local_acct, f.transaction_status, decode(a.reversal_flag, '000', 'No', 'Yes') reversal_flag, g.transaction_category, h.transaction_class, " +
                             "substr(a.CARD_NUMBER,1, 4)||'********'||substr(a.CARD_NUMBER, 13, 4) card_number, a.auth_code, i.business_class, j.clearing_channel transaction_source, a.account_amount_net, a.account_amount_chg, a.transaction_slip, a.retrieval_reference, a.transaction_id, (select PE.PAN_ENTRY_DESC from bw3.CHT_PAN_ENTRY pe where pe.index_field = a.pan_entry) as pan_entry  " +
                             "from bw3.int_online_transactions a " +
                             "left outer join bw3.BWT_country b on a.merchant_country = b.index_field and a.institution_number = b.institution_number " +
                             "left outer join bw3.BWT_TRANSACTION_TYPE c on a.transaction_type = c.index_field and a.institution_number = c.institution_number " +
                             "LEFT OUTER JOIN bw3.cht_currency d on a.acct_currency = d.iso_code " +
                             "left outer join bw3.CHT_RESPONSE e on a.response_code = e.index_field " +
                             "left outer join bw3.CHT_TRANSACTION_STATUS f on a.transaction_status = f.index_field " +
                             "left outer join bw3.CHT_TRANSACTION_CATEGORY g on a.transaction_category = g.index_field " +
                             "left outer join bw3.CHT_TRANSACTION_CLASS h on a.transaction_class = h.index_field " +
                             "left outer join bw3.BWT_ISO_BUSS_CLASS i on a.business_class = i.index_field and a.institution_number = i.institution_number " +
                             "left outer join bw3.CHT_CLEARING_CHANNEL j on a.transaction_source = j.index_field " +
                             "where a.card_number = '" + _cardno + "' AND a.transaction_date between '" + _start + "' and '" + _end + "' order by a.transaction_date, a.time_transaction";

            IDbConnection conn = ConnectionManager.GetConnection("B");

            var result = conn.Query<bwAuthorization>(strSQL, new { });

            ConnectionManager.CloseConnection(conn);

            return result;

        }

        public bool ImportExcel(string _path)
        {
            try
            {

				string sqlQry = "truncate table CardNo";
				db.Execute(sqlQry, new { });

                //int rowno = 1;
                //            XLWorkbook workbook = XLWorkbook.OpenFromTemplate(_path);
                //            var sheets = workbook.Worksheets.First();
                //            var rows = sheets.Rows().ToList();
                //            foreach (var row in rows)
                //            {
                //                if (rowno != 1)
                //                {
                //                    var cardno = row.Cell(1).Value.ToString();
                //                    if (string.IsNullOrWhiteSpace(cardno) || string.IsNullOrEmpty(cardno))
                //                    {
                //                        break;
                //                    }

                //                    string sql = "INSERT INTO CardNo (Cardno) " +
                //				"VALUES('" + cardno + "')";

                //		db.Execute(sql, new { });
                //	}
                //                else
                //                {
                //                    rowno = 2;
                //                }
                //            }

                using (var package = new ExcelPackage(_path))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 1; row <= rowCount; row++)
                    {

                        var cardno = worksheet.Cells[row, 1].Value.ToString().Trim();
                        if (string.IsNullOrWhiteSpace(cardno) || string.IsNullOrEmpty(cardno))
						{
							break;
						}

                        string sql = "INSERT INTO CardNo (Cardno) " +
									 "VALUES('" + cardno.Replace("'", "") + "')";

						db.Execute(sql, new { });

                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                string msg = ex.ToString();
                return false;
            }
        }

        public bool ClearCards()
        {

            try
            {

                string sqlQry = "truncate table CardNo";
                db.Execute(sqlQry, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool InserToCard(string _cardno)
		{

            try
            {

                string sql = "INSERT INTO CardNo (Cardno) " +
                             "VALUES('" + _cardno + "')";

                db.Execute(sql, new { });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string Documentupload(IFormFile fromFiles)
        {
            string uploadpath = _environment.WebRootPath;
            string dest_path = Path.Combine(uploadpath, "upload");

            if (!Directory.Exists(dest_path))
            {
                Directory.CreateDirectory(dest_path);
            }

            string sourcefile = Path.GetFileName(fromFiles.FileName);
            string path = Path.Combine(dest_path, sourcefile);

            if (File.Exists(path))
			{
                File.Delete(path);
			}

            using (FileStream filestream = new FileStream(path, FileMode.Create))
            {
                fromFiles.CopyTo(filestream);
            }
            return path;
        }

        public IEnumerable<InquiryBW> ListInquiryBWIRemit()
        {
            return this.db.Query<InquiryBW>("sp_BWPrepaidIRemit", new { }, commandTimeout: 0, commandType: CommandType.StoredProcedure);
        }

		#endregion

		public string GetIPAdd()
		{
            //IHttpContextAccessor currentRequest = IHttpContextAccessor.HttpContext.Current.Request;
            //string ipAddress = currentRequest.ServerVariables["HTTP_X_FORWARDED_FOR"];

            //if (ipAddress == null || ipAddress.ToLower() == "unknown")
            //	ipAddress = currentRequest.ServerVariables["REMOTE_ADDR"];

            var remoteIpAddress = _httpContextAccessor.HttpContext.Connection.RemoteIpAddress;

            if (remoteIpAddress != null)
            {
                // Check if the request is coming from a proxy
                if (_httpContextAccessor.HttpContext.Request.Headers.TryGetValue("X-Forwarded-For", out var forwardedFor))
                {
                    var forwardedIps = forwardedFor.ToString().Split(',');

                    // Use the first non-local IP address from the forwarded headers
                    foreach (var ip in forwardedIps)
                    {
                        if (IPAddress.TryParse(ip.Trim(), out var parsedIpAddress)
                            && !IPAddress.IsLoopback(parsedIpAddress))
                        {
                            remoteIpAddress = parsedIpAddress;
                            break;
                        }
                    }
                }
            }

            var clientIpAddress = remoteIpAddress?.ToString();

            return clientIpAddress;
		}



		public void Dispose()
		{
			db.Close();
			db.Dispose();
		}
	}
}
