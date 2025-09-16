using System;
using System.ComponentModel.DataAnnotations;

namespace AgentDesktop.Models
{
	public class Login
	{
		[Required]
		public string login_username { get; set; } = string.Empty;

		[Required]
		[DataType(DataType.Password)]
		public string login_password { get; set; } = string.Empty;
	}

	public class chkLogin
	{
		public string username { get; set; } = string.Empty;

		public string Password { get; set; } = string.Empty;
	}

	public class ForgotPass
	{
		[Required]
		[Display(Name = "User Name")]
		public string Username { get; set; } = string.Empty;
	}

	public class EmailForgot
	{
		public int ID { get; set; }
		public string FullName { get; set; }
		public string LoginID { get; set; }
		public string Email { get; set; }
	}

	public class LoginPass
	{
		[Key]
		public int ID { get; set; }
		public string FullName { get; set; }
		public string LoginID { get; set; }
		public int Active { get; set; }
		public int GroupID { get; set; }
		public string Email { get; set; }
		public string Dept { get; set; }
		public int Login { get; set; }
		public string chPass { get; set; }
		public string Password2 { get; set; }
		public int ExpiryDate { get; set; }

	}

	public class SessionModel
	{
		public int ID { get; set; }
		public string FullName { get; set; }
		public string LoginID { get; set; }
		public int UType { get; set; }
		public string chPass { get; set; }
		public string WrkStn { get; set; }
		public string Dept { get; set; }
		public string Stat { get; set; }

	}

	public class ResetPassword
	{

		[Required]
		[Display(Name = "Old Password")]
		[DataType(DataType.Password)]
		public string OldPassword { get; set; }

		[Required]
		[Display(Name = "New Password")]
		[DataType(DataType.Password)]
		//[RegularExpression("((?=.*\\d)(?=.*[A-Z])(?=.*[a-z])(?=.*\\W).{8,15})", ErrorMessage = "Password must have at least 8 characters and contain at least 1 uppercase letter, 1 lowercase letter, 1 numeric, and 1 special character.")]
		//[RegularExpression(@"^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[^a-zA-Z0-9])(?!.*(.).*\1).{8,15}$", ErrorMessage = "Password must have at least 8 characters and contain at least 1 uppercase letter, 1 lowercase letter, 1 numeric character, 1 special character, and cannot contain repeated characters.")]
		//[RegularExpression(@"^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[^a-zA-Z0-9])(?!.*(aaa|bbb|ccc|ddd|eee|fff|ggg|hhh|iii|jjj|kkk|lll|mmm|nnn|ooo|ppp|qqq|rrr|sss|ttt|uuu|vvv|www|xxx|yyy|zzz|111|222|333|444|555|666|777|888|999|000)).{8,15}$", ErrorMessage = "Password must have at least 8 characters and contain at least 1 uppercase letter, 1 lowercase letter, 1 numeric character, 1 special character, and cannot contain consecutive sequential letters or numbers.")]
		[RegularExpression(@"^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[^a-zA-Z0-9])(?!.*(.)\1\1)[^\s]{8,}$", ErrorMessage = "Password must have at least 8 characters and contain at least 1 uppercase letter, 1 lowercase letter, 1 numeric character, 1 special character, and cannot contain 3 consecutive characters.")]
		public string Password { get; set; }

		[Required]
		[Display(Name = "Confirm New Password")]
		[DataType(DataType.Password)]
		[Compare("Password", ErrorMessage = "Password did not match.")]
		public string ConfirmPassword { get; set; }
	}

	public class PassList
	{
		public string pass { get; set; }
	}
	public class CountPass
	{
		public int oldpass { get; set; }
		public int count_pass { get; set; }
	}


}
