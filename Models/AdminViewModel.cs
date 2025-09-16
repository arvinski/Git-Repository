using System;
using System.ComponentModel.DataAnnotations;

namespace AgentDesktop.Models
{
	public class AdminViewModel
	{
	}

	public class ViewUsers
	{
		public int id { get; set; }
		public string loginid { get; set; }
		public string fullname { get; set; }
		public string group { get; set; }
		public string email { get; set; }
		public string department { get; set; }
		public string stat { get; set; }

		public string expirydate { get; set; }

	}

    public class NewUser
    {
        [Required]
        [StringLength(20, MinimumLength = 5, ErrorMessage = "User ID should be between 5 and 20 characters")]
        [Display(Name = "User ID")]
        public string loginid { get; set; }

        [Required]
        [Display(Name = "User Name")]
        public string fullname { get; set; }

        [Required]
        [Display(Name = "Group")]
        public int group { get; set; }

        [Required]
        [EmailAddress]
        [Display(Name = "Email")]
        public string email { get; set; }

        [Required]
        [Display(Name = "Department")]
        public string department { get; set; }

    }

    public class EditUser
    {
		[Key]
		public int id { get; set; }

		[Required]
        [Display(Name = "User ID")]
        public string loginid { get; set; }

        [Required]
        [Display(Name = "User Name")]
        public string fullname { get; set; }

        [Required]
        [Display(Name = "Group")]
        public int group { get; set; }

        [Required]
        [EmailAddress]
        [Display(Name = "Email")]
        public string email { get; set; }

        [Required]
        [Display(Name = "Department")]
        public string department { get; set; }

        [Required]
        [Display(Name = "User Status")]
        public string chpass { get; set; }
    }

    public class GroupName
	{
		public int GroupID { get; set; }
		public string GroupDescription { get; set; }
	}

    public class Departments
	{
		public string Dept { get; set; }
		public string Department { get; set; }
	}

    public class ReasonCode
	{
		public int RNID { get; set; }
		public string RCode { get; set; }
		public string Reason { get; set; }
		public string Description { get; set; }
		public string FirstLetter { get; set; }
		public string SecondLetter { get; set; }
		public string ThirdLetter { get; set; }
		public string Abbreviation { get; set; }
		public string Class { get; set; }
	}

	public class NewReason
	{
		[Required]
		[Display(Name = "Reason Code")]
		public string RCode { get; set; }

		[Required]
		[Display(Name = "Reason")]
		public string Reason { get; set; }

		[Required]
		[Display(Name = "Deascription")]
		public string Description { get; set; }

		[Required]
		[Display(Name = "First Letter")]
		public string FirstLetter { get; set; }

		[Required]
		[Display(Name = "Second Letter")]
		public string SecondLetter { get; set; }

		[Required]
		[Display(Name = "Third Letter")]
		public string ThirdLetter { get; set; }

		[Required]
		[Display(Name = "Abbreviation")]
		public string Abbreviation { get; set; }

		[Required]
		[Display(Name = "Classification")]

		public string Class { get; set; }
	}

	public class EditReason
	{
		[Key]
		public int RNID { get; set; }

		[Required]
		[Display(Name = "Reason Code")]
		public string RCode { get; set; }

		[Required]
		[Display(Name = "Reason")]
		public string Reason { get; set; }

		[Required]
		[Display(Name = "Deascription")]
		public string Description { get; set; }

		[Required]
		[Display(Name = "First Letter")]
		public string FirstLetter { get; set; }

		[Required]
		[Display(Name = "Second Letter")]
		public string SecondLetter { get; set; }

		[Required]
		[Display(Name = "Third Letter")]
		public string ThirdLetter { get; set; }

		[Required]
		[Display(Name = "Abbreviation")]
		public string Abbreviation { get; set; }

		[Required]
		[Display(Name = "Classification")]
		public string Class { get; set; }

	}

	public class ATM
	{
		[Key]
		public int ID { get; set; }
		public string Description { get; set; }
		public string DesType { get; set; }
	}

	public class NewATM
	{
		[Required]
		public string Description { get; set; }
	}

	public class EditATM
	{
		[Key]
		public int ID { get; set; }
		[Required]
		public string Description { get; set; }
	}


	public class NewATMLoc
	{
		[Required]
		[Display(Name = "ATM Name")]
		public int ATM_ID { get; set; }

		[Required]
		[Display(Name = "ATM Location")]
		public string ATM_LOCATION { get; set; }

	}

	public class EditATMLoc
	{
		[Key]
		public int ID { get; set; }

		[Required]
		public int ATM_ID { get; set; }

		public int ATM_LOC_ID { get; set; }

		[Required]
		public string ATM_LOCATION { get; set; }

	}

	public class UserMenu
	{
		public int ID { get; set; }
		public string View { get; set; }
		public string Page { get; set; }
		public string Title { get; set; }
	}

	public class NewMenu
	{
		[Required]
		[Display(Name = "View")]
		public string View { get; set; }

		[Required]
		[Display(Name = "Page")]
		public string Page { get; set; }

		[Required]
		[Display(Name = "Title")]
		public string Title { get; set; }

	}

	public class EditMenu
	{
		[Key]
		public int ID { get; set; }
		[Required]
		[Display(Name = "View")]
		public string View { get; set; }

		[Required]
		[Display(Name = "Page")]
		public string Page { get; set; }

		[Required]
		[Display(Name = "Title")]
		public string Title { get; set; }

	}
	public class UserRight
	{
		[Key]
		public int ID { get; set; }
		public string Title { get; set; }
		public int Rights { get; set; }
	}

	public class Audit_Logs
	{
		[Key]
		public int aid { get; set; }
		public DateTime logdate { get; set; }
		public string description { get; set; }
		public string workstn { get; set; }
		public string activity { get; set; }

	}

	public class UserPass
	{
		public string LogindID { get; set; }
		public string Password { get; set; }
	}

}
