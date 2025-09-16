using EncryptionClassLibrary.Encryption;
using System.Text;

namespace AgentDesktop.Services
{
	public class Security
	{
		public static string DeCryptor(string val)
		{
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

			string deVal = val;

			secureString ss = new secureString();

			ss.Value = "a";
			ss.Encrypt();
			ss.Value = deVal;

			return ss.Value;
		}

		public static string Encryptor(string val)
		{
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

			string deVal = val.Trim();

			secureString ss = new secureString
			{
				Value = deVal
			};
			ss.Encrypt();
			deVal = ss.Value;

			return deVal;
		}
	}
}
