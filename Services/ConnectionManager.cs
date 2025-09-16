using Oracle.ManagedDataAccess.Client;
using System.Data;


namespace AgentDesktop.Services
{
	public static class ConnectionManager
	{
        //private static string connString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(" +
        //    "HOST=10.89.10.88) " +
        //    "(PORT=1524)))(CONNECT_DATA=(SERVER=DEDICATED)" +
        //    "(SERVICE_NAME=COREBANKDR)));" +
        //    "User Id = DEVFINAL01; " +
        //    "Password = passDEVFINAL01; ";

        public static IDbConnection GetConnection(string server)
        {

            string sqlString = "";

            if (server == "I")
			{
                sqlString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(" +
                        "HOST=" + Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppORAICBAConn:HOST"].ToString()) + ")" +
                        "(PORT=" + Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppORAICBAConn:PORT"].ToString()) + ")))(CONNECT_DATA=(SERVER=DEDICATED)" +
                        "(SERVICE_NAME=" + Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppORAICBAConn:SRVNAM"].ToString()) + ")));" +
                        "User Id=" + Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppORAICBAConn:USRID"].ToString()) + "; " +
                        "Password=" + Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppORAICBAConn:PAWD"].ToString()) + "; ";

            }
            else
			{
                sqlString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(" +
                        "HOST=" + Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppORABWConn:HOST"].ToString()) + ")" +
                        "(PORT=" + Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppORABWConn:PORT"].ToString()) + ")))(CONNECT_DATA=(SERVER=DEDICATED)" +
                        "(SERVICE_NAME=" + Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppORABWConn:SRVNAM"].ToString()) + ")));" +
                        "User Id=" + Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppORABWConn:USRID"].ToString()) + "; " +
                        "Password=" + Security.DeCryptor(Connect.ConfigurationManager.AppSetting["AppORABWConn:PAWD"].ToString()) + "; ";

            }


            var conn = new OracleConnection(sqlString);
            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }

            return conn;
        }

        public static void CloseConnection(IDbConnection conn)
        {
            if (conn.State == ConnectionState.Open || conn.State == ConnectionState.Broken)
            {
                conn.Close();
            }
        }
    }
}
