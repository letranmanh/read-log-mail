using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadFileToExcel
{
    class DBConnection
    {
        public static SqlConnection GetDBConnection()
        {
            //
            // Data Source=TRAN-VMWARE\SQLEXPRESS;Initial Catalog=simplehr;Persist Security Info=True;User ID=sa;Password=12345
            //
            string connString = ConfigurationManager.ConnectionStrings["strATTGroupConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connString);

            return conn;
        }

    }
}
