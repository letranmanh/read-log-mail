using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadFileToExcel
{
    class DBUtils
    {
        public static SqlConnection GetDBConnection()
        {
            return DBConnection.GetDBConnection();
        }
    }
}
