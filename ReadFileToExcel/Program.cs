using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
namespace ReadFileToExcel
{
    class Program
    {
        
        static void Main(string[] args)
        {
            string[] filepaths = Directory.GetFiles(ConfigurationManager.AppSettings["AppLogPath"], "*.log",
                                         SearchOption.AllDirectories);
            string line;
            int counter = 0;
            int koCheckDublication = 0;
            // create a hashmap with three key/value pairs.
            Dictionary<string, Email> ht = new Dictionary<string, Email>();
            string _Key = String.Empty;
            string _Email = String.Empty;
            string _Date = String.Empty;
            string _Time = String.Empty;
            Email emailObj;
            int lastindex = 0;
            foreach (string filepath in filepaths)
            {

                System.IO.StreamReader file =
                    new System.IO.StreamReader(filepath);
                while ((line = file.ReadLine()) != null)
                {
                    if (line.Contains("Date        :"))
                    {
                        lastindex = line.LastIndexOf(":") + 1;
                        string aDt = line.Substring(lastindex).Trim();
                        if (aDt.Length == 9)
                            _Date = "0" + aDt;
                        else
                            _Date = aDt;
                    }
                    else if (line.Contains("Time        :"))
                    {
                        lastindex = line.IndexOf(":") + 1;
                        string at = (line.Substring(lastindex).Trim()).Substring(0, 8).Trim();
                        if (at.Length == 7)
                            _Time = "0" + at;
                        else
                            _Time = at;
                    }
                    else if (line.Contains("Message     :") && line.Contains("@"))
                    {
                        lastindex = line.LastIndexOf(":") + 1;
                        _Email = line.Substring(lastindex).Trim();
                        _Key = _Date + _Email;
                        koCheckDublication++;
                        if (!ht.ContainsKey(_Key))
                        {
                            emailObj = new Email();
                            emailObj.SendDate = _Date;
                            emailObj.SendTime = _Time;
                            emailObj.EmailAddr = _Email;
                            ht.Add(_Key, emailObj);
                            counter++;
                        }
                    }
                }
                file.Close();
            }

            // lay tap hop cac key
            ICollection<Email> lstEmail = ht.Values;
            string isInsertEmail = System.Configuration.ConfigurationSettings.AppSettings["IsInsert"];
            if ("yes".Equals(isInsertEmail))
                //writeEmailAllLog(lstEmail);
                insertEmailToDB(lstEmail);
            else if (lstEmail.Count > 0)
                writeEmailAllLog(lstEmail);
        }
        
        private static void insertEmailToDB(ICollection<Email> lstEmail)
        {
            //Mo ket noi va insert DB
            SqlConnection conn = DBUtils.GetDBConnection();
            //SqlDataReader rdr = null;
            List<string> lstEmailError = new List<string>();
            List<string> lstEmailOK = new List<string>();
            try
            {
                //conn.Open();
                int i = 0;
                foreach (Email item in lstEmail)
                {
                    if (i == 0)
                    {
                        conn.Open();
                    }
                    //conn.Open();
                    // 1.  create a command object identifying
                    //     the stored procedure
                    SqlCommand cmd = new SqlCommand("sp_NSW_InsertLogToNotifyEmailOpened", conn);
                    // 2. set the command object so it knows
                    //    to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;
                    string tDate = item.SendDate + " " + item.SendTime;
                    DateTime sendDateTime = DateTime.ParseExact(tDate, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                    cmd.Parameters.Add(new SqlParameter("@SendDate", sendDateTime));
                    cmd.Parameters.Add(new SqlParameter("@EmailAddr", item.EmailAddr));
                    cmd.Parameters.Add("@OutputEmail", SqlDbType.VarChar, 255);
                    cmd.Parameters["@OutputEmail"].Direction = ParameterDirection.Output;
                    // execute the command
                    cmd.ExecuteReader();
                    if ("ok".Equals(cmd.Parameters["@OutputEmail"].Value.ToString()))
                        lstEmailOK.Add(item.EmailAddr);
                    else
                        lstEmailError.Add(cmd.Parameters["@OutputEmail"].Value.ToString());
                    i++;
                    if (i > 1000)
                    {
                        conn.Close();
                        //reset
                        i = 0;
                    }
                }
                //Console.WriteLine("So luong Email doc duoc tu file log: " + lstEmail.Count);
            }
            catch (Exception e)
            {
                conn.Close();
                if (lstEmail.Count > 0)
                    writeEmailAllLog(lstEmail);
                if (lstEmailOK.Count > 0)
                    writeEmailOKLog(lstEmailOK);
                Console.WriteLine("Error: " + e.Message);

            }
            finally
            {
                conn.Close();
                if (lstEmailError.Count > 0)
                    writeEmailErrorLog(lstEmailError);
            }
        }
        public static void writeEmailErrorLog(List<string> lstEmailError)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("EmailNotUserId.log");
            foreach(string email in lstEmailError)
            {
                file.WriteLine(email);
            }
            file.Close();
        }

        public static void writeEmailOKLog(List<string> lstEmailOk)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("EmailOk.log");
            foreach (string email in lstEmailOk)
            {
                file.WriteLine(email);
            }
            file.Close();
        }

        public static void writeEmailAllLog(ICollection<Email> lstEmail)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("EmailAll.log");
            foreach (Email _email in lstEmail)
            {
                file.WriteLine(_email.EmailAddr);
                file.WriteLine(_email.SendDate);
            }
            file.Close();
        }
    }
}
