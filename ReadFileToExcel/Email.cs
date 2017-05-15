using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadFileToExcel
{
    class Email
    {
        private int _id = -1;
        private string _emailAddr = string.Empty;
        private string _sendDate = string.Empty;
        private string _sendTime = string.Empty;
        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }
        public string EmailAddr
        {
            get { return _emailAddr; }
            set { _emailAddr = value; }
        }
       
        public string SendDate
        {
            get { return _sendDate; }
            set { _sendDate = value; }
        }

        public string SendTime
        {
            get { return _sendTime; }
            set { _sendTime = value; }
        }
    }
}
