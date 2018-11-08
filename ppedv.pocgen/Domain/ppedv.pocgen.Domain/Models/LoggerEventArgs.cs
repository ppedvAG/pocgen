using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace ppedv.pocgen.Domain.Models
{
    public class LoggerEventArgs : EventArgs
    {
        public LoggerEventArgs(string ClassName, string MemberName, string Message)
        {
            this.ClassName = ClassName;
            this.MemberName = MemberName;
            this.Message = Message;
            Time = DateTime.Now;
        }
        public DateTime Time { get;}
        public string ClassName { get;}
        public string MemberName { get;}
        public string Message { get;}
    }
}
