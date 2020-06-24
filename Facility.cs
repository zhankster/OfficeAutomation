using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAutomation
{
    public class Facility
    {
        public Facility(string code, string name, string phone, string fax, string email, string notify_type, bool valid_code)
        {
            this.code = code;
            this.name = name;
            this.phone = phone;
            this.fax = fax;
            this.email = email;
            this.notify_type = notify_type;
            this.valid_code = valid_code;
        }

        public string code { get; set; }
        public string name { get; set; }
        public string phone { get; set; }
        public string fax { get; set; }
        public string email { get; set; }
        public string notify_type { get; set; }
        public bool valid_code { get; set; }
    }
}
