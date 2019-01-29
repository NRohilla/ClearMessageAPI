using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClearMessageOutlookAddIn
{
    public class ClearMailModel
    {
        public List<Personalizations> personalizations { get; set; }
        public From from { get; set; }
        public List<Content> content { get; set; }
        public List<Attachments> attachments { get; set; }
    }

    public class Personalizations
    {
        public List<To> to { get; set; }
        public string subject { get; set; }
    }

    public class To
    {
        public string email { get; set; }
    }

    public class From
    {
        public string email { get; set; }
    }

    public class Content
    {
        public string type { get; set; }
        public string value { get; set; }
    }

    public class Attachments
    {
        public string content { get; set; }
        public string type { get; set; }
        public string filename { get; set; }
        public string disposition { get; set; }
    }
}
