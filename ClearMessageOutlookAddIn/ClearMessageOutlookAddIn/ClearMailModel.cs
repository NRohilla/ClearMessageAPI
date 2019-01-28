using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClearMessageOutlookAddIn
{
    public class ClearMailModel
    {
        public List<Personalizations> Personalizations { get; set; }
        public From From { get; set; }
        public List<Content> Content { get; set; }
        public List<Attachments> Attachments { get; set; }
    }

    public class Personalizations
    {
        public List<To> To { get; set; }
        public string Subject { get; set; }
    }

    public class To
    {
        public string Email { get; set; }
    }

    public class From
    {
        public string Email { get; set; }
    }

    public class Content
    {
        public string Type { get; set; }
        public string Value { get; set; }
    }

    public class Attachments
    {
        public string Content { get; set; }
        public string Type { get; set; }
        public string Filename { get; set; }
        public string Disposition { get; set; }
    }
}
