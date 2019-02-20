using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ClearMessageOutlookAddIn
{
    public class SettingsModel
    {
        public string BearerKey { get; set; }
        public string ApiBaseUrl { get; set; }
        public string FilePath
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }
        public string AuditSetting { get; set; }
    }
}
