using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ClearMessageOutlookAddIn
{
    [RunInstaller(true)]
    public partial class SetupHelper : System.Configuration.Install.Installer
    {
        public SetupHelper()
        {
            InitializeComponent();
        }

        public override void Install(System.Collections.IDictionary stateSaver)
        {
            try
            {
                base.Install(stateSaver);

                //Saving the CustomActionData parameters to the install state dictionary to access afterwards in Install Commit
                stateSaver.Add("TargetDir", Context.Parameters["targetdir"].ToString());
                stateSaver.Add("BearerKey",Context.Parameters["bearerKey"].ToString());
                stateSaver.Add("ApiBaseUrl", Context.Parameters["apiBaseUrl"].ToString());
            }
            catch (Exception e)
            {
                string s = e.Message;
            }
        }

        public override void Commit(IDictionary savedState)
        {
            try
            {
                base.Commit(savedState);

                //Getting the location of the directory where the plugin files will get installed
                string filePath = Path.GetDirectoryName(Context.Parameters["AssemblyPath"]);

                //MessageBox.Show(filePath);

                //Getting the settings.json file to update with bearer and endpoints
                string jsonSettingsPath = filePath + "\\settings.json";

                //MessageBox.Show(jsonSettingsPath);

                SettingsModel settingsModel = new SettingsModel();

                //Reading the settings.json file from the stream
                string json = string.Empty;
                using (StreamReader sr = new StreamReader(jsonSettingsPath))
                {
                    json = sr.ReadToEnd();
                }

                //MessageBox.Show(json);

                //Deserialzed the settings.json file to the SettingsModel object 
                settingsModel = JsonConvert.DeserializeObject<SettingsModel>(json);

                //MessageBox.Show(settings.ToString());
                //MessageBox.Show(Context.Parameters["bearerKey"]);
                //MessageBox.Show(Context.Parameters["apiBaseUrl"]);
                //MessageBox.Show(Context.Parameters["targetdir"]);

                //If settingsModel is not null then we will update the bearer token and endpoints
                if (settingsModel != null)
                {
                    settingsModel.BearerKey = Context.Parameters["bearerKey"];
                    settingsModel.ApiBaseUrl = Context.Parameters["apiBaseUrl"];
                }

                //MessageBox.Show(settingsModel.ToString());

                //Finally write and replace all the text in the settings.json file.
                File.WriteAllText(jsonSettingsPath, JsonConvert.SerializeObject(settingsModel));
                //MessageBox.Show("Done: " + jsonSettingsPath);
            }
            catch (Exception e)
            {
                string s = e.Message;
            }
        }
    }
}
