using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.Net.Http;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Reflection;
using System.Runtime.Versioning;
using System.Net;

namespace ClearMessageOutlookAddIn
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        ApiHelper apiHelper = new ApiHelper();
        Microsoft.Office.Interop.Outlook.FormRegion _formRegion;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var attributes = assembly.GetCustomAttributes(typeof(TargetFrameworkAttribute), false);
            var version = (TargetFrameworkAttribute)attributes[0];

            SecurityProtocolType flag;
            if (Enum.TryParse("Tls11", out flag))
                ServicePointManager.SecurityProtocol |= flag;
            if (Enum.TryParse("Tls12", out flag))
                ServicePointManager.SecurityProtocol |= flag;

            //_formRegion = new Microsoft.Office.Interop.Outlook.FormRegion();
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            this.Application.ItemSend += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        public void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            //Outlook.Folder folder = Application.ActiveExplorer().CurrentFolder as Outlook.Folder;
            //Outlook.ContactItem contactItem = folder.Items.Add("IPM.Contact") as Outlook.ContactItem;

            //// Add UserProperty to contactItem
            //contactItem.UserProperties.Add("SendViaClearMessage", Outlook.OlUserPropertyType.olYesNo, true, Type.Missing);
            //contactItem.UserProperties["SendViaClearMessage"].Value = true;
            //contactItem.Subject = "UserProperty Example";
            //contactItem.Save();


            try
            {
                Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
                if (mailItem != null)
                {
                    if (mailItem.EntryID == null)
                    {
                        mailItem.Subject = "This text was added by using code";
                        mailItem.Body = "This text was added by using code";
                    }
                }

                bool itemSend = false;
                if (!string.IsNullOrWhiteSpace(mailItem.EntryID))
                    Application_ItemSend(mailItem, ref itemSend);
            }
            catch (Exception e)
            {

            }
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            //testSendClearMessageEmailAsync();
            if (Item is Outlook.MailItem)
            {
                /*temp variable just to check email exists in contact - will remove it later*/
                bool hasEmail = false;

                Outlook.MailItem mail = (Outlook.MailItem)Item;
                Outlook.Folder contacts = (Outlook.Folder)Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);

                #region Logic for TO recipents


                if (mail.To.Contains(';'))
                {
                    //Splitting the TO email addresses for checking whether it exists in the contacts
                    string[] toEmailAddressList = mail.To.Split(';');

                    //Loop over all the "TO:" email addresses entered in the recipients
                    foreach (var toEmail in toEmailAddressList)
                    {
                        //Loop till the toEmail found in the contacts list - once contact found loop will break and run for another recipient
                        CheckRecipientsInContactsAsync(toEmail, contacts, mail);
                    }
                }
                else
                {
                    //Loop till the toEmail found in the contacts list - once contact found loop will break and run for another recipient
                    CheckRecipientsInContactsAsync(mail.To, contacts, mail);
                }
                #endregion

                #region Logic for CC recipents

                if (!string.IsNullOrWhiteSpace(mail.CC))
                {
                    if (mail.CC.Contains(';'))
                    {
                        //Splitting the CC email addresses for checking whether it exists in the contacts
                        string[] ccEmailAddressList = mail.CC.Split(';');

                        //Loop over all the "CC:" email addresses entered in the recipients
                        foreach (var ccEmail in ccEmailAddressList)
                        {
                            //Loop till the ccEmail found in the contacts list - once contact found loop will break and run for another recipient
                            CheckRecipientsInContactsAsync(ccEmail, contacts, mail);
                        }
                    }
                    else
                    {
                        //Loop till the ccEmail found in the contacts list - once contact found loop will break and run for another recipient
                        CheckRecipientsInContactsAsync(mail.CC, contacts, mail);
                    }
                }

                #endregion

                mail.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
                //mail.Send();
            }
        }

        private async Task CheckRecipientsInContactsAsync(string emailAddress, Outlook.Folder contacts, Outlook.MailItem mail)
        {
            //Loop till the toEmail found in the contacts list - once contact found loop will break and run for another recipient
            foreach (Outlook.ContactItem contactItem in contacts.Items)
            {
                //Check the toEmail matches the contact's email and display name
                if (emailAddress == contactItem.Email1DisplayName || emailAddress == contactItem.Email2DisplayName || emailAddress == contactItem.Email3DisplayName)
                {
                    //Check the user defined property "SendViaClearMessage" exists for the contact
                    var CustomProperty = contactItem.UserProperties.Find("SendViaClearMessage", true);
                    if (CustomProperty != null)
                    {
                        if (contactItem.UserProperties["SendViaClearMessage"].Value)
                        {
                            //Logic for calling the Clear Message API method for sending emails
                            await SendClearMessageEmailAsync(mail, contactItem);
                        }
                    }
                    break;
                }
            }
        }

        private async Task SendClearMessageEmailAsync(Outlook.MailItem mail, Outlook.ContactItem contact)
        {
            //Logic for sending the emails with attachments count.
            int attachmentCount = mail.Attachments.Count;
            decimal splitEmails;

            decimal.TryParse(Convert.ToString(attachmentCount / 10), out splitEmails);

            if (attachmentCount > 0 && attachmentCount <= 10)
            {

            }

            string mailJson = "{  \"personalizations\": [    {      \"to\": [        {          \"email\": \"rspandey@virtualemployee.com\"        }      ],      \"subject\": \"Sending with outlook addin is Fun\"    }  ],  \"from\": {    \"email\": \"gulrezansari@virtualemployee.com\"  },  \"content\": [    {      \"type\": \"text/plain\",      \"value\": \"This email has been sent via outlook plugin and clear message api\"    }  ],\"attachments\":[{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"}]}";

            ClearMailModel mailObject = JsonConvert.DeserializeObject<ClearMailModel>(mailJson);

            string mailObjectJson = JsonConvert.SerializeObject(mailJson);

            try
            {
                HttpClient client = apiHelper.InitializeClient();
                using (var content = new StringContent(mailJson, System.Text.Encoding.Default, "application/json"))
                {
                    using (HttpResponseMessage response = await client.PostAsync("v1/mail/send", content))
                    {
                        string responseData = await response.Content.ReadAsStringAsync();
                    }
                }
            }
            catch (Exception ex)
            {

            }

        }


        //private async Task SendClearMessageEmailAsync()
        //{
        //    string mailJson = "{  \"personalizations\": [    {      \"to\": [        {          \"email\": \"rspandey@virtualemployee.com\"        }      ],      \"subject\": \"Sending with outlook addin is Fun\"    }  ],  \"from\": {    \"email\": \"gulrezansari@virtualemployee.com\"  },  \"content\": [    {      \"type\": \"text/plain\",      \"value\": \"This email has been sent via outlook plugin and clear message api\"    }  ],\"attachments\":[{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"}]}";

        //    ClearMailModel mailObject = JsonConvert.DeserializeObject<ClearMailModel>(mailJson);

        //    string mailObjectJson = JsonConvert.SerializeObject(mailJson);

        //    try
        //    {
        //        HttpClient client = apiHelper.InitializeClient();
        //        using (var content = new StringContent(mailJson, System.Text.Encoding.Default, "application/json"))
        //        {
        //            using (HttpResponseMessage response = await client.PostAsync("v1/mail/send", content))
        //            {
        //                string responseData = await response.Content.ReadAsStringAsync();
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //    }

        //}

        //private void EnumerateAddressLists()
        //{
        //    Outlook.AddressLists addrLists = Application.Session.AddressLists;
        //    foreach (Outlook.AddressList addrList in addrLists)
        //    {
        //        StringBuilder sb = new StringBuilder();
        //        sb.AppendLine("Display Name: " + addrList.Name);
        //        sb.AppendLine("Resolution Order: "
        //            + addrList.ResolutionOrder.ToString());
        //        sb.AppendLine("Read-only : "
        //            + addrList.IsReadOnly.ToString());
        //        sb.AppendLine("Initial Address List: "
        //            + addrList.IsInitialAddressList.ToString());
        //        sb.AppendLine("");
        //    }
        //}

        //private void EnumerateGAL()
        //{
        //    Outlook.AddressList gal =
        //        Application.Session.GetGlobalAddressList();
        //    if (gal != null)
        //    {
        //        for (int i = 1;
        //            i <= gal.AddressEntries.Count - 1; i++)
        //        {
        //            Outlook.AddressEntry addrEntry =
        //                gal.AddressEntries[i];
        //            if (addrEntry.AddressEntryUserType ==
        //                Outlook.OlAddressEntryUserType.
        //                olExchangeUserAddressEntry
        //                || addrEntry.AddressEntryUserType ==
        //                Outlook.OlAddressEntryUserType.
        //                olExchangeRemoteUserAddressEntry)
        //            {
        //                Outlook.ExchangeUser exchUser =
        //                    addrEntry.GetExchangeUser();
        //                Debug.WriteLine(exchUser.Name + " "
        //                    + exchUser.PrimarySmtpAddress);
        //            }
        //            if (addrEntry.AddressEntryUserType ==
        //                Outlook.OlAddressEntryUserType.
        //                olExchangeDistributionListAddressEntry)
        //            {
        //                Outlook.ExchangeDistributionList exchDL =
        //                    addrEntry.GetExchangeDistributionList();
        //                Debug.WriteLine(exchDL.Name + " "
        //                    + exchDL.PrimarySmtpAddress);
        //            }
        //        }
        //    }
        //}

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
