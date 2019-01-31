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
using System.Windows.Forms;
using System.IO;

namespace ClearMessageOutlookAddIn
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        ApiHelper apiHelper = new ApiHelper();

        ClearMailModel clearMailModel;
        Personalizations personalizations;
        Attachments attachments;
        List<Attachments> attachmentList = new List<Attachments>();
        List<TempAttachments> tempAttachmentList = new List<TempAttachments>();

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

            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            this.Application.ItemSend += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        public void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
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

                //Calling the event before Attachment is fully added to the mail item
                mailItem.BeforeAttachmentAdd += MailItem_BeforeAttachmentAdd;
                mailItem.AttachmentRemove += MailItem_AttachmentRemove;

                bool itemSend = false;
                if (!string.IsNullOrWhiteSpace(mailItem.EntryID))
                    Application_ItemSend(mailItem, ref itemSend);
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void MailItem_BeforeAttachmentAdd(Outlook.Attachment Attachment, ref bool Cancel)
        {
            string filePath = Attachment.GetTemporaryFilePath();

            byte[] fileInBytes = System.IO.File.ReadAllBytes(filePath);

            string fileContent = Convert.ToBase64String(fileInBytes, 0, fileInBytes.Length);
            Attachment.DisplayName = Attachment.DisplayName + Attachment.Index.ToString();
            attachmentList.Add(new Attachments()
            {
                filename = Attachment.FileName,
                content = fileContent,
                type = Path.GetExtension(filePath),
                disposition = "attachment"
            });

            tempAttachmentList.Add(new TempAttachments()
            {
                displayName = Attachment.DisplayName,
                filename = Attachment.FileName,
                content = fileContent,
                type = Path.GetExtension(filePath),
                disposition = "attachment"
            });
        }

        private void MailItem_AttachmentRemove(Outlook.Attachment Attachment)
        {
            if (attachmentList.Count > 0)
            {
                var attachment = (from file in tempAttachmentList
                                  where file.displayName == Attachment.DisplayName
                                  select file).FirstOrDefault();

                tempAttachmentList.Remove(attachment);

                attachmentList.Clear();

                //attachmentList.Add(new Attachments()
                //{
                //    filename = Attachment.FileName,
                //    content = fileContent,
                //    type = Path.GetExtension(filePath),
                //    disposition = "attachment"
                //});
            }
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            //testSendClearMessageEmailAsync();
            if (Item is Outlook.MailItem)
            {
                /*Initializing the clear message email model objects*/
                InitializeMailObjects();

                Outlook.MailItem mail = (Outlook.MailItem)Item;
                Outlook.Folder contacts = (Outlook.Folder)Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                string toEmailDisplayName = string.Empty;

                #region Logic for TO recipients
                
                List<string> smptEmailAddressList = new List<string>();

                foreach (Outlook.Recipient recipient in mail.Recipients)
                {
                    dynamic address = recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E");

                    smptEmailAddressList.Add(address.ToString().Trim());
                }


                if (mail.To.Contains(';'))
                {
                    //Splitting the TO email addresses for checking whether it exists in the contacts
                    string[] toEmailAddressList = mail.To.Split(';');

                    //Loop over all the "TO:" email addresses entered in the recipients
                    foreach (Outlook.Recipient recipient in mail.Recipients)
                    {
                        dynamic toEmail = recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E");

                        string emailToFind = string.Empty;

                        foreach (string email in toEmailAddressList)
                        {
                            emailToFind = email.Contains("(") ? email.Remove(email.IndexOf("("), (email.Length - email.IndexOf("("))).Trim() : email.Trim();
                        }


                        var foundEmail = Array.FindAll(toEmailAddressList, s => s.Equals(toEmail.ToString().Trim())).First();

                        if (!string.IsNullOrWhiteSpace(foundEmail))
                        {

                        }

                        //Removing the email address from the Display name so that it should always match with the contact
                        toEmailDisplayName = toEmail.Contains("(") ? toEmail.Remove(toEmail.IndexOf("("), (toEmail.Length - toEmail.IndexOf("("))).Trim() : toEmail.Trim();

                        //Loop till the toEmail found in the contacts list - once contact found loop will break and run for another recipient
                        CheckRecipientsInContactsAsync(toEmailDisplayName, contacts, mail);
                    }
                }
                else
                {
                    //Removing the email address from the Display name so that it should always match with the contact
                    toEmailDisplayName = mail.To.Contains("(") ? mail.To.Remove(mail.To.IndexOf("("), (mail.To.Length - mail.To.IndexOf("("))).Trim() : mail.To.Trim();

                    //Loop till the toEmail found in the contacts list - once contact found loop will break and run for another recipient
                    CheckRecipientsInContactsAsync(toEmailDisplayName, contacts, mail);
                }
                #endregion

                #region Logic for CC recipients

                //if (!string.IsNullOrWhiteSpace(mail.CC))
                //{
                //    if (mail.CC.Contains(';'))
                //    {
                //        //Splitting the CC email addresses for checking whether it exists in the contacts
                //        string[] ccEmailAddressList = mail.CC.Split(';');

                //        //Loop over all the "CC:" email addresses entered in the recipients
                //        foreach (var ccEmail in ccEmailAddressList)
                //        {
                //            //Loop till the ccEmail found in the contacts list - once contact found loop will break and run for another recipient
                //            CheckRecipientsInContactsAsync(ccEmail.Trim(), contacts, mail);
                //        }
                //    }
                //    else
                //    {
                //        //Loop till the ccEmail found in the contacts list - once contact found loop will break and run for another recipient
                //        CheckRecipientsInContactsAsync(mail.CC.Trim(), contacts, mail);
                //    }
                //}

                #endregion

                //Call for the Clear Message API method for sending emails
                SendClearMessageEmailAsync();
            }
        }

        private async Task CheckRecipientsInContactsAsync(string emailAddress, Outlook.Folder contacts, Outlook.MailItem mail)
        {
            //Loop till the toEmail found in the contacts list - once contact found loop will break and run for another recipient
            foreach (Outlook.ContactItem contactItem in contacts.Items)
            {
                #region Logic for removing the email address from the contactItem so that it matches parameter emailAddress

                string Email1DisplayName = string.Empty;
                string Email2DisplayName = string.Empty;
                string Email3DisplayName = string.Empty;

                if (!string.IsNullOrEmpty(contactItem.Email1DisplayName))
                    Email1DisplayName = contactItem.Email1DisplayName.Contains("(") ? contactItem.Email1DisplayName.Remove(contactItem.Email1DisplayName.IndexOf("("), (contactItem.Email1DisplayName.Length - contactItem.Email1DisplayName.IndexOf("("))).Trim() : contactItem.Email1DisplayName;

                if (!string.IsNullOrEmpty(contactItem.Email2DisplayName))
                    Email2DisplayName = contactItem.Email2DisplayName.Contains("(") ? contactItem.Email2DisplayName.Remove(contactItem.Email2DisplayName.IndexOf("("), (contactItem.Email2DisplayName.Length - contactItem.Email2DisplayName.IndexOf("("))).Trim() : contactItem.Email2DisplayName;

                if (!string.IsNullOrEmpty(contactItem.Email3DisplayName))
                    Email3DisplayName = contactItem.Email3DisplayName.Contains("(") ? contactItem.Email3DisplayName.Remove(contactItem.Email3DisplayName.IndexOf("("), (contactItem.Email3DisplayName.Length - contactItem.Email3DisplayName.IndexOf("("))).Trim() : contactItem.Email3DisplayName;

                #endregion

                //Check the toEmail matches the contact's email and display name
                if (emailAddress == Email1DisplayName || emailAddress == Email2DisplayName || emailAddress == Email3DisplayName)
                {
                    //getting the email address from the contact instead of displayName
                    string contactEmailAddress = string.Empty;
                    //if (contactItem.Email1AddressType)
                    contactEmailAddress = !string.IsNullOrWhiteSpace(contactItem.Email1Address.Trim()) ? contactItem.Email1Address.Trim() : !string.IsNullOrWhiteSpace(contactItem.Email2Address.Trim()) ? contactItem.Email2Address.Trim() : !string.IsNullOrWhiteSpace(contactItem.Email3Address.Trim()) ? contactItem.Email3Address.Trim() : contactItem.Email3Address.Trim();

                    //Check the user defined property "SendViaClearMessage" exists for the contact
                    var CustomProperty = contactItem.UserProperties.Find("SendViaClearMessage", true);
                    if (CustomProperty != null)
                    {
                        if (contactItem.UserProperties["SendViaClearMessage"].Value)
                        {
                            //The call for the adding the clear message object once found checked
                            PerpareClearMessageModel(contactEmailAddress, mail);
                        }
                    }
                    break;
                }
            }
        }

        private void PerpareClearMessageModel(string emailAddress, Outlook.MailItem mail)
        {
            //Setting the personalization class object
            personalizations.to.Add(new To() { email = emailAddress });
            personalizations.subject = mail.Subject;

            //Adding the personalization object to the list of Persaonalizations
            clearMailModel.personalizations.Add(personalizations);

            //Adding the list of Content class object 
            clearMailModel.content.Add(new Content() { type = "text/plain", value = mail.Body });

            //Assigning the list of Attachents to the ClearMailModel.Attachment object
            if (!clearMailModel.attachments.Any())
                clearMailModel.attachments = attachmentList;

            // The propetry set for the FROM email address
            if (string.IsNullOrWhiteSpace(clearMailModel.from.email))
                clearMailModel.from.email = mail.SendUsingAccount.SmtpAddress;
        }

        private async Task SendClearMessageEmailAsync()
        {
            //string mailJson = "{  \"personalizations\": [    {      \"to\": [        {          \"email\": \"gulrezansari@virtualemployee.com\"        }      ],      \"subject\": \"Sending with outlook addin is Fun\"    }  ],  \"from\": {    \"email\": \"gulrezansari@virtualemployee.com\"  },  \"content\": [    {      \"type\": \"text/plain\",      \"value\": \"This email has been sent via outlook plugin and clear message api\"    }  ],\"attachments\":[{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"},{\"content\": \"filedata\",\"type\": \"jpeg\", \"filename\": \"testfile\", \"disposition\": \"attachment\"}]}";

            //ClearMailModel mailObject = JsonConvert.DeserializeObject<ClearMailModel>(mailJson);

            string mailObjectJson = JsonConvert.SerializeObject(clearMailModel);

            try
            {
                HttpClient client = apiHelper.InitializeClient();
                using (var content = new StringContent(mailObjectJson, System.Text.Encoding.Default, "application/json"))
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

        private void InitializeMailObjects()
        {
            clearMailModel = new ClearMailModel();
            clearMailModel.personalizations = new List<Personalizations>();
            clearMailModel.from = new From();
            clearMailModel.content = new List<Content>();
            clearMailModel.attachments = new List<Attachments>();

            personalizations = new Personalizations();
            personalizations.to = new List<To>();
            attachments = new Attachments();
        }

        private string GetSenderEmailAddress(Outlook.MailItem mail)
        {
            Outlook.AddressEntry sender = mail.Sender;
            string SenderEmailAddress = "";

            if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry || sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
            {
                Outlook.ExchangeUser exchUser = sender.GetExchangeUser();
                if (exchUser != null)
                {
                    SenderEmailAddress = exchUser.PrimarySmtpAddress;
                }
            }
            else
            {
                SenderEmailAddress = mail.SenderEmailAddress;
            }

            return SenderEmailAddress;
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
