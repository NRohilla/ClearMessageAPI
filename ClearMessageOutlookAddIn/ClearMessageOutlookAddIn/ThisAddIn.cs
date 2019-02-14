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
using System.Runtime.InteropServices;
using System.Resources;

namespace ClearMessageOutlookAddIn
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        ApiHelper apiHelper = new ApiHelper();

        Outlook.MailItem olMailItem = null;
        Outlook.Recipients olRecipients = null;
        Outlook.Recipient olRecipientTO = null;
        Outlook.Recipient cmRecipientTO = null;
        Outlook.Recipient olRecipientCC = null;
        Outlook.Recipient olRecipientBCC = null;
        Outlook.Attachments olAttachments = null;

        ClearMailModel clearMailModel;
        Personalizations personalizations;
        Attachments attachments;
        List<Attachments> attachmentList = new List<Attachments>();
        List<TempAttachments> tempAttachmentList = new List<TempAttachments>();
        List<string> cmRecipientsList = new List<string>();
        string sentEmailEntryID = string.Empty;

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
            inspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            this.Application.ItemSend += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);

            //Initializing the outlook email object
            olMailItem = Application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

            //((Outlook.ItemEvents_10_Event)olMailItem).Send += ThisAddIn_Send; ;

            //IResourceWriter resourceWriter =new ResourceWriter("Resources.resources");
            //ClearMessageOutlookAddIn.Properties.Resources.ResourceManager.GetString("ResxKey", System.Globalization.CultureInfo.GetCultureInfo("en-US"));

        }

        //private void ThisAddIn_Send(ref bool Cancel)
        //{
        //    if (olMailItem.Sent)
        //    {

        //    }
        //}

        public void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            try
            {
                Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
                //if (mailItem != null)
                //{
                //    if (mailItem.EntryID == null)
                //    {
                //        mailItem.Subject = DateTime.Now.ToString();
                //        mailItem.Body = "This mail was sent using ClearMessage API";
                //    }
                //}

                //Initializing the olRecipients which is not marked as Clear Message recipients
                olRecipients = olMailItem.Recipients;

                //Initializing the outlook email attachments object
                olAttachments = olMailItem.Attachments;

                //Calling the event before Attachment is fully added to the mail item
                mailItem.BeforeAttachmentAdd += MailItem_BeforeAttachmentAdd;

                //Calling the event when attachment is remnoved
                mailItem.AttachmentRemove += MailItem_AttachmentRemove;

                bool CancelSend = false;
                if (!string.IsNullOrWhiteSpace(mailItem.EntryID))
                {
                    Application_ItemSend(mailItem, ref CancelSend);
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            if (Item is Outlook.MailItem)
            {
                Outlook.MailItem mail = (Outlook.MailItem)Item;

                //Send the ClearMessage api email message which are marked with ClearMessage checked
                CreateClearMessageMailItem(mail);

                //Send the outlook normal email message which are not marked with ClearMessage
                SendOutlookEmails(mail);

                //Canceling the ItemSend event of outlook and overiding by our own method 
                Cancel = true;

                //Releasing the mail object from the memory
                if (mail != null)
                    Marshal.FinalReleaseComObject(mail);

                //Closing the new mail window and discarding the composed message.
                this.Application.ActiveInspector().Close(Outlook.OlInspectorClose.olDiscard);

                /*===============================================================================*/
                //Outlook.Items sentEmailItems = null;
                //Outlook.Items sentEmailItems = (Outlook.Items)Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Items;
                //Outlook.Folder sentFolder = (Outlook.Folder)Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);

                //Outlook.NameSpace ns = Application.GetNamespace("MAPI");

                //Outlook.MailItem objMail = ns.GetItemFromID(sentEmailEntryID, sentFolder.StoreID);

                ////sentEmailItems = sentEmailsList.Items;

                //sentEmailItems.Sort("[SentOn]", true);

                //dynamic obj = sentEmailItems.Find(sentEmailEntryID);

                //foreach (dynamic emailItem in sentEmailItems)
                //{
                //    if (olMailItem.EntryID != null)
                //    {

                //    }
                //}

                //if (cmRecipientsList.Any())
                //{
                //    foreach (string cmEmail in cmRecipientsList)
                //    {
                //        cmRecipientTO = olRecipients.Add(cmEmail);
                //        cmRecipientTO.Type = 1;
                //    }
                //}
                /*===============================================================================*/

                Dispose();
                ResetObjects();
            }
        }

        private void CreateClearMessageMailItem(Outlook.MailItem mailObject)
        {
            if (mailObject is Outlook.MailItem)
            {
                /*Initializing the clear message email model objects*/
                InitializeMailObjects();

                Outlook.MailItem mail = (Outlook.MailItem)mailObject;
                Outlook.Folder contacts = (Outlook.Folder)Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);

                #region Logic for TO recipients to send ClearMessage API

                //Loop over all the "TO:" email addresses entered in the recipients
                foreach (Outlook.Recipient recipient in mail.Recipients)
                {
                    //Converting the exchange email address to SMTP email address
                    dynamic toEmail = recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E");

                    //Checking if the recipient is in To address bar
                    if (recipient.Type == 1)
                    {
                        //Loop till the toEmail found in the contacts list - once contact found loop will break and run for another recipient
                        CheckRecipientsInContactsAsync(toEmail.ToString().Trim(), contacts, mail);
                    }
                    else
                    {
                        if (recipient.Type == 2) // For CC
                        {
                            //Adding the CC recipient to the outlook  olRecipientCC
                            olRecipientCC = olRecipients.Add(toEmail.ToString().Trim());
                            olRecipientCC.Type = 2;
                        }
                        else //(recipient.Type == 3) //For BCC
                        {
                            //Adding the BCC recipient to the outlook  olRecipientBCC
                            olRecipientBCC = olRecipients.Add(toEmail.ToString().Trim());
                            olRecipientBCC.Type = 3;
                        }
                    }
                }

                //Call for the Clear Message API method for sending emails
                SendClearMessageEmailAsync();

                #endregion
            }

        }

        private void SendOutlookEmails(Outlook.MailItem mailItem)
        {
            if (olRecipients.Count > 0)
            {
                olMailItem.Subject = mailItem.Subject;
                olMailItem.Body = mailItem.Body;
                olMailItem.BodyFormat = mailItem.BodyFormat;

                olMailItem.Save();

                sentEmailEntryID = olMailItem.EntryID;

                olMailItem.Send();
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

                //Checking the Email address is of type "SMTP" or "EX"
                if (!string.IsNullOrEmpty(contactItem.Email1Address))
                    Email1DisplayName = contactItem.Email1AddressType == "SMTP" ? contactItem.Email1Address.Trim() : GetSmtpEmaillAddress(contactItem);

                if (!string.IsNullOrEmpty(contactItem.Email2Address))
                    Email2DisplayName = contactItem.Email1AddressType == "SMTP" ? contactItem.Email1Address.Trim() : GetSmtpEmaillAddress(contactItem);

                if (!string.IsNullOrEmpty(contactItem.Email3Address))
                    Email3DisplayName = contactItem.Email1AddressType == "SMTP" ? contactItem.Email1Address.Trim() : GetSmtpEmaillAddress(contactItem);

                #endregion

                //Check the toEmail matches the contact's email and display name
                if (emailAddress == Email1DisplayName || emailAddress == Email2DisplayName || emailAddress == Email3DisplayName)
                {
                    //getting the email address from the contact instead of displayName
                    string contactEmailAddress = string.Empty;

                    contactEmailAddress = !string.IsNullOrWhiteSpace(Email1DisplayName.Trim()) ? Email1DisplayName.Trim() : !string.IsNullOrWhiteSpace(Email2DisplayName.Trim()) ? Email2DisplayName.Trim() : !string.IsNullOrWhiteSpace(Email3DisplayName.Trim()) ? Email3DisplayName.Trim() : Email3DisplayName.Trim();

                    //Check the user defined property "SendViaClearMessage" exists for the contact
                    var CustomProperty = contactItem.UserProperties.Find("SendViaClearMessage", true);
                    if (CustomProperty != null)
                    {
                        if (contactItem.UserProperties["SendViaClearMessage"].Value)
                        {
                            //Adding the Clear Message recipients to the list for saving the email in sent items with all emails
                            cmRecipientsList.Add(emailAddress);

                            //The call for the adding the clear message object once found true for the property
                            PerpareClearMessageModel(contactEmailAddress, mail);
                        }
                        else
                        {
                            //Else add the recipient to the normal outlook recipient for sending email (w/o clear message)
                            olRecipientTO = olRecipients.Add(emailAddress);
                            olRecipientTO.Type = 1;
                        }
                    }
                    break;
                }
            }
        }

        private string GetSmtpEmaillAddress(Outlook.ContactItem contactItem)
        {
            dynamic contactProp = contactItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/8084001F");

            return contactProp.ToString().Trim();
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

        private void MailItem_BeforeAttachmentAdd(Outlook.Attachment Attachment, ref bool Cancel)
        {
            //Getting attachment path
            string filePath = Attachment.GetTemporaryFilePath();

            //Convert to Byte arraty to read file data
            byte[] fileInBytes = System.IO.File.ReadAllBytes(filePath);

            //Getting the type (extension) of the attachment file
            string fileExtension = Path.GetExtension(filePath);

            //File data in Base64 String
            string fileContent = Convert.ToBase64String(fileInBytes, 0, fileInBytes.Length);

            //Getting the Attachment's file name without extension
            string fileDisplayName = Path.GetFileNameWithoutExtension(filePath);

            //Making the display name unqiue  - comes handy when deleting the attachments
            Attachment.DisplayName = fileDisplayName + Attachment.Index.ToString() + fileExtension;

            //Adding the files to the attachment list
            attachmentList.Add(new Attachments()
            {
                filename = Attachment.FileName,
                content = fileContent,
                type = Path.GetExtension(filePath),
                disposition = "attachment"
            });

            //Adding the attachments to the Outlook mailItems
            olAttachments.Add(filePath, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, 1, Attachment.DisplayName);

            //Replica of Attachments object used when attachment is removed from the list
            tempAttachmentList.Add(new TempAttachments()
            {
                displayName = Attachment.DisplayName,
                filename = Attachment.FileName,
                content = fileContent,
                type = fileExtension,
                disposition = "attachment"
            });
        }

        private void MailItem_AttachmentRemove(Outlook.Attachment Attachment)
        {
            //Finding the attachment if it exist in the list
            var attachment = (from file in tempAttachmentList
                              where file.displayName == Attachment.DisplayName
                              select file).FirstOrDefault();

            //Removing the attachmnet from the Temp list
            tempAttachmentList.Remove(attachment);

            string fileDisplayName = Path.GetFileNameWithoutExtension(Attachment.DisplayName);

            //Removed an attachment from the outlook mail items which send separate email
            olAttachments.Remove(Convert.ToInt16(fileDisplayName.Remove(0, fileDisplayName.Length - 1)));

            //Initializing the new mail inspector to get the mail Item object
            Outlook.Inspector mailInspector = Globals.ThisAddIn.Application.ActiveInspector();

            //If the mail object not Null then rename the display names of the attachments (For Global MailItem)
            if (mailInspector != null)
            {
                Outlook.MailItem item = (Outlook.MailItem)mailInspector.CurrentItem;
                if (item.Attachments.Count > 0)
                {
                    int index = 0;
                    string objAttachmentDisplayName;
                    string objAttachmentExtension;

                    //Traversing the all the Global mailItem Attachments to remove the indexing and then append the fileExtension to the display names
                    foreach (Outlook.Attachment objAttachment in item.Attachments)
                    {
                        objAttachmentDisplayName = Path.GetFileNameWithoutExtension(objAttachment.DisplayName);
                        objAttachmentExtension = Path.GetExtension(objAttachment.DisplayName);

                        objAttachment.DisplayName = objAttachmentDisplayName.Remove(objAttachmentDisplayName.Length - 1) + objAttachmentExtension;
                    }

                    //Traversing the all the Global mailItem Attachments to change/add the indexing after file removed and then append the fileExtension to the display names
                    foreach (Outlook.Attachment objAttachment in item.Attachments)
                    {
                        index++;

                        objAttachmentDisplayName = Path.GetFileNameWithoutExtension(objAttachment.DisplayName);
                        objAttachmentExtension = Path.GetExtension(objAttachment.DisplayName);

                        objAttachment.DisplayName = objAttachmentDisplayName + index + objAttachmentExtension;
                    }
                }

                if (item != null)
                    Marshal.ReleaseComObject(item);

                if (mailInspector != null)
                    Marshal.ReleaseComObject(mailInspector);
            }

            //If the attchments is not zero then rename the display names of the attachments (For outlook MailItem)
            if (olAttachments.Count > 0)
            {
                int index = 0;

                string objAttachmentDisplayName;
                string objAttachmentExtension;

                //Traversing the all the outlook mailItem Attachments to remove the indexing and then append the fileExtension to the display names
                foreach (Outlook.Attachment objAttachment in olAttachments)
                {
                    objAttachmentDisplayName = Path.GetFileNameWithoutExtension(objAttachment.DisplayName);
                    objAttachmentExtension = Path.GetExtension(objAttachment.DisplayName);

                    objAttachment.DisplayName = objAttachmentDisplayName.Remove(objAttachmentDisplayName.Length - 1) + objAttachmentExtension;
                }

                //Traversing the all the outlook mailItem Attachments to change/add the indexing after file removed and then append the fileExtension to the display names
                foreach (Outlook.Attachment objAttachment in olAttachments)
                {
                    index++;

                    objAttachmentDisplayName = Path.GetFileNameWithoutExtension(objAttachment.DisplayName);
                    objAttachmentExtension = Path.GetExtension(objAttachment.DisplayName);

                    objAttachment.DisplayName = objAttachmentDisplayName + index + objAttachmentExtension;
                }
            }

            attachmentList.Clear();

            //Again copying the attachments from temp to Original
            foreach (TempAttachments tempAttachment in tempAttachmentList)
            {
                attachmentList.Add(new Attachments()
                {
                    content = tempAttachment.content,
                    filename = tempAttachment.filename,
                    type = tempAttachment.type,
                    disposition = tempAttachment.disposition
                });
            }
        }

        private void ResetObjects()
        {
            attachmentList.Clear();
            tempAttachmentList.Clear();

            clearMailModel.personalizations.Clear();
            clearMailModel.attachments.Clear();
            clearMailModel.content.Clear();
            cmRecipientsList.Clear();

            //Releasing all the com objects to avoid error for the next email
            if (olMailItem != null)
                Marshal.FinalReleaseComObject(olMailItem);

            if (olRecipients != null)
                Marshal.FinalReleaseComObject(olRecipients);

            if (olRecipientTO != null)
                Marshal.FinalReleaseComObject(olRecipientTO);

            if (olRecipientCC != null)
                Marshal.FinalReleaseComObject(olRecipientCC);

            if (olRecipientBCC != null)
                Marshal.FinalReleaseComObject(olRecipientBCC);

            if (olAttachments != null)
                Marshal.ReleaseComObject(olAttachments);
        }

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
