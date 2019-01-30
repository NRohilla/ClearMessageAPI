using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClearMessageOutlookAddIn
{
    partial class ClearMessage
    {
        ApiHelper apiHelper = new ApiHelper();

        #region Form Region Factory 


        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)]
        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Contact)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("ClearMessageOutlookAddIn.ClearMessageRegion")]
        public partial class ClearMessageFactory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void ClearMessageFactory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
                //dynamic sampleObject = (dynamic)e.OutlookItem;

                //if (sampleObject.MessageClass == "IPM.Contact")
                //    e.Cancel = false;

                //if (sampleObject.MessageClass == "IPM.Note")
                //    e.Cancel = true;
            }
        }

        #endregion

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void ClearMessage_FormRegionShowing(object sender, System.EventArgs e)
        {
            dynamic sampleObject = (dynamic)this.OutlookItem;

            if (sampleObject.MessageClass == "IPM.Contact")
            {
                Outlook.ContactItem contactItem = (Outlook.ContactItem)this.OutlookItem;
                if (!string.IsNullOrWhiteSpace(contactItem.FirstName) && !string.IsNullOrWhiteSpace(contactItem.Email1Address) && !string.IsNullOrWhiteSpace(contactItem.FileAs))
                {
                    var CustomProperty = contactItem.UserProperties.Find("SendViaClearMessage", true);
                    if (CustomProperty != null)
                    {
                        chkSendViaClearMessage.Checked = contactItem.UserProperties["SendViaClearMessage"].Value;
                    }
                    else
                    {
                        contactItem.UserProperties.Add("SendViaClearMessage", Outlook.OlUserPropertyType.olYesNo, true, Type.Missing);
                        contactItem.UserProperties["SendViaClearMessage"].Value = chkSendViaClearMessage.Checked;
                        contactItem.Subject = contactItem.LastNameAndFirstName;
                        contactItem.Save();
                    }
                }
            }
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void ClearMessage_FormRegionClosed(object sender, System.EventArgs e)
        {
            dynamic sampleObject = (dynamic)this.OutlookItem;
            try
            {
                if (sampleObject.MessageClass == "IPM.Contact")
                {
                    Outlook.ContactItem contactItem = (Outlook.ContactItem)this.OutlookItem;

                    if (chkSendViaClearMessage.Checked)
                    {
                        if (!string.IsNullOrWhiteSpace(contactItem.FirstName) && !string.IsNullOrWhiteSpace(contactItem.Email1Address) && !string.IsNullOrWhiteSpace(contactItem.FileAs))
                        {
                            contactItem.UserProperties.Add("SendViaClearMessage", Outlook.OlUserPropertyType.olYesNo, true, Type.Missing);
                            contactItem.UserProperties["SendViaClearMessage"].Value = chkSendViaClearMessage.Checked;
                            contactItem.Subject = contactItem.LastNameAndFirstName;
                            contactItem.Save();
                        }

                        if (!string.IsNullOrEmpty(contactItem.MobileTelephoneNumber))
                        {
                            RegisterModel registerModel = new RegisterModel();

                            registerModel.email = contactItem.Email1Address;
                            registerModel.phone = contactItem.MobileTelephoneNumber;

                            string jsonRegisterModel = JsonConvert.SerializeObject(registerModel);
                            SendRegisterationMail(jsonRegisterModel);
                        }
                        else
                        {
                            MessageBox.Show("Please enter the mobile number otherwise contact will not get registered on Clear Message portal.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrWhiteSpace(contactItem.FirstName) && !string.IsNullOrWhiteSpace(contactItem.Email1Address) && !string.IsNullOrWhiteSpace(contactItem.FileAs))
                        {
                            var CustomProperty = contactItem.UserProperties.Find("SendViaClearMessage", true);
                            if (CustomProperty != null)
                            {
                                contactItem.UserProperties["SendViaClearMessage"].Value = chkSendViaClearMessage.Checked;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            { }
            finally
            {
            }
        }

        private async Task SendRegisterationMail(string registerModel)
        {
            HttpClient client = apiHelper.InitializeClient();
            using (var content = new StringContent(registerModel, System.Text.Encoding.Default, "application/json"))
            {
                using (HttpResponseMessage response = await client.PostAsync("v1/admin/receiver", content))
                {
                    string responseData = await response.Content.ReadAsStringAsync();
                }
            }
        }

    }
}
