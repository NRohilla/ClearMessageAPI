using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
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

            }
        }

        #endregion

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void ClearMessage_FormRegionShowing(object sender, System.EventArgs e)
        {
            dynamic outlookObject = (dynamic)this.OutlookItem;
            if (outlookObject.MessageClass == "IPM.Note")
            {
                chkSendViaClearMessage.Visible = false;
            }


            //Check the object is of type Contact
            if (outlookObject.MessageClass == "IPM.Contact")
            {
                chkSendViaClearMessage.Visible = true;
                Outlook.ContactItem contactItem = (Outlook.ContactItem)this.OutlookItem;

                //Getting the custom property defined for the checkbox
                var CustomProperty = contactItem.UserProperties.Find("SendViaClearMessage", true);
                if (CustomProperty != null)
                {
                    //Setting the state of the checkbox from the property
                    chkSendViaClearMessage.Checked = contactItem.UserProperties["SendViaClearMessage"].Value;
                }
                else
                {
                    //This process will get exceuted if the old contact doesn't have the property.
                    contactItem.UserProperties.Add("SendViaClearMessage", Outlook.OlUserPropertyType.olYesNo, true, Type.Missing);
                    contactItem.UserProperties["SendViaClearMessage"].Value = chkSendViaClearMessage.Checked;
                    contactItem.Subject = contactItem.LastNameAndFirstName;
                    contactItem.Save();
                }

                Marshal.FinalReleaseComObject(contactItem);
            }
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void ClearMessage_FormRegionClosed(object sender, System.EventArgs e)
        {
            dynamic outlookObject = (dynamic)this.OutlookItem;
            try
            {
                if (outlookObject.MessageClass == "IPM.Contact")
                {
                    Outlook.ContactItem contactItem = (Outlook.ContactItem)this.OutlookItem;

                    if (chkSendViaClearMessage.Checked)
                    {
                        //Check the below fields are not empty or null
                        if (!string.IsNullOrWhiteSpace(contactItem.FirstName) && !string.IsNullOrWhiteSpace(contactItem.Email1Address) && !string.IsNullOrWhiteSpace(contactItem.FileAs))
                        {
                            //Create the user defined property for clear message checkbox
                            contactItem.UserProperties.Add("SendViaClearMessage", Outlook.OlUserPropertyType.olYesNo, true, Type.Missing);
                            contactItem.UserProperties["SendViaClearMessage"].Value = chkSendViaClearMessage.Checked;
                            contactItem.Subject = contactItem.LastNameAndFirstName;
                        }

                        //If Checkbox is checked and not the mobile number is null - We will register the user on Clear Message Portal
                        if (!string.IsNullOrEmpty(contactItem.MobileTelephoneNumber))
                        {
                            RegisterModel registerModel = new RegisterModel();

                            if (!string.IsNullOrWhiteSpace(contactItem.Email1Address))
                                registerModel.email = contactItem.Email1AddressType == "SMTP" ? contactItem.Email1Address.Trim() : GetSmtpEmaillAddress(contactItem);

                            if (!string.IsNullOrWhiteSpace(contactItem.Email2Address))
                                registerModel.email = contactItem.Email2AddressType == "SMTP" ? contactItem.Email2Address.Trim() : GetSmtpEmaillAddress(contactItem);

                            if (!string.IsNullOrWhiteSpace(contactItem.Email3Address))
                                registerModel.email = contactItem.Email3AddressType == "SMTP" ? contactItem.Email3Address.Trim() : GetSmtpEmaillAddress(contactItem);

                            registerModel.phone = contactItem.MobileTelephoneNumber;

                            //Serialize the Register model for sending to API
                            string jsonRegisterModel = JsonConvert.SerializeObject(registerModel);

                            //Call for the API endpoint for making a new receiver
                            RegisterUserOnClearMessage(jsonRegisterModel);
                        }
                        else
                        {
                            MessageBox.Show("Please enter the mobile number otherwise contact will not get registered on Clear Message portal.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        //If the checkbox is not checked then update the custom property if created.
                        if (!string.IsNullOrWhiteSpace(contactItem.FirstName) && !string.IsNullOrWhiteSpace(contactItem.Email1Address) && !string.IsNullOrWhiteSpace(contactItem.FileAs))
                        {
                            var CustomProperty = contactItem.UserProperties.Find("SendViaClearMessage", true);
                            if (CustomProperty != null)
                            {
                                contactItem.UserProperties["SendViaClearMessage"].Value = chkSendViaClearMessage.Checked;
                            }
                        }
                    }

                    contactItem.Save();

                    Marshal.FinalReleaseComObject(contactItem);
                }
            }
            catch (Exception ex)
            { }
            finally
            { }
        }

        private string GetSmtpEmaillAddress(Outlook.ContactItem contactItem)
        {
            dynamic contactProp = contactItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/8084001F");

            return contactProp.ToString().Trim();
        }

        private async Task RegisterUserOnClearMessage(string registerModel)
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
