/* 
   Outlook File Guard Add-in

   This VSTO Outlook add-in hooks the Application.ItemSend event to inspect outgoing emails
   and help prevent accidental disclosure of reserved/private documents. Behavior:

   - Resolves the sender SMTP domain from the MailItem (handles Exchange/MAPI).
   - Scans recipients and flags the message if any recipient is outside the sender's domain.
   - Checks for attachments and (placeholder) detects whether attachments are marked as "private" or "reserved".
   - If external recipients and reserved attachments are detected, prompts the user to confirm sending.
   - Cancels the send when the user does not confirm.

   Notes and TODOs:
   - Debug MessageBox calls remain in the code for testing and should be removed or gated by a debug flag.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

using System.IO;
using System.Windows.Forms;
//using Word = Microsoft.Office.Interop.Word;  //TODO: Verificare se serve
using System.Runtime.InteropServices; // Necessario per ReleaseComObject


namespace OutlookFileGuard_addin
{
    public partial class ThisAddIn
    {
        //const string InternalDomain = "gmail.com";  // sostituito con funzione GetSenderDomainFromMailItem che acquisisce il dominio utente da Outlook Mail Item


        /// <summary>
        /// get the sender SMTP domain from the 'FROM:' field of the current email
        /// </summary>
        /// <param name="mailItem"> the MailItem object being to be sent</param>
        /// <returns> sender SMTP domain</returns>
        private string GetSenderDomainFromMailItem(Outlook.MailItem mailItem)
        {
            string senderAddress = mailItem.SenderEmailAddress;

            // if the address is an Exchange (MAPI), must be resolved in SMTP using PropertyAccessor
            if (mailItem.SenderEmailType == "EX")
                try
                {
                    dynamic dynamicMailItem = mailItem; // casting (static) mailItem on a dynamic variable (dynamic is needed by VSTO)
                    Outlook.Account senderAccount = dynamicMailItem.SendUsingAccount;
                    if (senderAccount != null && !string.IsNullOrEmpty(senderAccount.SmtpAddress))
                        senderAddress = senderAccount.SmtpAddress;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // Se si verifica questo errore, l'Add-in è probabilmente su una versione di Outlook
                    // precedente a 2010 e la proprietà non è disponibile (si passa al fallback).
                }
                catch (Exception ex)
                {
                    // In case of any other error
                    System.Diagnostics.Debug.WriteLine($"Errore in SendUsingAccount: {ex.Message}");
                }

            // Extracting domain from senderAddress
            int atIndex = senderAddress.LastIndexOf('@');
            if (atIndex > 0 && atIndex < senderAddress.Length - 1)  // if there is '@' and is not the first or the last char 
                return senderAddress.Substring(atIndex + 1).ToLowerInvariant();  // return the substring after '@'
            else
                return string.Empty;
        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Hooking the function Application_ItemSend to the ItemSend event (fired by Outlook just before sending an email)
            Application.ItemSend += Application_ItemSend;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        /// <summary>
        /// Main method of this Plugin
        /// here happen:
        ///  - if an email is being sent
        ///  - and if there are external recipients
        ///  - and if there are attachments
        ///  - and if there are private attachments
        ///  - then warn the user asking to confirm sending
        /// </summary>
        {
            try 
            {
                System.Windows.Forms.MessageBox.Show("ITEMSEND HIT");
                if (Item is Outlook.MailItem mailItem)  // checking if the item being sent is an email
                {
                    Cancel = false; // this row could be deleted: Cancel = false by default
                    string senderDomain = GetSenderDomainFromMailItem(mailItem);  //MessageBox.Show($"Sender domain: {senderDomain}", "ccc", MessageBoxButtons.OK, MessageBoxIcon.Stop);//{senderDomain}"); // for testing
                    if (AreExternalRecipientsPresent(mailItem, senderDomain)) //InternalDomain)) // checking if there are external recipients TODO: implementare la funzione
                        if (mailItem.Attachments.Count > 0) // checking if there are attachments
                            if (AreAttachementPrivate(mailItem)) // checking if there are private attachments TODO: implementare la funzione
                                if (!UserConfirmesSending())
                                {
                                    Cancel = true; // Cancel sending the email
                                    //MessageBox.Show("Sender domain: {senderDomain}", "titolo", MessageBoxButtons.OK, MessageBoxIcon.Stop);// {senderDomain}");
                                    //MessageBox.Show("Cancelled sending due to Private documents in attachments", "Sending Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in Application_ItemSend: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cancel = true;
            }
        }


        private bool AreExternalRecipientsPresent(Outlook.MailItem mailItem, string internalDomain)
        /// <summary>
        /// Verify if mailItem includes external domains (domains different from internalDomain)
        /// </summary>
        {
            foreach (Outlook.Recipient recipient in mailItem.Recipients)
            {
                //string address = recipient.Address;  non funge con gli indirizzi di exchange (raro ma non impossibile) quindi uso PropertyAccessor
                try
                {
                    recipient.Resolve();
                    if (recipient.Resolved)
                    {
                        Outlook.PropertyAccessor pa = recipient.PropertyAccessor;
                        string emailAddress = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString();  // Always returns SMTP address

                        // Se l'indirizzo SMTP NON finisce con il dominio interno, è esterno.
                        if (!emailAddress.EndsWith($"@{internalDomain}", StringComparison.OrdinalIgnoreCase)) return true; // at the first external address found, return true
                    }
                    else return true;  // if recipient address is not resolved, consider it as external by default
                }
                catch
                {
                    return true; // in case of any error, consider the address as external by default
                }
            }
            return false; // No external recipients found
        }


        private bool AreAttachementPrivate(Outlook.MailItem mailItem)
        {
            foreach (Outlook.Attachment attachment in mailItem.Attachments)
            {
                string tempPath = Path.Combine(Path.GetTempPath(), attachment.FileName);
                try
                {
                    // controlla se l'estensione è supportata prima di salvare (DocumentFormat.OpenXml ha bisogno di leggere il file dal disco)
                    if (!CrocoOfficeFileGuard.Core.ReservedFileManager.IsSupportedExtension(tempPath))
                        continue;  // passa al prossimo allegato
                    attachment.SaveAsFile(tempPath);

                    if (CrocoOfficeFileGuard.Core.ReservedFileManager.IsReservedInFile(tempPath))
                        return true;   // al primo allegato 'reserved' ritorna true 
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    // cancella sempre il file temporaneo
                    try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
                }
            }
            return false;
        }
         
        private bool UserConfirmesSending()
        {
            using (var form = new GuardAlertForm())
            {
                form.ShowDialog();
                return form.UserConfirmed;
            }
           /*
            // TODO: Implementare una finestra di dialogo che richiede all'utente di digitare 'yes' per confermare l'invio
            DialogResult result = MessageBox.Show("This email is addressed to external recipients and contains documents marked ‘reserved’. If you confirm sending, type ‘yes’; otherwise cancel", "Confirm Send", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            return result == DialogResult.Yes;
            */
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
