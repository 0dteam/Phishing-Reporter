/*
 * Developer: Abdulla Albreiki
 * Github: https://github.com/0dteam
 * licensed under the GNU General Public License v3.0
 */
 
using Microsoft.Office.Core;
using PhishingReporter.Properties;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Security.Cryptography;
using HtmlAgilityPack;
using System.Collections.Generic;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PhishingReporter
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public Bitmap getGroup1Image(IRibbonControl control)
        {
            return Resources.phishing;
        }

        // Functions
        public void reportPhishing(Office.IRibbonControl control)
        {
            var areYouSure = MessageBox.Show("Do you want to report this email to the Information Security Team as a potential phishing attempt?", "Are you sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if(areYouSure == DialogResult.Yes)
            {
                reportPhishingEmailToSecurityTeam(control);
            }
        }

        /*
         *  Helper functions 
         */

        private void reportPhishingEmailToSecurityTeam(IRibbonControl control)
        {

            Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
            string reportedItemType = "NaN"; // email, contact, appointment ...etc
            string reportedItemHeaders = "NaN";

            if(selection.Count < 1) // no item is selected
            {
                MessageBox.Show("Select an email before reporting.", "Error");
            }
            else if(selection.Count > 1) // many items selected
            {
                MessageBox.Show("You can report 1 email at a time.", "Error");
            }
            else // only 1 item is selected
            {
                if (selection[1] is Outlook.MeetingItem || selection[1] is Outlook.ContactItem || selection[1] is Outlook.AppointmentItem || selection[1] is Outlook.TaskItem || selection[1] is Outlook.MailItem)
                {
                    // Identify the reported item type
                    if (selection[1] is Outlook.MeetingItem)
                    {
                        reportedItemType = "MeetingItem";
                    }
                    else if (selection[1] is Outlook.ContactItem)
                    {
                        reportedItemType = "ContactItem";
                    }
                    else if (selection[1] is Outlook.AppointmentItem)
                    {
                        reportedItemType = "AppointmentItem";
                    }
                    else if (selection[1] is Outlook.TaskItem)
                    {
                        reportedItemType = "TaskItem";
                    }
                    else if (selection[1] is Outlook.MailItem)
                    {
                        reportedItemType = "MailItem";
                    }

                    // Prepare Reported Email
                    Object mailItemObj = (selection[1] as object) as Object;
                    MailItem mailItem = (reportedItemType == "MailItem") ? selection[1] as MailItem : null; // If the selected item is an email

                    MailItem reportEmail = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                    reportEmail.Attachments.Add(selection[1] as Object);

                    try
                    {

                        reportEmail.To = Properties.Settings.Default.infosec_email;
                        reportEmail.Subject = (reportedItemType == "MailItem") ? "[POTENTIAL PHISH] " + mailItem.Subject : "[POTENTIAL PHISH] " + reportedItemType; // If reporting email, include subject; otherwise, state the type of the reported item

                        // Get Email Headers
                        if (reportedItemType == "MailItem")
                        {
                            reportedItemHeaders = mailItem.HeaderString();
                        }
                        else
                        {
                            reportedItemHeaders = "Headers were not extracted because the reported item is not an email. It is " + reportedItemType;
                        }

                        // Check if the email is a simulated phishing campaign by Information Security Team
                        string simulatedPhishingURL = GoPhishIntegration.setReportURL(reportedItemHeaders);

                        if (simulatedPhishingURL != "NaN")
                        {
                            string simulatedPhishingResponse = GoPhishIntegration.sendReportNotificationToServer(simulatedPhishingURL);
                            // DEBUG: to check if reporting email reaches GoPhish Portal
                            // MessageBox.Show(simulatedPhishingURL + " --- " + simulatedPhishingResponse);

                            // Update GoPhish Campaigns Reported counter
                            Properties.Settings.Default.gophish_reports_counter++;

                            // Thanks
                            MessageBox.Show("Good job! You have reported a simulated phishing campaign sent by the Information Security Team.", "We have a winner!");
                        }
                        else
                        {

                            // Update Suspecious Emails Reported counter
                            Properties.Settings.Default.suspecious_reports_counter++;

                            // Prepare the email body
                            reportEmail.Body = GetCurrentUserInfos();
                            reportEmail.Body += "\n";
                            reportEmail.Body += GetBasicInfo(mailItem);
                            reportEmail.Body += "\n";
                            reportEmail.Body += GetURLsAndAttachmentsInfo(mailItem);
                            reportEmail.Body += "\n";
                            reportEmail.Body += "---------- Headers ----------";
                            reportEmail.Body += "\n" + reportedItemHeaders;
                            reportEmail.Body += "\n";
                            reportEmail.Body += GetPluginDetails() + "\n\n";

                            reportEmail.Save();
                            //reportEmail.Display(); // Helps in debugginng
                            reportEmail.Send(); // Automatically send the email

                            // Enable if you want a second popup for confirmation
                            // MessageBox.Show("Thank you for reporting. We will review this report soon. - Information Security Team", "Thank you");
                        }

                        // Delete the reported email
                        mailItem.Delete();

                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("There was an error! An automatic email was sent to the support to resolve the issue.", "Do not worry");

                        MailItem errorEmail = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                        errorEmail.To = Properties.Settings.Default.support_email;
                        errorEmail.Subject = "[Outlook Addin Error]";
                        errorEmail.Body = ("Addin error message: " + ex);
                        errorEmail.Save();
                        //errorEmail.Display(); // Helps in debugginng
                        errorEmail.Send(); // Automatically send the email
                    }
                }
                else
                {
                    MessageBox.Show("You cannot report this item", "Error");
                }
            }
        }

        public String GetBasicInfo(MailItem mailItem)
        {
            Outlook.MAPIFolder parentFolder = mailItem.Parent as Outlook.MAPIFolder;
            string FolderLocation = parentFolder.FolderPath;
            string basicInfo = "---------- Basic Info ----------";
            basicInfo += "\n - Reported from: \"" + FolderLocation + "\" Folder";
            basicInfo += "\n - OS: " + Environment.OSVersion + " " + (Environment.Is64BitOperatingSystem ? "(64bit)" : "(32bit)");
            basicInfo += "\n - Agent: " + Globals.ThisAddIn.Application.Name + " "  + Globals.ThisAddIn.Application.Version;
            basicInfo += "\n - Suspecious emails reported: " + Properties.Settings.Default.suspecious_reports_counter;
            basicInfo += "\n - GoPhish campaigns reported: " + Properties.Settings.Default.gophish_reports_counter;
            basicInfo += "\n";
            return basicInfo;
        }


        public String GetCurrentUserInfos()
        {
            string str = "---------- User Information ----------";
            str += "\n - Domain:" + Environment.UserDomainName;
            str += "\n - Username:" + Environment.UserName;
            str += "\n - Machine name:" + Environment.MachineName;

            Outlook.AddressEntry addrEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
            if (addrEntry.Type == "EX")
            {
                Outlook.ExchangeUser currentUser =
                    Globals.ThisAddIn.Application.Session.CurrentUser.
                    AddressEntry.GetExchangeUser();
                if (currentUser != null)
                {
                    str += "\n - Name: " + currentUser.Name;
                    str += "\n - STMP address: " + currentUser.PrimarySmtpAddress;
                    str += "\n - Title: " + currentUser.JobTitle;
                    str += "\n - Department: " + currentUser.Department;
                    str += "\n - Location: " + currentUser.OfficeLocation;
                    str += "\n - Business phone: " + currentUser.BusinessTelephoneNumber;
                    str += "\n - Mobile phone: " + currentUser.MobileTelephoneNumber;

                }
            }
            return str + "\n";
        }

        public String GetURLsAndAttachmentsInfo(MailItem mailItem)
        {
            string urls_and_attachments = "---------- URLs and Attachments ----------";

            var domainsInEmail = new List<string>();

            var emailHTML = mailItem.HTMLBody;
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(emailHTML);

            // extracting all links
            var urlsText = "";
            var urlNodes = doc.DocumentNode.SelectNodes("//a[@href]");
            if(urlNodes != null)
            {
                urlsText = "\n\n # of URLs: " + doc.DocumentNode.SelectNodes("//a[@href]").Count;
                foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//a[@href]"))
                {
                    HtmlAttribute att = link.Attributes["href"];
                    if (att.Value.Contains("a"))
                    {
                        urlsText += "\n --> URL: " + att.Value.Replace(":", "[:]");
                        // Domain Extraction
                        try
                        {
                            domainsInEmail.Add(new Uri(att.Value).Host);
                        }
                        catch (UriFormatException)
                        {
                            // Try to process URL as email address. Example -> <a href="mailto:ask@0d.ae">...etc
                            String emailAtChar = "@";
                            int ix = att.Value.IndexOf(emailAtChar);
                            if (ix != -1)
                            {
                                string emailDomain = att.Value.Substring(ix + emailAtChar.Length);
                                try
                                {
                                    domainsInEmail.Add(new Uri(emailDomain).Host);
                                }
                                catch (UriFormatException)
                                {
                                    // if it fails again, ignore domain extraction
                                    Console.WriteLine("Bad url: {0}", emailDomain);
                                }
                            }
                        }
                    }
                }
            }
            else
                urlsText = "\n\n # of URLs: 0";

            // Get domains
            domainsInEmail = domainsInEmail.Distinct().ToList();
            urls_and_attachments += "\n # of unique Domains: " + domainsInEmail.Count;
            foreach (string item in domainsInEmail)
            {
                urls_and_attachments += "\n --> Domain: " + item.Replace(":", "[:]");
            }

            // Add Urls
            urls_and_attachments += urlsText;

            urls_and_attachments += "\n\n # of Attachments: " + mailItem.Attachments.Count;
            foreach (Attachment a in mailItem.Attachments)
            {
                // Save attachment as txt file temporarily to get its hashes (saves under User's Temp folder)
                var filePath = Environment.ExpandEnvironmentVariables(@"%TEMP%\Outlook-Phishaddin-" + a.DisplayName + ".txt");
                a.SaveAsFile(filePath);

                string fileHash_md5 = "";
                string fileHash_sha256 = "";
                if (File.Exists(filePath))
                {
                    fileHash_md5 = CalculateMD5(filePath);
                    fileHash_sha256 = GetHashSha256(filePath);
                    // Delete file after getting the hashes
                    File.Delete(filePath);
                }
                urls_and_attachments += "\n --> Attachment: " + a.FileName + " (" + a.Size + " bytes)\n\t\tMD5: " + fileHash_md5 + "\n\t\tSha256: " + fileHash_sha256 + "\n";
            }
            return urls_and_attachments;
        }



        public String GetPluginDetails()
        {
            string pluginDetails = "---------- Report Phishing Plugin ----------";
            pluginDetails += "\n - Version: " + Properties.Settings.Default.plugin_version;
            pluginDetails += "\n - Usage: Report phishing emails to the Information Security Team.";
            pluginDetails += "\n - Support: " + Properties.Settings.Default.support_email;
            // Do not delete this. I worked hard to deliver this product for FREE.
            pluginDetails += "\n - Developer: Abdulla Albreiki (aalbraiki@hotmail.com)";
            return pluginDetails;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PhishingReporter.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        static string CalculateMD5(string filename)
        {
            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(filename))
                {
                    var hash = md5.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                }
            }
        }
        private string GetHashSha256(string filename)
        {
            using (FileStream stream = File.OpenRead(filename))
            {
                SHA256Managed sha = new SHA256Managed();
                byte[] shaHash = sha.ComputeHash(stream);
                string result = "";
                foreach (byte b in shaHash) result += b.ToString("x2");
                return result;
            }
        }

        #endregion
    }

    public static class MailItemExtensions
    {
        private const string HeaderRegex =
            @"^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)" +
                "(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)";
        private const string TransportMessageHeadersSchema =
            "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        public static string[] Headers(this MailItem mailItem, string name)
        {
            var headers = mailItem.HeaderLookup();
            if (headers.Contains(name))
                return headers[name].ToArray();
            return new string[0];
        }

        public static ILookup<string, string> HeaderLookup(this MailItem mailItem)
        {
            var headerString = mailItem.HeaderString();
            var headerMatches = Regex.Matches
                (headerString, HeaderRegex, RegexOptions.Multiline).Cast<Match>();
            return headerMatches.ToLookup(
                h => h.Groups["header_key"].Value,
                h => h.Groups["header_value"].Value);
        }

        public static string HeaderString(this MailItem mailItem)
        {
            return (string)mailItem.PropertyAccessor
                .GetProperty(TransportMessageHeadersSchema);
        }

    }
}
 