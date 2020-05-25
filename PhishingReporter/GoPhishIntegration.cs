/*
 * Developer: Abdulla Albreiki
 * Github: https://github.com/0dteam
 * licensed under the GNU General Public License v3.0
 */

using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Text;
using System;
using System.Security.Authentication;
using PhishingReporter;

namespace PhishingReporter
{
    static class GoPhishIntegration
    {

        static string GoPhishURL = PhishingReporter.Properties.Settings.Default.gophish_url + ":" + Properties.Settings.Default.gophish_listener_port;
        static string URLrequest = GoPhishURL + "/report?rid=USERID";
        static string GoPhishHeader = PhishingReporter.Properties.Settings.Default.gophish_custom_header;
        static string WebExpID = GoPhishHeader + @": [0-9a-zA-Z]+";
        static string WebExpPrefix = GoPhishHeader + @": ";

        // This function constructs GoPhish report url from a custom header in the simulated phishing campaign email
        public static string setReportURL(string headers)
        {
            // Extract GoPhish Custom Header (X-GOPHISH-ASMN: USERID0123)
            var match = new Regex(WebExpID).Match(headers);

            foreach (var group in match.Groups)
            {
                if(group.ToString().Trim()!=string.Empty)
                {
                    // Extract User ID from the header (USERID0123)
                    string user_id = group.ToString().Replace(WebExpPrefix, string.Empty);

                    // Build reporting URL, something like this -> https[:]//GOPHISHURL:PORT/report?rid=USERID
                    string report_url = URLrequest.Replace(@"USERID", user_id);
                    return report_url;
                }
            }

            // else, no header was found -> No report tracking URL
            return "NaN";
        }

        public const SslProtocols _strTls12 = (SslProtocols)0x00000C00;
        public const SecurityProtocolType Tls12 = (SecurityProtocolType)_strTls12;

        public static string sendReportNotificationToServer(string reportURL)
        {
            ServicePointManager.SecurityProtocol = Tls12;

            try
            {
                var request = (HttpWebRequest)WebRequest.Create(reportURL);            
                var response = (HttpWebResponse)request.GetResponse();
                string html = new StreamReader(response.GetResponseStream()).ReadToEnd();
                return "OK";
            }
            catch (System.Exception exc)
            {
                return "ERROR"; // "GoPhish Listener is not responding or there is no Internet connection."
            }
        }
    }
}

