using System.Data.SqlClient;
using NLog;
using SponsorPortal_Training.AppCode;
using System.Data;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Text;
using System.Xml.Linq;
using SponsorPortal_Training.Report;
using System.Linq;
using SponsorPortal_Training.Model;
using SponsorPortal_Training.Ise;
using SponsorPortal_Training.Email;

namespace SponsorPortal_Training
{
    class Program
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static DataTable _dt;
        private static string _iseUri;
        private static List<GuestAccount> _guestAccounts = new List<GuestAccount>();
        private static string _guestPrefix = "";
        public static List<RdpSession> rdpAvailableSessions;
        private static IseHandler _iseHandler = new IseHandler();

        static void Main()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

            Logger.Info("App SponsorPortal-Training is starting");
            PrintSettingToConsole();
            Logger.Info("Set guest prefix");
            _guestPrefix = Settings.Instance.GuestnameModification == true ? Settings.Instance.GuestNamePrefix : "";

            Logger.Info("Creating IseUri");
            _iseUri = GetUri(Settings.Instance.SponsorPortalUrl);

            Logger.Info("Creating dataTable");
            _dt = InitDataTable();

            Logger.Info("Sending request to sql - read candidates info");
            SendSqlRequest();

            Logger.Info("Print all candidates info to console");
            PrintDataTableToConsole();
            Logger.Info("Print all candidates info to console - done");

            Logger.Info("Sending request to create ISE sponsor portal users");

            if (Settings.Instance.RdpFeature.Enabled)
            {
                List<string> guestsCompanyIds = _iseHandler.GetGuestsCompanyIds();
                rdpAvailableSessions = CreateArrayByMinMax(Settings.Instance.RdpFeature.Counter.Min, Settings.Instance.RdpFeature.Counter.Max);
                rdpAvailableSessions.ForEach(x =>
                {
                    if (guestsCompanyIds.Exists(c=>c.Equals(x.ValueAsText)))
                    {
                        x.IsAvailable = false;
                    }
                });
            }

            for (int i = 0; i < _dt.Rows.Count; i++)
            {
                Logger.Info($"Creating {i + 1} guest account");
                string aditionalAttribute = "";

                if (Settings.Instance.RdpFeature.Enabled)
                {
                    if (Settings.Instance.RdpFeature.ResourceVersionEnabled)
                    {
                        if (_dt.Rows[i].Field<string>(10).ToLower().Contains(Settings.Instance.RdpFeature.ResourceVersion.ToLower()))
                        {
                            if (rdpAvailableSessions.Exists(x => x.IsAvailable == true))
                            {
                                RdpSession rdpAvailableItem = rdpAvailableSessions.First(x => x.IsAvailable == true);
                                string rdpId = rdpAvailableItem.ValueAsText;
                                aditionalAttribute = @"<company>" + Settings.Instance.RdpFeature.Prefix + rdpId + @"</company>";
                                rdpAvailableItem.IsAvailable = false;
                            }
                            else
                            {
                                Logger.Info($"There is no free RdpId - guest account will not be generated!!!");
                                continue;
                            }
                        }
                    }else if (Settings.Instance.RdpFeature.CourseListEnabled)
                    {
                        if (Settings.Instance.RdpFeature.Courses.Contains(_dt.Rows[i].Field<string>(1)))
                        {
                            if (rdpAvailableSessions.Exists(x => x.IsAvailable == true))
                            {
                                RdpSession rdpAvailableItem = rdpAvailableSessions.First(x => x.IsAvailable == true);
                                string rdpId = rdpAvailableItem.ValueAsText;
                                aditionalAttribute = @"<company>" + Settings.Instance.RdpFeature.Prefix + rdpId + @"</company>";
                                rdpAvailableItem.IsAvailable = false;
                            }
                            else
                            {
                                Logger.Info($"There is no free RdpId - guest account will not be generated!!!");
                                continue;
                            }
                        }
                    }
                }
                GuestAccount guestAccount = CreateGuestAccount(i, _iseUri, aditionalAttribute);

                if (guestAccount != null)
                {
                    _guestAccounts.Add(guestAccount);
                }
            }

            PrintAccountListToConsole();

            Logger.Info("Generate reports for accounts and send to printer if enabled");
            var generator = new ReportGenerator();
            foreach (var account in _guestAccounts)
            {
                generator.GenerateUserReport(account);
            }

            if (Settings.Instance.EmailJobs.Exists(emailJob => emailJob.Name == "SummaryEmail"))
            {
                var emailJobObject = Settings.Instance.EmailJobs.Find(emailJob => emailJob.Name == "SummaryEmail");
                Logger.Info($"Summary Email Feature - {(emailJobObject.Enabled ? "enabled" : "disabled")}");
                if (emailJobObject.Enabled)
                {
                    Logger.Info($"Summary Email Feature - Start");
                    var emailSender = new EmailSender(emailJobObject);
                    emailSender.SendEmail(_guestAccounts, _dt);
                    Logger.Info($"Summary Email Feature - End");
                }
            }
        }

        private static List<RdpSession> CreateArrayByMinMax(int min, int max)
        {
            List<RdpSession> rdpSession = new List<RdpSession>();
            for (int i = min; i <= max; i++)
            {
                rdpSession.Add(new RdpSession
                {
                    IsAvailable = true,
                    ValueAsInt = i,
                    ValueAsText = i <= 9 ? "0" + i.ToString() : i.ToString()
                });
            }

            return rdpSession;
        }

        private static void PrintAccountListToConsole()
        {
            Logger.Info("Printing all created accounts");
            foreach (var account in _guestAccounts)
            {
                Logger.Info(account.Username + " " + account.Password + " " + account.FirstName + " " + account.LastName + " " +account.AccountStartDate + " " + account.AccountExpirationDate + " " + account.AccountDuration + " " + account.OptData1Course + " " + account.OptData2Mail + " " + account.OptData3Company + " " + account.OptData4Tel);
            }
        }
        private static string GetUri(string sponsorPortalUrl)
        {
            int indexOf = sponsorPortalUrl.IndexOf(":");
            if (indexOf == -1)
            {
                Logger.Error("Error parsing IP adress from SponsorPortalUrl, please, check configuration");
                Logger.Error("Exiting application");
                Environment.Exit(0);
            }
            string iseIpAddress = sponsorPortalUrl.Substring(0, indexOf);
            string uri = string.Format("https://{0}:9060/ers/config/guestuser", iseIpAddress);
            return uri;
        }
        private static GuestAccount CreateGuestAccount(int i, string uri, string additionalAttribute)
        {            
            try
            {
                string guestUserName = _guestPrefix + _dt.Rows[i].Field<string>(2).Substring(0, 1) + _dt.Rows[i].Field<string>(3);
                Logger.Info($"Start creating guest-account - Guest UserName: {guestUserName}");
                if (IseGuestAccountExist(uri + "/name/" + guestUserName))
                {
                    Logger.Info($"Guest with UserName {guestUserName} exist. Generating new GuestName.");
                    guestUserName = GenerateNewGuestName(guestUserName);
                }

                using (var client = new WebClient())
                {
                    client.Headers.Add("Accept", "*/*");
                    client.Credentials = new NetworkCredential(Settings.Instance.SponsorPortalUsername, Settings.Instance.SponsorPortalPassword);
                    client.Headers.Add("Content-type", "application/vnd.com.cisco.ise.identity.guestuser.2.0+xml;charset=utf-8");

                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11| SecurityProtocolType.Tls12;
                    ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                    var formatedUsername = FormatingUserName(guestUserName);

                    var payload = @"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
                                    <ns4:guestuser description=""ERS Example user "" id =""1234567890"" name =""guestUser"" xmlns:ers=""ers.ise.cisco.com"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"" xmlns:ns4=""identity.ers.ise.cisco.com"" >
                                        <customFields>
                                            <entry>
                                                <key>ui_name_of_training_class_or_person_being_visited_text_label</key>
                                                <value>" + _dt.Rows[i].Field<string>(1) + @"</value>
                                            </entry>
                                            <entry>
                                                <key>ui_email_text_label</key>
                                                <value>" + _dt.Rows[i].Field<string>(5) + @"</value>
                                            </entry>
                                            <entry>
                                                <key>ui_phone_number_text_label</key>
                                                <value>" + _dt.Rows[i].Field<string>(6) + @"</value>
                                            </entry>
                                        </customFields>
                                        <guestAccessInfo>
                                            <fromDate>" + _dt.Rows[i].Field<string>(7) + @"</fromDate>
                                            <location>" + Settings.Instance.SponsorPortalLocation + @"</location>
                                            <toDate>" + _dt.Rows[i].Field<string>(8) + @"</toDate>
                                            <validDays>" + _dt.Rows[i].Field<string>(9) + @"</validDays>                                            
                                        </guestAccessInfo>
                                        <guestInfo>
                                            "
                                                         + additionalAttribute +
                                            @"
                                            <enabled>true</enabled>
                                            <firstName>" + _dt.Rows[i].Field<string>(2) + @"</firstName>
                                            <lastName>" + _dt.Rows[i].Field<string>(3) + @"</lastName>
                                            <userName>" + formatedUsername + @"</userName>
                                        </guestInfo>
                                        <guestType>" + Settings.Instance.SponsorPortalGuestType + @"</guestType>
                                        <portalId>" + Settings.Instance.SponsorPortalPortalId + @"</portalId>
                                    </ns4:guestuser>";

                    Logger.Debug(payload);
                    Logger.Debug("uri + " + uri);

                    byte[] data = Encoding.UTF8.GetBytes(payload);
                    client.UploadData(new Uri(uri), "POST", data);

                    //Get User Id from WebClient header
                    WebHeaderCollection myWebHeaderCollection = client.ResponseHeaders;
                    string userId = GetUserId(myWebHeaderCollection);
                    Logger.Debug("userID: " + userId);

                    //Get User info By Id - from ISE API
                    string responseBody = GetUserInfoById(uri + "/" + userId);
                    Logger.Debug("responseBody: " + responseBody);

                    //Parse Data from responseBody and get GuestAccount
                    GuestAccount guestAccount = GetGuestAccount(responseBody);
                    return guestAccount;
                }
            }
            catch (WebException ex)
            {
                Logger.Error("Error creating guest-account: " + ex);
                return null;
            }
        }

        private static string GenerateNewGuestName(string oldUserName)
        {
            string newUserName = oldUserName;
            string lastChar = oldUserName.Substring(oldUserName.Length - 1);

            int accountNumber = 0;
            bool success = Int32.TryParse(lastChar, out accountNumber);

            if (success)
            {
                accountNumber += 1;
                newUserName = newUserName.Remove(newUserName.Length - 1) + accountNumber;
            }
            else
            {
                newUserName += 1;
            }

            return newUserName;
        }

        private static bool IseGuestAccountExist(string uri)
        {
            Logger.Info($"I'm checking to see if a user account exists on the ISE - uri: {uri}");

            try
            {                
                using (var webClient = new WebClient()) {
                    webClient.Headers.Add("Accept", "application/json");
                    webClient.Credentials = new NetworkCredential(Settings.Instance.SponsorPortalUsername, Settings.Instance.SponsorPortalPassword);
                    byte[] bResponseData = webClient.DownloadData(uri);
                    return Encoding.UTF8.GetString(bResponseData).Contains(uri);
                }
            }
            catch (WebException e)
            {
                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpWebResponse myHttpWebResponse = (HttpWebResponse)e.Response;
                    Logger.Debug("IseGuestAccountExist() WebException - Errorcode: {0}", (int)myHttpWebResponse.StatusCode);
                    Logger.Debug("IseGuestAccountExist() WebException - Message: {0}", e.Message);

                    switch (myHttpWebResponse.StatusCode)
                    {
                        case HttpStatusCode.NotFound:
                            Logger.Debug("User account doesn´t exist in ISE.");
                            break;
                        default:
                            Logger.Debug("Unexpected response from ISE if an user account exist. We will try to create an account with an existing username.");
                            break;
                    }
                }
                else
                {
                    Logger.Debug("Unexpected Protocol error in response form ISE, Status: {0}", e.Status);
                    Logger.Debug("We will try to create an account with an existing username.");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"An attempt to determine the existence of a user account ended in error. We will try to create an account with an existing username. Exc message: {ex.Message}");                
            }

            return false;
        }

        private static GuestAccount GetGuestAccount(string responseBody)
        {
            try
            {
                Logger.Info("GetGuestAccount - parsing data");
                XDocument xdoc = XDocument.Parse(responseBody);
                if (xdoc.Root == null)
                {
                    Logger.Error("Invalid data exception, xdoc.Root is null");
                    return null;
                }

                string optData1Course = "";
                string optData2Mail = "";
                string optData3Company = "";
                string optData4Tel = "";

                var customFields = xdoc.Root.Descendants();
                foreach (var entry in customFields.Elements("entry"))
                {
                    switch (entry.Element("key").Value)
                    {
                        case "ui_name_of_training_class_or_person_being_visited_text_label":
                            optData1Course = entry.Element("value").Value;
                            break;
                        case "ui_email_text_label":
                            optData2Mail = entry.Element("value").Value;
                            break;
                        case "ui_company_text_label":
                            optData3Company = entry.Element("value").Value;
                            break;
                        case "ui_phone_number_text_label":
                            optData4Tel = entry.Element("value").Value;
                            break;
                    }
                };

                var guestInfo = xdoc.Root.Element("guestInfo");
                string firstname = guestInfo?.Element("firstName")?.Value;
                string lastname = guestInfo?.Element("lastName")?.Value;
                string userName = guestInfo?.Element("userName")?.Value;
                string password = guestInfo?.Element("password")?.Value;

                var guestAccessInfo = xdoc.Root.Element("guestAccessInfo");
                DateTime fromDateIseFormat = DateTime.ParseExact(guestAccessInfo?.Element("fromDate")?.Value, "MM/dd/yyyy HH:mm", CultureInfo.InvariantCulture);
                string fromDatePortalFormat = fromDateIseFormat.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                DateTime toDateIseFormat = DateTime.ParseExact(guestAccessInfo?.Element("toDate")?.Value, "MM/dd/yyyy HH:mm", CultureInfo.InvariantCulture);
                string toDatePortalFormat = toDateIseFormat.ToString("yyyy-MM-dd " + new TimeSpan(23, 59, 59), CultureInfo.InvariantCulture);
                string accountDuration = (((toDateIseFormat - fromDateIseFormat).Days) + 1).ToString();

                if (userName == null || password == null || accountDuration == null)
                {
                    Logger.Error("Invalid data exception, some data is null");
                    return null;
                }

                var account = new GuestAccount
                {
                    FirstName = firstname,
                    LastName = lastname,
                    Username = userName,
                    Password = password,
                    AccountDuration = accountDuration,
                    AccountStartDate = DateTime.ParseExact(fromDatePortalFormat, "yyyy-MM-dd HH:mm:ss", null),
                    AccountExpirationDate = DateTime.ParseExact(toDatePortalFormat, "yyyy-MM-dd HH:mm:ss", null),
                    OptData1Course = optData1Course,
                    OptData2Mail = optData2Mail,
                    OptData3Company = optData3Company,
                    OptData4Tel = optData4Tel,
                };
                return account;
            }
            catch (Exception ex)
            {
                Logger.Error("Creating account failed. ");
                Logger.Error("Exception in parsing data from after creating account: " + ex);
                return null;
            }
        }
        private static string GetUserInfoById(string uriId)
        {
            try
            {
                Logger.Info("GetUserInfoById starting");
                var client = new WebClient();
                client.Headers.Add("Accept", "application/vnd.com.cisco.ise.identity.guestuser.2.0+xml");
                client.Credentials = new NetworkCredential(Settings.Instance.SponsorPortalUsername, Settings.Instance.SponsorPortalPassword);
                byte[] bResponseData = client.DownloadData(uriId);
                string sResponseData = Encoding.UTF8.GetString(bResponseData);

                Logger.Info("GetUserInfoById is completed");
                return sResponseData;
            }
            catch (WebException e)
            {
                Logger.Error("Cannot get user info by Id: " + e);
                return null;
            }
        }
        private static string GetUserId(WebHeaderCollection myWebHeaderCollection)
        {
            if (myWebHeaderCollection.Count != 0)
            {
                const int lenghtUserId = 36;
                string userId = myWebHeaderCollection.Get("Location").Substring(myWebHeaderCollection.Get("Location").Length - lenghtUserId);
                Logger.Info("userId");
                return userId;
            }
            return null;
        }
        private static void PrintDataTableToConsole()
        {

            for (int i = 0; i < _dt.Rows.Count; i++)
            {
                Logger.Info($"Attendee [{i+1}]");
                for (int j = 0; j < _dt.Columns.Count-1; j++)
                {
                    Logger.Info(_dt.Rows[i].Field<string>(j));
                }
            }
        }
        private static void PrintSettingToConsole()
        {
            Logger.Debug(Settings.Instance.DbName);
            Logger.Debug(Settings.Instance.DbAuth.ToString);
            Logger.Debug(Settings.Instance.DbIpAddress);
            Logger.Debug(Settings.Instance.DbUsername);
            Logger.Debug(HidePassword(Settings.Instance.ConDbString));
            Logger.Debug(Settings.Instance.IseUrl);
            Logger.Debug(Settings.Instance.SponsorPortalGuestType);
            Logger.Debug(Settings.Instance.SponsorPortalLocation);
            Logger.Debug(Settings.Instance.SponsorPortalPortalId);
            Logger.Debug(Settings.Instance.SponsorPortalUrl);
            Logger.Debug(Settings.Instance.SponsorPortalUsername);
            Logger.Debug(Settings.Instance.Printer);
            Logger.Debug(Settings.Instance.PrintingEnabled);
            Logger.Debug(Settings.Instance.SqlQuery);
            Logger.Debug(Settings.Instance.GuestnameModification);
            Logger.Debug(Settings.Instance.GuestNamePrefix);
            Logger.Debug(Settings.Instance.RdpFeature.Enabled);
            if (Settings.Instance.RdpFeature.Enabled)
            {
                Logger.Debug(Settings.Instance.RdpFeature.Counter.Max);
                Logger.Debug(Settings.Instance.RdpFeature.Counter.Min);
                Logger.Debug(Settings.Instance.RdpFeature.CustomFieldName);                
                Logger.Debug(Settings.Instance.RdpFeature.Prefix);
                if (Settings.Instance.RdpFeature.CourseListEnabled)
                {
                    Logger.Debug(Settings.Instance.RdpFeature.CourseListPath);
                    Logger.Debug(string.Join(",", Settings.Instance.RdpFeature.Courses));
                }
                if (Settings.Instance.RdpFeature.ResourceVersionEnabled)
                {
                    Logger.Debug(Settings.Instance.RdpFeature.ResourceVersion);
                }
            }
        }

        private static string HidePassword(string conDbString)
        {
            return conDbString.Substring(0, conDbString.LastIndexOf('=')) + "=***";
        }

        private static void SendSqlRequest()
        {
            try
            {
                SqlConnection conn = new SqlConnection(Settings.Instance.ConDbString);
                conn.Open();
                Logger.Info("Connect to Db OK");
                SqlCommand cmd = new SqlCommand(Settings.Instance.SqlQuery, conn);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    FulfilDataTable(reader);
                }
                conn.Close();
            }
            catch (Exception e)
            {
                Logger.Error("Cannot connect to DB, msg: " + e.Message);
                Logger.Error("Stop App");
                Environment.Exit(1);
            }
        }
        private static void FulfilDataTable(SqlDataReader reader)
        {
            while (reader.Read())
            {
                Logger.Info($"Candidate info: {reader[0]}, {reader[1]}, {reader[2]}, {reader[3]}, {reader[4]}, {reader[5]}, {reader[6]}, {reader[7]}, {reader[8]}, {reader[9]}, {reader[10]}");

                string fromDate = ConvertToIseTimeFrom(reader[7]);
                string toDate = ConvertToIseTimeTo(reader[8]);
                string validDays = (((DateTime) reader[8] - (DateTime) reader[7]).Days).ToString();
                validDays = validDays == "0" ? "1" : validDays;

                _dt.Rows.Add(reader[0], reader[1], RemDiakritics(reader[2].ToString()), RemDiakritics(reader[3].ToString()), RemDiakritics(reader[4].ToString()), RemDiakritics(reader[5].ToString()), reader[6], fromDate, toDate, validDays, reader[9], reader[10]);           
            }
        }
        private static string RemDiakritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }
        private static string ConvertToIseTimeTo(object sqlTime)
        {
            DateTime dtSqlTime = (DateTime)sqlTime;
            string toDate = dtSqlTime.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture) + " 23:59";
            return toDate;
        }
        private static string ConvertToIseTimeFrom(object sqlTime)
        {
            DateTime dtSqlTime = (DateTime) sqlTime;
            string fromDate = dtSqlTime.ToString("MM/dd/yyyy HH:mm", CultureInfo.InvariantCulture);
            return fromDate;
        }
        private static DataTable InitDataTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("No", typeof(string));
            table.Columns.Add("SeminarNo", typeof(string));
            table.Columns.Add("FirstName", typeof(string));
            table.Columns.Add("Surname", typeof(string));
            table.Columns.Add("Company", typeof(string));
            table.Columns.Add("Email", typeof(string));
            table.Columns.Add("Tel", typeof(string));
            table.Columns.Add("StartDate", typeof(string));
            table.Columns.Add("EndDate", typeof(string));
            table.Columns.Add("ValidDays", typeof(string));
            table.Columns.Add("Version", typeof(string));
            table.Columns.Add("Event Type", typeof(string));

            return table;
        }

        private static string FormatingUserName(string userName)
        {
            string[] parts = userName.Trim().Split(' ');
            string formattedUserName = string.Join("-", parts);
            return formattedUserName;
        }
    }
}
