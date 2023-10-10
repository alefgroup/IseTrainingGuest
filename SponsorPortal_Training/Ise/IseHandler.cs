using NLog;
using RestSharp;
using RestSharp.Authenticators;
using SponsorPortal_Training.AppCode;
using SponsorPortal_Training.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace SponsorPortal_Training.Ise
{
    public class IseHandler
    {
        private static readonly string IseIpAddress = Settings.Instance.IseUrl;
        private static readonly string IseUser = Settings.Instance.SponsorPortalUsername;
        private static readonly string IsePass = Settings.Instance.SponsorPortalPassword;

        private Logger _logger = LogManager.GetCurrentClassLogger();

        public List<string> GetGuestsCompanyIds()
        {
            _logger.Info("IseHandler - GetGuestList() - Getting guest user data from ISE");
            var iseUrl = GetUrlAddress(IseIpAddress);
            _logger.Debug($"IseHandler - GetGuestList() - ISE URL: {iseUrl}");
            List<IseGuestUser> guestList = new List<IseGuestUser>();
            var client = new RestClient(iseUrl);
            client.Authenticator = new HttpBasicAuthenticator(IseUser, IsePass);

            var request = new RestRequest();
            request.Method = Method.GET;
            request.AddHeader("Accept", "application/vnd.com.cisco.ise.identity.guestuser.2.0+xml");
            request.AddHeader("Accept-Search-Result", "application/vnd.com.cisco.ise.ers.searchresult.2.0+xml");

            TimeSpan differenceNext = DateTimeExtensions.EndOfWeek(DateTime.Now.AddDays(7D)) - DateTimeExtensions.StartOfWeek(DateTime.Now.AddDays(7D), DayOfWeek.Monday);

            var daysRemain = (int)Math.Abs(Math.Round(differenceNext.TotalDays)) + 1;
            _logger.Debug($"IseHandler - GetGuestList() - nextWeek=YES daysRemain:{daysRemain} ");

            CultureInfo culture = new CultureInfo("en-US");

            List<string> filters = new List<string>();
            List<string> searchStrings = new List<string>();

            for (int i = 0; daysRemain >= i; i++)
            {
                filters.Add("toDate");
                string dateSearch = DateTimeExtensions.StartOfWeek(DateTime.Now.AddDays(7D), DayOfWeek.Monday).AddDays(i).ToString("dd-MMM-yy", culture);
                searchStrings.Add("CONTAINS." + dateSearch);
                string resourceUrl = GetUrlWithFilter(GetResourceGuestAll(), filters, searchStrings);
                _logger.Debug($"IseHandler - GetGuestList() - searching resource={resourceUrl}");
                request.Resource = resourceUrl;
                IRestResponse response = client.Execute(request);
                var content = response.Content; // raw content as string
                XDocument doc = XDocument.Parse(content);
                var namespaceManager = new XmlNamespaceManager(new NameTable());
                namespaceManager.AddNamespace("ns5", "ers.ise.cisco.com");
                var listOfLinks = doc.XPathSelectElements("//ns5:resource", namespaceManager).ToList();
                if (listOfLinks.Any())
                {
                    foreach (var node in listOfLinks)
                    {
                        guestList.Add(new IseGuestUser { IseUserId = node.Attribute("id").Value ?? "" });
                    }
                }
            }

            RestRequest detailRequest = new RestRequest();
            detailRequest.Method = Method.GET;
            detailRequest.AddHeader("Accept", "application/vnd.com.cisco.ise.identity.guestuser.2.0+xml");
            if (guestList.Any())
            {
                foreach (var iseGuestUser in guestList)
                {
                    _logger.Debug($"IseHandler - GetGuestList() - GuestUserId: {iseGuestUser.IseUserId}");
                }

                guestList = guestList.Select(c => GetUpdatedGuestUser(c, client, detailRequest))
                                        .Where(x => x.OptData2Company.StartsWith(Settings.Instance.RdpFeature.Prefix))
                                                .ToList();

                if (guestList.Any())
                {
                    _logger.Info($"IseHandler - GetGuestList() - List of users");
                    return guestList.Select(x => x.OptData2Company.Substring(Settings.Instance.RdpFeature.Prefix.Count())).ToList();

                }
                else
                {
                    _logger.Debug("IseHandler - GetGuestList() - No RDP guest users has been found!");
                    return new List<string>();
                }
            }
            else
            {
                _logger.Debug("IseHandler - GetGuestList() - No guest users has been found!");
                return new List<string>();
            }            
        }

        private IseGuestUser GetUpdatedGuestUser(IseGuestUser iseGuestUser, RestClient client, RestRequest request)
        {

            request.Resource = GetResourceGuestById(iseGuestUser.IseUserId);
            IRestResponse<IseGuestUser> response = client.Execute<IseGuestUser>(request);
            var content = response.Data; // raw content as string
            _logger.Debug($"Retrieving data for user ", iseGuestUser.IseUserId);
            XDocument doc = XDocument.Parse(response.Content);
            var namespaceManager = new XmlNamespaceManager(new NameTable());
            namespaceManager.AddNamespace("ns4", "identity.ers.ise.cisco.com");

            string companyName;
            try
            {
                companyName = doc.XPathSelectElement("//ns4:guestuser/customFields/entry/value[../key/text()=\'ui_company_text_label\']", namespaceManager).Value;

            }
            catch (Exception)
            {
                companyName = "";
                _logger.Error("Element key ui_company_text_label missing");
            }


            var iseUpdatedUser = new IseGuestUser
            {
                IseUserId = iseGuestUser.IseUserId,
                AccountExpirationDate = content.AccountExpirationDate,
                AccountStartDate = content.AccountStartDate,
                FirstName = content.FirstName,
                LastName = content.LastName,
                Password = content.Password,
                Username = content.Username,
                OptData2Company = companyName,
                AccountDuration = content.AccountDuration
            };
            _logger.Debug($"Updated user data:\n" +
                          $"ID: {iseUpdatedUser.IseUserId}\n" +
                          $"Name: {iseUpdatedUser.FirstName} {iseUpdatedUser.LastName}\n" +
                          $"Date: {iseUpdatedUser.AccountStartDate} {iseUpdatedUser.AccountExpirationDate}\n" +
                          $"Company: {iseUpdatedUser.OptData2Company}\n");
            return iseUpdatedUser;
        }

        private string GetResourceGuestById(string id)
        {
            return "/ers/config/guestuser/" + id;
        }
        private string GetResourceGuestAll()
        {
            return "/ers/config/guestuser";
        }

        private string GetUrlAddress(string ipAddress)
        {

            _logger.Info($"IseHandler - GetUrlAddress() - Getting url address with ip {ipAddress}");
            var url = $"https://{ipAddress}:9060";
            _logger.Debug($"IseHandler - GetUrlAddress() - Url address:{url}");
            return url;
        }

        private string GetUrlWithFilter(string urlAddress, List<string> filters, List<string> searchString)
        {
            if (filters.Count != searchString.Count)
            {
                throw new InvalidOperationException("Length of filter lists differs from each other. ");
            }
            StringBuilder builder = new StringBuilder();
            builder.Append(urlAddress);

            string fill = "&filter=";
            builder.AppendFormat("{0}", "?size=100");
            for (int i = 0; i < filters.Count; i++)
            {
                if (string.IsNullOrEmpty(searchString[i]))
                {
                    continue;
                };

                if (filters.Count == i + 1)
                {
                    builder.Append(fill);
                    builder.Append(filters[i]);
                    builder.Append(".");
                    builder.Append(searchString[i]);
                    continue;
                };
                builder.Append(fill);
                builder.Append(filters[i]);
                builder.Append(".");
                builder.Append(searchString[i]);
            }
            return builder.ToString();
        }
    }
}
