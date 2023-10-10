using NLog;
using System;
using AleFIT.Configuration;
using SponsorPortal_Training.Model;
using System.IO;
using System.Collections.Generic;

namespace SponsorPortal_Training.AppCode
{
    public class Settings
    {
        public static Settings Instance { get; } = new Settings();

        public bool DbAuth;
        public string DbIpAddress;
        public string DbName;
        public string DbUsername;
        public string DbPassword;
        public string ConDbString;

        public string IseUrl;

        public string SponsorPortalUrl;
        public string SponsorPortalUsername;
        public string SponsorPortalPassword;
        public string SponsorPortalGuestType;
        public string SponsorPortalLocation;
        public string SponsorPortalPortalId;

        public string Printer;
        public bool PrintingEnabled;
        public string SqlQuery;

        public bool GuestnameModification;
        public string GuestNamePrefix;

        public RdpFeature RdpFeature;

        public List<EmailJob> EmailJobs;

        private Settings()
        {
            Logger logger = LogManager.GetCurrentClassLogger();
            logger.Info("Reading settings");
            
            var appDir = AppDomain.CurrentDomain.BaseDirectory;
            Config config = new Config(appDir + @"AppData/Config.xml", "Config");

            DbName = config.ReadConfig("Database", "DbName");
            DbIpAddress = config.ReadConfig("Database", "IPAddress");
            DbUsername = config.ReadConfig("Database", "DbUsername");
            DbPassword = config.ReadConfig("Database", "DbPassword");
            DbAuth = config.ReadConfigAttrBool("Database", "UseWindowsAuthentication");

            if (DbAuth)
            {
                ConDbString = $"data source={DbIpAddress};initial catalog={DbName};integrated security=true;persist security info=True;";
            }
            else
            {
                ConDbString = $"data source={DbIpAddress};initial catalog={DbName};integrated security=false;persist security info=True;" +
                $"User ID={DbUsername};Password={DbPassword}";
            }

            IseUrl = config.ReadConfig("Ise", "BaseUrl");

            SponsorPortalUrl = config.ReadConfig("SponsorPortal", "SponsorPortalUrl");
            SponsorPortalUsername = config.ReadConfig("SponsorPortal", "Username");
            SponsorPortalPassword = config.ReadConfig("SponsorPortal", "Password");
            SponsorPortalGuestType = config.ReadConfig("SponsorPortal", "GuestType");
            SponsorPortalPortalId = config.ReadConfig("SponsorPortal", "PortalId");
            SponsorPortalLocation = config.ReadConfig("SponsorPortal", "Location");

            Printer = config.ReadConfig("Printers", "Printer");
            PrintingEnabled = config.ReadConfigAttrBool("Printers", "enabled");
            SqlQuery = config.ReadConfig("SqlQuerys", "Query");

            GuestnameModification = config.ReadConfigAttrBool("Output/GuestName/Prefix", "enabled");
            GuestNamePrefix = config.ReadConfig("Output", "GuestName/Prefix");

            RdpFeature = new RdpFeature
            {
                Enabled = config.ReadConfigAttrBool("AdvancedFeatures/RDPExtension", "enabled"),
                Counter = new Counter {
                    Max = config.ReadConfigInt("AdvancedFeatures/RDPExtension/Counter", "Max"),
                    Min = config.ReadConfigInt("AdvancedFeatures/RDPExtension/Counter", "Min")
                },
                CustomFieldName = config.ReadConfig("AdvancedFeatures/RDPExtension", "CustomFieldName"),
                CourseListEnabled = config.ReadConfigAttrBool("AdvancedFeatures/RDPExtension/Conditions/CourseListPath", "enabled"),
                CourseListPath = config.ReadConfig("AdvancedFeatures/RDPExtension/Conditions", "CourseListPath"),
                ResourceVersionEnabled = config.ReadConfigAttrBool("AdvancedFeatures/RDPExtension/Conditions/ResourceVersion", "enabled"),
                ResourceVersion = config.ReadConfig("AdvancedFeatures/RDPExtension/Conditions", "ResourceVersion"),
                Prefix = config.ReadConfig("AdvancedFeatures/RDPExtension", "Prefix"),
                Courses = config.ReadConfigAttrBool("AdvancedFeatures/RDPExtension/Conditions/CourseListPath", "enabled") == true ?
                                        File.ReadAllLines(config.ReadConfig("AdvancedFeatures/RDPExtension/Conditions", "CourseListPath")) :
                                            null
            };

            EmailJobs = new List<EmailJob>();
            var emailJobsList = config.GetConfigNodes("AdvancedFeatures/Emailing");
            for (var i = 0; i < emailJobsList.Count; i++)
            {
                var emailJob = emailJobsList[i].SelectSingleNode("EmailJob");

                EmailJobs.Add(
                    new EmailJob { 
                        Enabled = bool.Parse(emailJob.Attributes["enabled"].Value),
                        Name = emailJob.SelectSingleNode("Name")?.InnerText,
                        Sender = emailJob.SelectSingleNode("Sender")?.InnerText,
                        Subject = emailJob.SelectSingleNode("Subject")?.InnerText,
                        Recipients = GetRecipientsFromConfig(emailJob.SelectNodes("Recipients/Recipient")),
                        ExportFileName = emailJob.SelectSingleNode("ExportFileName")?.InnerText,
                        SmtpServer = new SmtpServer {                             
                            Server = emailJob.SelectSingleNode("SMTPServer/Address")?.InnerText,
                            Port = int.Parse(emailJob.SelectSingleNode("SMTPServer/Port")?.InnerText)                       
                        }
                    });                 
            }

            logger.Info("Reading settings done");
         }

        private List<string> GetRecipientsFromConfig(System.Xml.XmlNodeList xmlNodeList)
        {
            var recList = new List<string>();
            for (var i = 0; i < xmlNodeList.Count; i++)
            {
                var recipient = xmlNodeList[i].InnerText;
                recList.Add(recipient);
            };

            return recList;
        }
    }
}
