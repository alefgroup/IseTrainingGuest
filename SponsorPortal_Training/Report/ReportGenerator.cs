using System;
using System.Collections.Generic;
using System.IO;
using NLog;
using Microsoft.Reporting.WinForms;
using SponsorPortal_Training.AppCode;

namespace SponsorPortal_Training.Report
{
    class ReportGenerator
    {
        private readonly Logger _logger;

        public ReportGenerator()
        {
            _logger = LogManager.GetCurrentClassLogger();
        }

        public void GenerateUserReport(GuestAccount account)
        {
            _logger.Info("Trying generate report for user: " + account.Username);

            var report = new LocalReport();
            report.ReportPath =
                @"Report\UserReport.rdlc";

            _logger.Info("Generating DataSet for guest account: " + account.Username);
            List<GuestAccount> listGuestAccounts = GetListOfObject(account);
            var rds = new ReportDataSource
            {
                Name = "DataSet1", 
                Value = listGuestAccounts
            };
            report.DataSources.Add(rds);

            //Generate PDF document for guest account
            //GeneratePdf(report, account);

            //Send guest account report to printer
            _logger.Info($"Printing accounts on the printer is {(Settings.Instance.PrintingEnabled == true ? "enabled" : "disabled")}");
            if (Settings.Instance.PrintingEnabled)
            {
                _logger.Info($"Printing accounts ...");
                PrintGuestReport(listGuestAccounts);
                _logger.Info($"Printing accounts done");
            }            
        }

        private void PrintGuestReport(List<GuestAccount> listGuestAccounts)
        {
            _logger.Info("Sending reports to printer");
            using (Printer print = new Printer())
            {
                try
                {
                    print.Run(listGuestAccounts);
                    _logger.Info("Print OK");
                }
                catch (Exception e)
                {
                    _logger.Error("Error with printing: " + e.Message);
                    _logger.Error("Stop App");
                    Environment.Exit(1);
                }

            }
        }

        private void GeneratePdf(LocalReport report, GuestAccount account)
        {
            _logger.Info("Generating PDF");
            byte[] bytes = report.Render("PDF");

            using (FileStream fs = new FileStream($"Report_for_user_{account.Username}.pdf", FileMode.Create))
            {
                fs.Write(bytes, 0, bytes.Length);
            }
        }

        private List<GuestAccount> GetListOfObject(GuestAccount account)
        {
            var list = new List<GuestAccount>(1);
            list.Add(account);
            return list;
        }
    }
}
