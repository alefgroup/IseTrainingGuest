using SponsorPortal_Training.Model;
using System;
using System.Collections.Generic;
using NLog;
using System.Net.Mail;
using System.IO;
using AleFIT.Utils;
using AleFIT_Library;
using System.Data;
using System.Linq;

namespace SponsorPortal_Training.Email
{
    public class EmailSender
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly EmailJob _sumEmailSettings;        
        public EmailSender(EmailJob sumEmailSettings)
        {
            _sumEmailSettings = sumEmailSettings;
        }

        internal void SendEmail(List<GuestAccount> guestAccounts, DataTable dataTableAccounts)
        {
            try
            {
                using (var smtp = new SmtpClient(_sumEmailSettings.SmtpServer.Server))
                {
                    using (var mem = new MemoryStream())
                    {
                        CreateCsvReport(guestAccounts, dataTableAccounts, mem);

                        foreach (var recipient in _sumEmailSettings.Recipients)
                        {
                            try
                            {
                                Logger.Info($"EmailSender - Send() - SumEmail - to: {recipient}, from: {_sumEmailSettings.Sender}, subject: {_sumEmailSettings.Subject}");

                                var message = new MailMessage(_sumEmailSettings.Sender, recipient,
                                        FormatSubject(_sumEmailSettings.Subject), null);

                                message.Attachments.Add(EmailUtilities.CreateUtf8Attachment(_sumEmailSettings.ExportFileName, mem.ToArray()));

                                smtp.Send(message);

                                Logger.Info("EmailSender - Send() - Exception export for special devices report successfully sent.");
                            }
                            catch (Exception ex)
                            {
                                Logger.Error($"EmailSender - Send() - Exception while sending email, ex message: {ex.Message} ");
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Error($"EmailSender - Send() - Exception in sending method, ex.msg: {e.Message}");
            }
        }

        private string FormatSubject(string subject)
        {
            return subject?.Replace("\r", " ").Replace("\n", " ").Substring(0, subject.Length < 168 ? subject.Length : 168);
        }

        private void CreateCsvReport(List<GuestAccount> guestAccounts, DataTable dataTableAccounts, Stream stream)
        {
            using (var sw = CsvUtils.CreateStreamWriter(stream, CsvUtils.CsvStreamOptions.Utf8WithBom))
            {
                sw.WriteLine(string.Join(";", new List<string> { "Attendee", "No_", "Seminar No_", "Username", "FIRST Name", "Surname", "RDP-Extension","Password",
                    "Company Name", "Participant E-Mail", "Participant Phone No_", "StartDate", "EndDate", "Version", "Event Type" }));
                
                var counter = 1;
                foreach (var guest in guestAccounts)
                {
                    
                    var guestDbRow = dataTableAccounts.Select().FirstOrDefault(db => db["Email"].ToString() == guest.OptData2Mail);

                    sw.WriteLine(string.Join(";", new List<string>
                        {
                            counter.ToString(),
                            guestDbRow == null ? "" : guestDbRow[0].ToString(),
                            guest.OptData1Course ?? string.Empty,
                            guest.Username ?? string.Empty,
                            guest.FirstName ?? string.Empty,
                            guest.LastName ?? string.Empty,
                            guest.OptData3Company ?? string.Empty, // !! - v pripade RDP enable, se do company extra attributu dava RDP-xx - xx je cislo rdp sessiony
                            guest.Password ?? string.Empty,
                            guestDbRow == null ? "" : guestDbRow[4].ToString(),
                            guest.OptData2Mail ?? string.Empty,
                            guest.OptData4Tel ?? string.Empty,
                            guest.AccountStartDate.ToString("dd.MM.yyyy"),
                            guest.AccountExpirationDate.ToString("dd.MM.yyyy"),
                            "",
                            guestDbRow == null ? "" : guestDbRow[11].ToString(),
                        }));
                    counter++;
                }
                sw.Flush();
            }
        }
    }
}