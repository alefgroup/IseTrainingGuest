using System.Collections.Generic;

namespace SponsorPortal_Training.Model
{
    public class EmailJob
    {
        public bool Enabled { get; set; }
        public string Name { get; set; }
        public string Sender { get; set; }
        public string Subject { get; set; }
        public List<string> Recipients { get; set; }
        public SmtpServer SmtpServer { get; set; }
        public string ExportFileName { get; set; }
    }
}
