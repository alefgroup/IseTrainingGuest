using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SponsorPortal_Training
{
    public class GuestAccount
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string AccountDuration { get; set; }
        public DateTime AccountStartDate { get; set; }
        public DateTime AccountExpirationDate { get; set; }
        public string OptData1Course { get; set; }
        public string OptData2Mail { get; set; }
        public string OptData3Company { get; set; }
        public string OptData4Tel { get; set; }
    }
}
