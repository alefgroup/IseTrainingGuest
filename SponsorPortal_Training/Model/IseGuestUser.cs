using System;

namespace SponsorPortal_Training.Model
{
    public class IseGuestUser
    {
        public string IseUserId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public DateTime AccountStartDate { get; set; }
        public DateTime AccountExpirationDate { get; set; }
        public string AccountDuration { get; set; }
        public string OptData2Company { get; set; }

    }
}
