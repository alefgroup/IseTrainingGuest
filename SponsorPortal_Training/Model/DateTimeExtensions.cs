using System;

namespace SponsorPortal_Training.Model
{
    public static class DateTimeExtensions
    {
        public static DateTime StartOfWeek(this DateTime dt, DayOfWeek startOfWeek)
        {
            int diff = dt.DayOfWeek - startOfWeek;
            if (diff < 0)
            {
                diff += 7;
            }
            return dt.AddDays(-1 * diff).Date;
        }

        public static DateTime EndOfWeek(this DateTime dt)
        {
            int diff = 7 - (int)dt.DayOfWeek;

            diff = diff == 7 ? 0 : diff;

            DateTime eow = dt.AddDays(diff).Date;

            return new DateTime(eow.Year, eow.Month, eow.Day, 23, 59, 59, 999) { };
        }

        public static DateTime Next(this DateTime from, DayOfWeek dayOfWeek)
        {
            int start = (int)from.DayOfWeek;
            int target = (int)dayOfWeek;
            if (target <= start)
                target += 7;
            return from.AddDays(target - start);
        }
    }
}
