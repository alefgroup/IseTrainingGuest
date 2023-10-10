using System.Collections.Generic;

namespace SponsorPortal_Training.Model
{
    public class RdpFeature
    {
        public bool Enabled { get; set; }
        public string CustomFieldName { get; set; }
        public string Prefix { get; set; }
        public Counter Counter { get; set; }
        public bool CourseListEnabled { get; set; }
        public string CourseListPath { get; set; }
        public bool ResourceVersionEnabled { get; set; }
        public string ResourceVersion { get; set; }
        public string[] Courses { get; set; }
    }
}
