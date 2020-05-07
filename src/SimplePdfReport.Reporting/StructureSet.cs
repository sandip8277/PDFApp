using System.Collections.Generic;

namespace SimplePdfReport.Reporting
{
    public class StructureSet
    {
        public string ReportParameters { get; set; }
        //public Image Image { get; set; }
        public List<Structure> Structures { get; set; }
    }
}