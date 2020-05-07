using System;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;

namespace SimplePdfReport.Reporting.MigraDoc.Internal
{
    internal class HeaderAndFooter
    {
        public void Add(Section section, ReportData data)
        {
            AddHeader(section, data.ClientReport);
            AddFooter(section);
        }
        private void AddHeading(Section section, ClientReport clientReport)
        {
            section.Headers.Primary.AddParagraph("Report Date Range: 04/2019-05/2020");
            section.Headers.Primary.AddParagraph(clientReport.ReportParameters);
            section.Headers.Primary.AddParagraph("Service:All");
            section.Headers.Primary.AddParagraph("----------------------------------------------------------------------------------");

        }
        private void AddHeader(Section section, ClientReport clientReport)
        {

            Image image = section.Headers.Primary.AddImage("D:/TestPDFData/Images/Nexstar.jpg");
            image.Height = "2.0cm";
            image.LockAspectRatio = true;
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Left;

            var header = section.Headers.Primary.AddParagraph();

            header.AddFormattedText("(SymphonyMedia-Linear) Billing Summary by Agreement", TextFormat.Bold);
            header.AddTab();
            header.AddLineBreak();
            header.AddLineBreak();
            header.AddLineBreak();
            header.AddFormattedText("Report Date Range: 04/2019-05/2020");
            header.AddLineBreak();
            header.AddFormattedText("Type: CLOSE");
            header.AddLineBreak();
            header.AddFormattedText(clientReport.ReportParameters);
            header.AddLineBreak();
            header.AddFormattedText("Service:All");
            header.AddLineBreak();
           

        }

        private void AddFooter(Section section)
        {
            var footer = section.Footers.Primary.AddParagraph();
            footer.Format.AddTabStop(Size.GetWidth(section), TabAlignment.Right);


            footer.AddTab();
            footer.AddText("Symphony Media -Linear");//Here need to print Client Short Name
            footer.AddTab();
            footer.AddText("Page ");
            footer.AddPageField();
            footer.AddText(" of ");
            footer.AddNumPagesField();
        }
    }
}