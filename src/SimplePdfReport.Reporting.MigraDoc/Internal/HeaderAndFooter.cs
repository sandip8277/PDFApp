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
      //section.AddParagraph($"Image {structureSet.Image.Id} " +
      //                     $"taken {structureSet.Image.CreationTime:g}");
      section.Headers.Primary.AddParagraph("Report Date Range: 04/2019-05/2020");
      section.Headers.Primary.AddParagraph(clientReport.ReportParameters);
      section.Headers.Primary.AddParagraph("Service:All");
      section.Headers.Primary.AddParagraph("----------------------------------------------------------------------------------");

    }
    private void AddHeader(Section section, ClientReport clientReport )
        {
      var header = section.Headers.Primary.AddParagraph();
      //Paragraph paragraph = section.Headers.Primary.AddParagraph();
      //paragraph.AddFormattedText("(SymphonyMedia-Linear) Billing Summary by Agreement", TextFormat.Bold);
      //paragraph.Format.Alignment = "Center";
      //paragraph.AddText("\nCreated on ");
      header.Format.AddTabStop(Size.GetWidth(section), TabAlignment.Center);
      //Image image = header.Section.AddImage("D:/Shubham/POC/Nexstar.jpg");
      Image image = section.Headers.Primary.AddImage("D:/APDFData/Images/Nexstar.jpg");
      image.Height = "2.0cm";
      image.LockAspectRatio = true;
      //image.RelativeVertical = RelativeVertical.Margin;
      //image.RelativeHorizontal = RelativeHorizontal.Margin;
      image.Top = ShapePosition.Top;
      image.Left = ShapePosition.Left;
      //image.WrapFormat.Style = WrapStyle.TopBottom;
      //image.RelativeHorizontal = RelativeHorizontal.Margin;
      //header.Style = "font-size:150px;";
      //image.WrapFormat.DistanceTop = "0.00 cm";
      header.AddTab();
      //header.AddLineBreak(); 
      //header.AddLineBreak();
      //header.AddLineBreak();
      header.AddFormattedText("(SymphonyMedia-Linear) Billing Summary by Agreement", TextFormat.Bold);
      //header.AddLineBreak();
      //header.AddTab();
      //AddHeading(section, clientReport);
      ////header.AddLineBreak();
      //header.AddText("(SymphonyMedia-Linear) Billing Summary by Agreement ");
      header.AddLineBreak();
      //header.Format.Alignment = ParagraphAlignment.Center;
      header.AddFormattedText("Report Date Range: 04/2019-05/2020");
      header.AddLineBreak();
            //header.Format.AddTabStop(Size.GetWidth(section), TabAlignment.Left);
      header.AddFormattedText(clientReport.ReportParameters);
      header.AddLineBreak();
      header.AddFormattedText("Service:All");
      header.AddLineBreak();
      header.AddFormattedText("----------------------------------------------------------------------------------");

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