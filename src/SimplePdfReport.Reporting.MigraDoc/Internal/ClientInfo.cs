using System;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;

namespace SimplePdfReport.Reporting.MigraDoc.Internal
{
    internal class ClientInfo
    {
        public static readonly Color Shading = new Color(243, 243, 243);

        public void Add(Section section, ClientReport clientReport)
        {
            var table = AddPatientInfoTable(section);

            //AddLeftInfo(table.Rows[0].Cells[0], clientReport);
            //AddLeftInfo(table.Rows[0], clientReport);
            //AddRightInfo(table.Rows[0].Cells[1], clientReport);
        }

        private Table AddPatientInfoTable(Section section)
        {
            var table = section.AddTable();
            table.Shading.Color = Shading;

            table.Rows.LeftIndent = 0;

            table.LeftPadding = Size.TableCellPadding;
            table.TopPadding = Size.TableCellPadding;
            table.RightPadding = Size.TableCellPadding;
            table.BottomPadding = Size.TableCellPadding;

            // Use two columns of equal width
            var columnWidth = Size.GetWidth(section) / 2.0;
            table.AddColumn(columnWidth);
            table.AddColumn(columnWidth);

            // Only one row is needed
            table.AddRow();

            return table;
        }

        //private void AddLeftInfo(Cell cell, ClientReport clientReport )
        //{
        //    // Add patient name and sex symbol
        //    var p1 = cell.AddParagraph();
        //    p1.Style = CustomStyles.ReportName;
        //    /* p1.AddText($"{clientReport.ClientName}");*///(SymphonyMedia-Linear)BillingSummarybyAgreement 
        //    p1.AddText("(SymphonyMedia-Linear) Billing Summary by Agreement");
        //    p1.AddSpace(2);
        //    //AddSexSymbol(p1, patient.Sex);

        //    // Add patient ID
        //   // var p2 = cell.AddParagraph();
        //    //p2.AddText("ID: ");
        //    //p2.AddFormattedText(patient.Id, TextFormat.Bold);
        //}

        private void AddLeftInfo(Row row, ClientReport clientReport)
        {
            // Add patient name and sex symbol
            //var p1 = row.Cells[0].AddParagraph();
            //p1.Style = CustomStyles.ReportName;
            /* p1.AddText($"{clientReport.ClientName}");*///(SymphonyMedia-Linear)BillingSummarybyAgreement 
            //p1.AddText("(SymphonyMedia-Linear) Billing Summary by Agreement ");
            //p1.AddSpace(2);
            //AddSexSymbol(p1, patient.Sex);

            // Add patient ID
            // var p2 = cell.AddParagraph();
            //p2.AddText("ID: ");
            //p2.AddFormattedText(patient.Id, TextFormat.Bold);
           // var p2 = row.Cells[1].AddParagraph();
           // p2.AddText("Billing Summary by Agreement");
        }
        private void AddSexSymbol(Paragraph p, Sex sex)
        {
            p.AddImage(new SexSymbol(sex).GetMigraDocFileName());
        }

        private void AddRightInfo(Cell cell, ClientReport clientReport)
        {
            var p = cell.AddParagraph();

            // Add birthdate
            p.AddText("Client Name: ");
            p.AddFormattedText((clientReport.ClientName), TextFormat.Bold);

            p.AddLineBreak();

            // Add doctor name
            p.AddText(" ");
            p.AddFormattedText($"{clientReport.ReportParameters}", TextFormat.Bold);
        }

        private string Format(DateTime birthdate)
        {
            return $"{birthdate:d} (age {Age(birthdate)})";
        }

        // See http://stackoverflow.com/a/1404/1383366
        private int Age(DateTime birthdate)
        {
            var today = DateTime.Today;
            int age = today.Year - birthdate.Year;
            return birthdate.AddYears(age) > today ? age - 1 : age;
        }
    }
}