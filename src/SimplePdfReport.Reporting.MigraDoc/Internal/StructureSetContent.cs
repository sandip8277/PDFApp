using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using System.Collections.Generic;
using System.Linq;

namespace SimplePdfReport.Reporting.MigraDoc.Internal
{
  internal class StructureSetContent
  {
    public void Add(Section section, StructureSet structureSet)
    {
      //AddHeading(section, structureSet);
      AddStructures(section, structureSet.Structures);
    }

    //    private void AddHeading(Section section, StructureSet structureSet)
    //    {
    //  section.AddParagraph(structureSet.ReportParameters, StyleNames.Heading1);
    //  //section.AddParagraph($"Image {structureSet.Image.Id} " +
    //  //                     $"taken {structureSet.Image.CreationTime:g}");
    //  section.AddParagraph("Report Date Range: 04/2019-05/2020");
    //  section.AddParagraph(structureSet.ReportParameters);
    //  section.AddParagraph("Service:All");
    //  section.AddParagraph("----------------------------------------------------------------------------------");

    //}

    private void AddStructures(Section section, List<Structure> structures)
    {
      //AddTableTitle(section, "(SymphonyMedia-Linear) Billing Summary by Agreement");
      AddStructureTable(section, structures);
    }

    private void AddTableTitle(Section section, string title)
    {
      var p = section.AddParagraph(title, StyleNames.Heading2);
      p.Format.KeepWithNext = true;
    }

    private void AddStructureTable(Section section, List<Structure> structures)
    {
      var table = section.AddTable();

      FormatTable(table);
      AddColumnsAndHeaders(table);
      AddStructureRows(table, structures);

      AddLastRowBorder(table);
      AlternateRowShading(table);
    }

    private static void FormatTable(Table table)
    {
      table.LeftPadding = 0;
      table.TopPadding = Size.TableCellPadding;
      table.RightPadding = 0;
      table.BottomPadding = Size.TableCellPadding;
      table.Format.LeftIndent = Size.TableCellPadding;
      table.Format.RightIndent = Size.TableCellPadding;
    }

    private void AddColumnsAndHeaders(Table table)
    {
      var width = Size.GetWidth(table.Section);
      table.AddColumn(width * 0.16);
      table.AddColumn(width * 0.16);
      table.AddColumn(width * 0.16);
      table.AddColumn(width * 0.16);
      table.AddColumn(width * 0.16);
      table.AddColumn(width * 0.16);
      table.AddColumn(width * 0.16);
      table.AddColumn(width * 0.16);
      table.AddColumn(width * 0.16);
      table.AddColumn(width * 0.16);

      var headerRow = table.AddRow();
      headerRow.Borders.Bottom.Width = 1;

      //AddHeader(headerRow.Cells[0], "Structure ID");
      //AddHeader(headerRow.Cells[1], "Volume [cc]");
      //AddHeader(headerRow.Cells[2], "Mean Dose [Gy]");

      //      AddHeader(headerRow.Cells[0], "Invoice Date");
      //      AddHeader(headerRow.Cells[1], "Invoice Number");
      //      AddHeader(headerRow.Cells[2], "Agreement Name");
      //      AddHeader(headerRow.Cells[3], "MVPD");
      //      AddHeader(headerRow.Cells[4], "Service");
      //AddHeader(headerRow.Cells[5], "Subscriber Type");
      //AddHeader(headerRow.Cells[6], "System Designation");
      //AddHeader(headerRow.Cells[7], "Subs");
      //      AddHeader(headerRow.Cells[8], "Rate");
      //      AddHeader(headerRow.Cells[9], "Revenue");
    }

    private void AddHeader(Cell cell, string header)
    {
      var p = cell.AddParagraph(header);
      p.Style = CustomStyles.ColumnHeader;
      cell.Shading.Color = Color.FromRgb(161, 161, 161);
    }

    private void AddStructureRows(Table table, List<Structure> structures)
    {
      List<Structure> st = structures;
      st.OrderBy(x => x.Agreement_Name);
      IEnumerable<string> lstDistictsAgreement;
      string previousAgreement = "";
      int nbrOfRecord = 0,
        temp = 0;
      double grandSubTotal = 0;
      double grandRateTotal = 0;
      double grandRevenueTotal = 0;
      lstDistictsAgreement = st.Select(x => x.Agreement_Name).Distinct();

      //foreach (var structure in structures)
      //      {
      //          var row = table.AddRow();
      //          row.VerticalAlignment = VerticalAlignment.Center;

      //          row.Cells[0].AddParagraph(structure.Invoice_Date);
      //          row.Cells[1].AddParagraph($"{structure.Invoice_Number}");
      //          row.Cells[2].AddParagraph($"{structure.Agreement_Name}");
      //          row.Cells[3].AddParagraph($"{structure.MVPD}");
      //          row.Cells[4].AddParagraph($"{structure.Service}");
      //  row.Cells[5].AddParagraph($"{structure.SubscriberType:f2}");
      //  row.Cells[6].AddParagraph($"{structure.SystemDesignation:f2}");
      //  //row.Cells[7].AddParagraph($"{structure.Subs:f2}");
      //  //row.Cells[8].AddParagraph($"{structure.Rate:f2}");
      //  //row.Cells[9].AddParagraph($"{structure.Revenue:f2}");
      //  row.Cells[7].AddParagraph($"{structure.Subs:f2}");
      //          row.Cells[8].AddParagraph($"{structure.Rate:f2}");
      //          row.Cells[9].AddParagraph($"{structure.Revenue:f2}");
      //      }
      foreach (var item in lstDistictsAgreement)
      {
        nbrOfRecord = 0;
        temp++;
        for (int j = 0; j < st.Count; j++)
        {
          if (st[j].Agreement_Name == item)
          {
            nbrOfRecord++;
            //index = worksheet.Dimension.End.Row + 1;

            if (j == 0 || st[j - 1].Agreement_Name != previousAgreement)
            {
              //row.Cell["A" + (index)].Value = "Invoice Date";
              //worksheet.Cells["" + worksheet.Cells["A" + index].Address + ":" + worksheet.Cells["C" + index].Address + ""].Style.Font.Bold = true;
              //worksheet.Cells["" + worksheet.Cells["A" + index].Address + ":" + worksheet.Cells["K" + index].Address + ""].Style.Border.Top.Style = ExcelBorderStyle.Thin;
              //worksheet.Cells["" + worksheet.Cells["A" + (index + 1)].Address + ":" + worksheet.Cells["K" + (index + 1)].Address + ""].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
              //worksheet.Cells["B" + (index)].Value = "Invoice Number";
              //worksheet.Cells["C" + (index)].Value = "Agreement";
              //worksheet.Cells["D" + (index + 1)].Value = "MVPD";
              //worksheet.Cells["E" + (index + 1)].Value = "Service";
              //worksheet.Cells["F" + (index + 1)].Value = "Subscriber Type";
              //if (row < (codeDetails.Count() - 1))
              //{
              //  if (codeDetails[row].SystemDesignation == "Broadcast" || codeDetails[row].SystemDesignation == "Hybrid")
              //  {
              //    worksheet.Cells["G" + (index + 1)].Value = "System Designation";
              //  }
              //}
              ////worksheet.Cells["G" + (index + 1)].Value = "System Designation";
              //worksheet.Cells["H" + (index + 1)].Value = "";
              //worksheet.Cells["I" + (index + 1)].Value = "Subs";
              //worksheet.Cells["J" + (index + 1)].Value = "Rate";
              //worksheet.Cells["K" + (index + 1)].Value = "Revenue";
              var headerRow = table.AddRow();
              headerRow.Borders.Bottom.Width = 1;
              AddHeader(headerRow.Cells[0], "Invoice Date");
              AddHeader(headerRow.Cells[1], "Invoice Number");
              AddHeader(headerRow.Cells[2], "Agreement Name");
              AddHeader(headerRow.Cells[3], "MVPD");
              AddHeader(headerRow.Cells[4], "Service");
              AddHeader(headerRow.Cells[5], "Subscriber Type");
              AddHeader(headerRow.Cells[6], "System Designation");
              //AddHeader(headerRow.Cells[7], "Subs");
              //AddHeader(headerRow.Cells[8], "Rate");
              //AddHeader(headerRow.Cells[9], "Revenue");
              AddHeader(headerRow.Cells[7], "Subs");
              AddHeader(headerRow.Cells[8], "Rate");
              AddHeader(headerRow.Cells[9], "Revenue");
            }
            if (j == 0)
            {
              if (nbrOfRecord == 1)
              {
                var row1 = table.AddRow();
                row1.VerticalAlignment = VerticalAlignment.Center;

                row1.Cells[0].AddParagraph(st[j].Invoice_Date);
                row1.Cells[1].AddParagraph($"{st[j].Invoice_Number}");
                row1.Cells[2].AddParagraph($"{st[j].Agreement_Name}");
                row1.Cells[3].AddParagraph($"{st[j].MVPD}");
                row1.Cells[4].AddParagraph($"{st[j].Service}");
                row1.Cells[5].AddParagraph($"{st[j].SubscriberType:f2}");
                row1.Cells[6].AddParagraph($"{st[j].SystemDesignation:f2}");
                row1.Cells[7].AddParagraph("$"+$"{st[j].Subs:f2}");
                row1.Cells[8].AddParagraph("$"+$"{st[j].Rate:f2}");
                row1.Cells[9].AddParagraph("$"+$"{st[j].Revenue:f2}");
              }
              else
              {
                //worksheet.Cells["A" + (index + 2)].Value = "";
                //worksheet.Cells["B" + (index + 2)].Value = "";
                //worksheet.Cells["C" + (index + 2)].Value = "";
                //worksheet.Cells["D" + (index + 2)].Value = "";
                //worksheet.Cells["E" + (index + 2)].Value = "";
                //worksheet.Cells["F" + (index + 2)].Value = "";
                //if (codeDetails[row].SystemDesignation == "Broadcast" || codeDetails[row].SystemDesignation == "Hybrid")
                //{
                //  worksheet.Cells["G" + (index + 2)].Value = "";
                //}
                //worksheet.Cells["G" + (index + 2)].Value = "";
                var row1 = table.AddRow();
                row1.VerticalAlignment = VerticalAlignment.Center;
                row1.Cells[0].AddParagraph("");
                row1.Cells[1].AddParagraph("");
                row1.Cells[2].AddParagraph("");
                row1.Cells[3].AddParagraph("");
                row1.Cells[4].AddParagraph("");
                row1.Cells[5].AddParagraph("");
                row1.Cells[6].AddParagraph("");
                row1.Cells[7].AddParagraph("$"+$"{st[j].Subs:f2}");
                row1.Cells[8].AddParagraph("$"+$"{st[j].Rate:f2}");
                row1.Cells[9].AddParagraph("$" + $"{st[j].Revenue:f2}");
              }

            }
            else
            {
              if (st[j - 1].Agreement_Name != previousAgreement)
              {
                if (nbrOfRecord == 1)
                {
                  var row1 = table.AddRow();
                  row1.VerticalAlignment = VerticalAlignment.Center;

                  row1.Cells[0].AddParagraph(st[j].Invoice_Date);
                  row1.Cells[1].AddParagraph($"{st[j].Invoice_Number}");
                  row1.Cells[2].AddParagraph($"{st[j].Agreement_Name}");
                  row1.Cells[3].AddParagraph($"{st[j].MVPD}");
                  row1.Cells[4].AddParagraph($"{st[j].Service}");
                  row1.Cells[5].AddParagraph($"{st[j].SubscriberType:f2}");
                  row1.Cells[6].AddParagraph($"{st[j].SystemDesignation:f2}");
                  row1.Cells[7].AddParagraph("$"+$"{st[j].Subs:f2}");
                  row1.Cells[8].AddParagraph("$"+$"{st[j].Rate:f2}");
                  row1.Cells[9].AddParagraph("$" + $"{st[j].Revenue:f2}");
                }
                else
                {
                  var row1 = table.AddRow();
                  row1.VerticalAlignment = VerticalAlignment.Center;
                  row1.Cells[0].AddParagraph("");
                  row1.Cells[1].AddParagraph("");
                  row1.Cells[2].AddParagraph("");
                  row1.Cells[3].AddParagraph("");
                  row1.Cells[4].AddParagraph("");
                  row1.Cells[5].AddParagraph("");
                  row1.Cells[6].AddParagraph("");
                  row1.Cells[7].AddParagraph("$"+$"{st[j].Subs:f2}");
                  row1.Cells[8].AddParagraph("$"+$"{st[j].Rate:f2}");
                  row1.Cells[9].AddParagraph("$" + $"{st[j].Revenue:f2}");
                }

              }
              else
              {
                if (nbrOfRecord == 1)
                {
                  var row1 = table.AddRow();
                  row1.VerticalAlignment = VerticalAlignment.Center;

                  row1.Cells[0].AddParagraph(st[j].Invoice_Date);
                  row1.Cells[1].AddParagraph($"{st[j].Invoice_Number}");
                  row1.Cells[2].AddParagraph($"{st[j].Agreement_Name}");
                  row1.Cells[3].AddParagraph($"{st[j].MVPD}");
                  row1.Cells[4].AddParagraph($"{st[j].Service}");
                  row1.Cells[5].AddParagraph($"{st[j].SubscriberType:f2}");
                  row1.Cells[6].AddParagraph($"{st[j].SystemDesignation:f2}");
                  row1.Cells[7].AddParagraph("$"+$"{st[j].Subs:f2}");
                  row1.Cells[8].AddParagraph("$"+$"{st[j].Rate:f2}");
                  row1.Cells[9].AddParagraph("$" + $"{st[j].Revenue:f2}");
                }
                else
                {
                  var row1 = table.AddRow();
                  row1.VerticalAlignment = VerticalAlignment.Center;
                  row1.Cells[0].AddParagraph("");
                  row1.Cells[1].AddParagraph("");
                  row1.Cells[2].AddParagraph("");
                  row1.Cells[3].AddParagraph("");
                  row1.Cells[4].AddParagraph("");
                  row1.Cells[5].AddParagraph("");
                  row1.Cells[6].AddParagraph("");
                  row1.Cells[7].AddParagraph("$"+$"{st[j].Subs:f2}");
                  row1.Cells[8].AddParagraph("$"+$"{st[j].Rate:f2}");
                  row1.Cells[9].AddParagraph("$" + $"{st[j].Revenue:f2}");
                }
              }
            }

            //if (j == (st.Count() - 1) || (previousAgreement != item && previousAgreement != ""))
            //{
            //  var row1 = table.AddRow();
            //  row1.VerticalAlignment = VerticalAlignment.Center;
            //  row1.Cells[6].AddParagraph($"Grand Total");

            //  row1.Cells[6].As = "SUM(" + worksheet.Cells["I" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["I" + (index - 2)].Address + ")";
            //  worksheet.Cells["" + worksheet.Cells["I" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["I" + (index - 2)].Address + ""].Style.Numberformat.Format = "$###,###,##0.00";
            //  worksheet.Cells["J" + (index)].Formula = "SUM(" + worksheet.Cells["J" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["J" + (index - 2)].Address + ")";
            //  worksheet.Cells["" + worksheet.Cells["J" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["J" + (index - 2)].Address + ""].Style.Numberformat.Format = "$###,###,##0.00";
            //  worksheet.Cells["K" + (index)].Formula = "SUM(" + worksheet.Cells["K" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["K" + (index - 2)].Address + ")";
            //  worksheet.Cells["" + worksheet.Cells["K" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["K" + (index - 2)].Address + ""].Style.Numberformat.Format = "$###,###,##0.00";
            //  //var abc = worksheet.Cells["I" + (index)].Text;
            //  //value = (int)worksheet.Cells["I" + (index)].Value;
            //  //value1 = value1 + value;
            //  //cellValue = (double)cellValue + (double)(int)worksheet.Cells["I" + (index)].Value;
            //  worksheet.Cells["I" + (index)].Calculate();
            //  worksheet.Cells["J" + (index)].Calculate();
            //  worksheet.Cells["K" + (index)].Calculate();
            //  //dValue = Convert.ToDouble(worksheet.Cells["I" + (index)].Value);
            //  grandSubTotal = grandSubTotal + Convert.ToDouble(worksheet.Cells["I" + (index)].Value);
            //  grandRateTotal = grandRateTotal + Convert.ToDouble(worksheet.Cells["J" + (index)].Value);
            //  grandRevenueTotal = grandRevenueTotal + Convert.ToDouble(worksheet.Cells["K" + (index)].Value);
            //  excelPackage.Workbook.Calculate();
            //}
            //if (row > 0 && row != (codeDetails.Count() - 1))
            //{
            //  if (codeDetails[row + 1].Agreement != previousAgreement)
            //  {
            //    index = worksheet.Dimension.End.Row + 2;
            //    worksheet.Cells["G" + (index)].Value = "(" + item + ") " + "Sub Total:";
            //    worksheet.Cells["" + worksheet.Cells["A" + index].Address + ":" + worksheet.Cells["K" + index].Address + ""].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //    worksheet.Cells["" + worksheet.Cells["G" + index].Address + ":" + worksheet.Cells["K" + index].Address + ""].Style.Font.Bold = true;
            //    worksheet.Cells["" + worksheet.Cells["G" + index].Address + ":" + worksheet.Cells["K" + index].Address + ""].Style.Font.UnderLine = true;
            //    worksheet.Cells["" + worksheet.Cells["A" + index].Address + ":" + worksheet.Cells["K" + index].Address + ""].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


            //    worksheet.Cells["I" + (index)].Formula = "SUM(" + worksheet.Cells["I" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["I" + (index - 2)].Address + ")";
            //    worksheet.Cells["" + worksheet.Cells["I" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["I" + (index - 2)].Address + ""].Style.Numberformat.Format = "$###,###,##0.00";
            //    worksheet.Cells["J" + (index)].Formula = "SUM(" + worksheet.Cells["J" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["J" + (index - 2)].Address + ")";
            //    worksheet.Cells["" + worksheet.Cells["J" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["J" + (index - 2)].Address + ""].Style.Numberformat.Format = "$###,###,##0.00";
            //    worksheet.Cells["K" + (index)].Formula = "SUM(" + worksheet.Cells["K" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["K" + (index - 2)].Address + ")";
            //    worksheet.Cells["" + worksheet.Cells["K" + (index - 1 - nbrOfRecord)].Address + ":" + worksheet.Cells["K" + (index - 2)].Address + ""].Style.Numberformat.Format = "$###,###,##0.00";
            //    //var abc = worksheet.Cells["I15"].Value;
            //    //abc = worksheet.Cells["I15"].Value;
            //    worksheet.Cells["I" + (index)].Calculate();
            //    worksheet.Cells["J" + (index)].Calculate();
            //    worksheet.Cells["K" + (index)].Calculate();
            //    //string a = worksheet.Cells["I15"].Value;
            //    //worksheet.Cells["I60"].Value = worksheet.Cells["I15"].Value;
            //    //cellValue = (double)cellValue + (double)(int)worksheet.Cells["I" + (index)].Value;
            //    //subValue = Convert.ToDouble(worksheet.Cells["I" + (index)].Value);
            //    grandSubTotal = grandSubTotal + Convert.ToDouble(worksheet.Cells["I" + (index)].Value);
            //    grandRateTotal = grandRateTotal + Convert.ToDouble(worksheet.Cells["J" + (index)].Value);
            //    grandRevenueTotal = grandRevenueTotal + Convert.ToDouble(worksheet.Cells["K" + (index)].Value);
            //    excelPackage.Workbook.Calculate();
            //  }
            //  worksheet.Cells["A15"].Calculate();
            //}

          }
          previousAgreement = item;
          if (j == (st.Count() - 1))
          {
            var row1 = table.AddRow();
            row1.VerticalAlignment = VerticalAlignment.Center;
            row1.Shading.Color = Color.FromRgb(216, 216, 216);
            row1.Cells[0].AddParagraph("");
            row1.Cells[1].AddParagraph("");
            row1.Cells[2].AddParagraph("");
            row1.Cells[3].AddParagraph("");
            row1.Cells[4].AddParagraph("");
            row1.Cells[5].AddParagraph("");
            row1.Cells[6].AddParagraph($"Grand Total");
            row1.Cells[7].AddParagraph("$" + CalculateSubs(item, st, "Subs"));
            row1.Cells[8].AddParagraph("$" + CalculateSubs(item, st, "Rate"));
            row1.Cells[9].AddParagraph("$" + CalculateSubs(item, st, "Revenue"));
          }
          //if (row == (codeDetails.Count() - 1) && temp == lstDistictsAgreement.Count())
          //{
          //  index = worksheet.Dimension.End.Row + 2;
          //  worksheet.Cells["G" + index].Value = "Grand Total";
          //  worksheet.Cells["" + worksheet.Cells["A" + index].Address + ":" + worksheet.Cells["K" + index].Address + ""].Style.Border.Top.Style = ExcelBorderStyle.Thin;
          //  worksheet.Cells["" + worksheet.Cells["G" + index].Address + ":" + worksheet.Cells["K" + index].Address + ""].Style.Font.Bold = true;
          //  worksheet.Cells["" + worksheet.Cells["G" + index].Address + ":" + worksheet.Cells["K" + index].Address + ""].Style.Font.UnderLine = true;
          //  worksheet.Cells["" + worksheet.Cells["A" + index].Address + ":" + worksheet.Cells["K" + index].Address + ""].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

          //  worksheet.Cells["I" + index].Formula = grandSubTotal.ToString();
          //  //worksheet.Cells["" + worksheet.Cells["I" + index].Address + ":" + worksheet.Cells["I" + index].Address + ""].Style.Numberformat.Format = "$###,###,##0.00";
          //  worksheet.Cells["J" + index].Formula = grandRateTotal.ToString();
          //  //worksheet.Cells["" + worksheet.Cells["J" + index].Address + ":" + worksheet.Cells["J" + index].Address + ""].Style.Numberformat.Format = "$###,###,##0.00";
          //  worksheet.Cells["K" + index].Formula = grandRevenueTotal.ToString();
          //  //worksheet.Cells["" + worksheet.Cells["K" + index].Address + ":" + worksheet.Cells["K" + index].Address + ""].Style.Numberformat.Format = "$###,###,##0.00";
          //  worksheet.Cells["K" + (index + 2)].Value = "Client Short Name";
          //  worksheet.Cells["K" + (index + 2)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

          //  excelPackage.Workbook.Calculate();
          //}
        }

      }
    }
    private double CalculateSubs(string currentItem, List<Structure> currentStruct, string columnName)
    {
      double totalSubs = 0;
      switch (columnName)
      {
        case "Subs":
          {
            totalSubs = currentStruct.Where(obj => obj.Agreement_Name == currentItem).Sum(obj => obj.Subs);

            break;
          }
        case "Rate":
          {
            totalSubs = currentStruct.Where(obj => obj.Agreement_Name == currentItem).Sum(obj => obj.Rate);
            break;
          }
        case "Revenue":
          {
            totalSubs = currentStruct.Where(obj => obj.Agreement_Name == currentItem).Sum(obj => obj.Revenue);
            break;
          }

      }
      return totalSubs;
    }
    private void AddLastRowBorder(Table table)
    {
      var lastRow = table.Rows[table.Rows.Count - 1];
      lastRow.Borders.Bottom.Width = 2;
    }

    private void AlternateRowShading(Table table)
    {
      // Start at i = 1 to skip column headers
      for (var i = 1; i < table.Rows.Count; i++)
      {
        //if (i % 2 == 0)  // Even rows
        //{
        //    table.Rows[i].Shading.Color = Color.FromRgb(216, 216, 216);
        //}
        //else if (i == 1)
        //{
        //  table.Rows[i].Shading.Color = Color.FromRgb(216, 216, 216);

        //}
      }
    }

  }
}