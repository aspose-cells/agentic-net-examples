using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsNumberFormatDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Apply a built‑in number format using Style.Number
            // -------------------------------------------------
            // Access cell A1 and put a numeric value
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue(1234.567);

            // Retrieve the cell style, set the built‑in format (5 = Currency with 2 decimals)
            Style styleA1 = cellA1.GetStyle();
            styleA1.Number = 5;               // "$#,##0_);($#,##0)" format
            cellA1.SetStyle(styleA1);

            // -------------------------------------------------
            // 2. Apply a custom number format using StyleFlag
            // -------------------------------------------------
            // Access cell B2 and put a numeric value
            Cell cellB2 = sheet.Cells["B2"];
            cellB2.PutValue(9876.543);

            // Create a custom style with a Euro currency format
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "_-€ #,##0.00;[Red]_-€ -#,##0.00";

            // Use StyleFlag to apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to cell B2
            Aspose.Cells.Range rangeB2 = sheet.Cells.CreateRange(1, 1, 1, 1);
            rangeB2.ApplyStyle(customStyle, flag);

            // -------------------------------------------------
            // 3. Set number format for a chart's data labels
            // -------------------------------------------------
            // Add sample data for the chart
            sheet.Cells["A4"].PutValue("Category");
            sheet.Cells["A5"].PutValue("A");
            sheet.Cells["A6"].PutValue("B");
            sheet.Cells["A7"].PutValue("C");
            sheet.Cells["B4"].PutValue("Value");
            sheet.Cells["B5"].PutValue(1500);
            sheet.Cells["B6"].PutValue(2500);
            sheet.Cells["B7"].PutValue(3500);

            // Insert a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 9, 0, 25, 15);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B5:B7", true);
            chart.NSeries.CategoryData = "A5:A7";

            // Enable data labels and set a custom number format
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.NumberFormat = "$#,##0";

            // -------------------------------------------------
            // Save the workbook to demonstrate all applied formats
            // -------------------------------------------------
            workbook.Save("NumberFormatDemo.xlsx");

            Console.WriteLine("Workbook created with various number formats.");
        }
    }
}