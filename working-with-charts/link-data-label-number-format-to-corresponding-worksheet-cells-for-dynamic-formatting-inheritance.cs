using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDynamicDataLabelFormatting
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1500);
            sheet.Cells["B3"].PutValue(2500);
            // Cells that contain formatted text for data labels
            sheet.Cells["C1"].PutValue("Formatted Value");
            sheet.Cells["C2"].PutValue("1.5k USD");
            sheet.Cells["C3"].PutValue("2.5k USD");

            // Apply a custom number format to the source cells (C2:C3)
            Style style = workbook.CreateStyle();
            style.Custom = "\"$\"#,##0\" USD\"";
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;
            sheet.Cells.CreateRange("C2:C3").ApplyStyle(style, flag);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set data range for the series
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Configure data labels to use the formatted cells and link number format
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // Show the numeric value
            series.DataLabels.LinkedSource = "C2:C3";          // Link to cells with custom text
            series.DataLabels.NumberFormatLinked = true;      // Inherit number format from linked cells

            // OPTIONAL: Change the number format of the linked cells after linking
            // This demonstrates that data label formatting updates dynamically.
            sheet.Cells["C2"].SetStyle(style);
            sheet.Cells["C3"].SetStyle(style);

            // Save the workbook
            workbook.Save("DynamicDataLabelFormatting.xlsx");
        }
    }
}