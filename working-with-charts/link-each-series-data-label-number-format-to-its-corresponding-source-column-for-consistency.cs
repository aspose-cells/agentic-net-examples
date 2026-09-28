// Title: How to bind each column chart series data label to its own formatted cell range in Aspose.Cells for .NET
// AI Prompts: Write C# code that sets DataLabels.NumberFormatLinked = true and assigns DataLabels.LinkedSource to a custom range for each series in an Aspose.Cells column chart. | Show how to display data label values and apply a dark blue font after linking them to formatted source cells using Aspose.Cells. | Create a workbook with raw and formatted columns, add two series, and programmatically associate each series with its corresponding formatted column for data label formatting.
// Common Searches: how to bind chart series labels to formatted cells using Aspose.Cells C# | setting DataLabels.LinkedSource for each series in a column chart Aspose.Cells | changing font color of data labels after linking number format in Aspose.Cells | sample code for using separate formatted columns for chart data labels in Aspose.Cells
// Tags: chart series data label formatting Aspose.Cells | assign formatted cell range to DataLabels in C# | custom font styling for Aspose.Cells chart labels | column chart with separate formatted source columns | bind series to formatted source cells Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    // The example creates a workbook, populates raw numeric columns and corresponding formatted text columns, adds a column chart with two series, enables NumberFormatLinked for each series, links each series' data labels to its own formatted cell range (C or E), shows the values, applies a dark blue font to the labels, and saves the file as an .xlsx workbook.
    public class LinkDataLabelNumberFormatDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data:
            // Column A – Category names
            // Column B – Raw numeric values for Series 1
            // Column C – Formatted values for Series 1 (e.g., with units)
            // Column D – Raw numeric values for Series 2
            // Column E – Formatted values for Series 2
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            // Series 1 data
            sheet.Cells["B1"].PutValue("Value1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);
            sheet.Cells["C1"].PutValue("Formatted1");
            sheet.Cells["C2"].PutValue("100 units");
            sheet.Cells["C3"].PutValue("200 units");
            sheet.Cells["C4"].PutValue("300 units");

            // Series 2 data
            sheet.Cells["D1"].PutValue("Value2");
            sheet.Cells["D2"].PutValue(400);
            sheet.Cells["D3"].PutValue(500);
            sheet.Cells["D4"].PutValue(600);
            sheet.Cells["E1"].PutValue("Formatted2");
            sheet.Cells["E2"].PutValue("400 pcs");
            sheet.Cells["E3"].PutValue("500 pcs");
            sheet.Cells["E4"].PutValue("600 pcs");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add two series: first uses B column, second uses D column
            chart.NSeries.Add("B2:B4", true); // Series 0
            chart.NSeries.Add("D2:D4", true); // Series 1
            chart.NSeries.CategoryData = "A2:A4";

            // Loop through each series and link its data label number format
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];
                series.DataLabels.ShowValue = true;               // Show the value
                series.DataLabels.NumberFormatLinked = true;      // Link number format to cells

                // Determine the formatted source column based on series index
                // Series 0 -> formatted values in column C, Series 1 -> column E
                string formattedRange = i == 0 ? "C2:C4" : "E2:E4";
                series.DataLabels.LinkedSource = formattedRange;   // Link to formatted cells
            }

            // Optional: style data labels (e.g., font color)
            foreach (Series s in chart.NSeries)
            {
                s.DataLabels.Font.Color = Color.DarkBlue;
                s.DataLabels.ApplyFont(); // Apply font to all child labels
            }

            // Save the workbook
            workbook.Save("LinkDataLabelNumberFormatDemo.xlsx");
        }
    }
}
