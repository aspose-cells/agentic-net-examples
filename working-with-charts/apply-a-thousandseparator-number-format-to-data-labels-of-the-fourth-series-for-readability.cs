// Title: How to apply a thousand‑separator number format to the data labels of the fourth series in an Aspose.Cells column chart using C#
// AI Prompts: Write C# that turns on data labels for the fourth series of an Aspose.Cells column chart and formats the values with a comma‑grouping pattern (#,##0). | Demonstrate how to assign a number format string to the DataLabels of a specific series in an Aspose.Cells chart using .NET.
// Common Searches: Aspose.Cells .NET how to show data labels with thousand separator for a chart series | C# set number format '#,##0' for fourth series data labels in column chart | Apply custom numeric format to chart series labels using Aspose.Cells | Formatting chart data labels with commas in Aspose.Cells C#
// Tags: Aspose.Cells column chart series label formatting | C# configure chart label numeric style | Aspose.Cells apply custom number format to fourth series | Aspose.Cells chart series label styling | Aspose.Cells column chart label appearance

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook with four data series, adds a column chart, enables data labels for the fourth series, and sets the label number format to "#,##0" so values appear with thousand separators before saving the file.
    public class ApplyThousandSeparatorToFourthSeries
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
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

            // Populate sample data for four series
            // Column A – Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");

            // Columns B‑E – Values for four different series
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(2500);
            sheet.Cells["B4"].PutValue(3700);
            sheet.Cells["B5"].PutValue(4800);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(1500);
            sheet.Cells["C3"].PutValue(2600);
            sheet.Cells["C4"].PutValue(3900);
            sheet.Cells["C5"].PutValue(5000);

            sheet.Cells["D1"].PutValue("Series3");
            sheet.Cells["D2"].PutValue(1800);
            sheet.Cells["D3"].PutValue(2700);
            sheet.Cells["D4"].PutValue(4100);
            sheet.Cells["D5"].PutValue(5200);

            sheet.Cells["E1"].PutValue("Series4");
            sheet.Cells["E2"].PutValue(2000);
            sheet.Cells["E3"].PutValue(3000);
            sheet.Cells["E4"].PutValue(4500);
            sheet.Cells["E5"].PutValue(6000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the four series to the chart
            chart.NSeries.Add("B2:B5", true); // Series1
            chart.NSeries.Add("C2:C5", true); // Series2
            chart.NSeries.Add("D2:D5", true); // Series3
            chart.NSeries.Add("E2:E5", true); // Series4

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the fourth series and apply thousand‑separator format
            Series fourthSeries = chart.NSeries[3]; // zero‑based index
            fourthSeries.DataLabels.ShowValue = true;
            fourthSeries.DataLabels.NumberFormat = "#,##0"; // thousand separator

            // Save the workbook
            string outputPath = "ThousandSeparatorFourthSeries.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                throw;
            }
        }
    }
}
