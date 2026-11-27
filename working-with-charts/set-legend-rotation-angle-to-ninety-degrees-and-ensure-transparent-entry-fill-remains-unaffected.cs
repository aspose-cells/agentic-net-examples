// Title: How to rotate a chart legend to 90° and keep legend entries transparent with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart legend rotation to 90 degrees | make legend entry background transparent in Aspose.Cells chart | preserve legend entry fill while rotating legend in a .NET workbook | Aspose.Cells example using LegendEntry.IsTextNoFill | rotate legend text without affecting fill in Aspose.Cells
// Tags: chart legend rotation Aspose.Cells | legend entry transparent fill .NET | set legend rotation angle C# | legendentry istextnofill property | column chart legend customization Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace LegendRotationDemo
{
    // The sample creates a workbook, adds a column chart, sets the legend's RotationAngle to 90 degrees, and uses LegendEntry.IsTextNoFill to make each legend entry background transparent, then saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(50);
                sheet.Cells["B4"].PutValue(70);

                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(40);
                sheet.Cells["C3"].PutValue(60);
                sheet.Cells["C4"].PutValue(80);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // ------------------------------
                // Set legend rotation angle to 90 degrees
                // ------------------------------
                // The Legend class provides a RotationAngle property (in degrees).
                // Setting it to 90 rotates the legend text accordingly.
                chart.Legend.RotationAngle = 90;

                // ------------------------------
                // Ensure each legend entry has transparent (no) fill
                // ------------------------------
                // LegendEntry.IsTextNoFill = true makes the text background transparent.
                // Iterate through all series to apply the setting.
                foreach (Series series in chart.NSeries)
                {
                    LegendEntry entry = series.LegendEntry;
                    entry.IsTextNoFill = true; // Transparent fill for the legend entry text
                }

                // Optional: verify the settings by outputting to console
                Console.WriteLine($"Legend rotation angle set to: {chart.Legend.RotationAngle} degrees");
                foreach (Series series in chart.NSeries)
                {
                    Console.WriteLine($"Series '{series.Name}' legend entry IsTextNoFill = {series.LegendEntry.IsTextNoFill}");
                }

                // Save the workbook
                string outputPath = "LegendRotationTransparentFillDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
