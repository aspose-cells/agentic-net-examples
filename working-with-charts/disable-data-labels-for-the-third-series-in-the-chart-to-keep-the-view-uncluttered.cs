// Title: Aspose.Cells C# – Disable Data Labels for the Third Series in a Column Chart
// Description: Learn how to create a workbook with three data series, add a column chart, enable data labels for the first two series, and hide all data labels for the third series using Aspose.Cells for .NET. The example saves the result as an XLSX file.
// Keywords: Aspose.Cells C# data labels | disable chart series labels | hide third series labels Aspose.Cells | column chart data label visibility | Aspose.Cells chart customization | C# Excel chart series label control
// Common Searches: Aspose.Cells hide data labels third series C# | disable specific series labels in Aspose.Cells chart | C# Aspose.Cells column chart label settings | turn off data labels for one series Aspose.Cells | Aspose.Cells chart series label visibility
// Developer Intent: The developer wants to suppress all data labels for the third series of a column chart while keeping labels visible for the other series, using Aspose.Cells in C#.
// Use Cases: Create a sales dashboard where only the primary product lines show values on the chart, keeping a secondary line label‑free for clarity. | Generate an Excel report that highlights two key metrics with data labels and omits labels for a comparison metric to reduce visual clutter. | Export a workbook with a column chart that displays detailed labels for selected series while hiding them for others to meet presentation standards.
// AI Prompts: Write C# code with Aspose.Cells that adds a column chart containing three series and disables data labels for the third series. | Explain which Aspose.Cells properties control data label visibility per series and how to configure them to hide labels for a specific series. | Show how to toggle data label visibility on and off for individual series in an existing Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Learn how to create a workbook with three data series, add a column chart, enable data labels for the first two series, and hide all data labels for the third series using Aspose.Cells for .NET. The example saves the result as an XLSX file.
    public class DisableThirdSeriesDataLabels
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for three series
                // Category labels
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                // Series 1 values
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Series 2 values
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Series 3 values
                sheet.Cells["D1"].PutValue("Series 3");
                sheet.Cells["D2"].PutValue(12);
                sheet.Cells["D3"].PutValue(22);
                sheet.Cells["D4"].PutValue(32);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add the three series to the chart
                chart.NSeries.Add("=Sheet1!$B$2:$B$4", true); // Series 1
                chart.NSeries.Add("=Sheet1!$C$2:$C$4", true); // Series 2
                chart.NSeries.Add("=Sheet1!$D$2:$D$4", true); // Series 3

                // Set category (X) data
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

                // Enable data labels for the first two series (optional, to show the effect)
                chart.NSeries[0].DataLabels.ShowValue = true;
                chart.NSeries[1].DataLabels.ShowValue = true;

                // Disable all data labels for the third series
                Series thirdSeries = chart.NSeries[2];
                thirdSeries.DataLabels.ShowValue = false;          // hide values
                thirdSeries.DataLabels.ShowCategoryName = false;   // hide category names
                thirdSeries.DataLabels.ShowPercentage = false;    // hide percentages (if applicable)
                thirdSeries.DataLabels.ShowSeriesName = false;    // hide series name

                // Save the workbook
                string outputPath = "ChartWithThirdSeriesLabelsDisabled.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
