// Title: Aspose.Cells C# – Hide Data Labels for the Third Series in a Column Chart
// Description: Demonstrates how to create a workbook with a column chart containing three series, enable data labels for the first two series, and completely hide all label types (value, category name, percentage, series name) for the third series before saving the file as an XLSX document.
// Keywords: Aspose.Cells hide data labels | C# chart series label control | disable third series labels Aspose.Cells | column chart data labels .NET | Aspose.Cells chart customization
// Common Searches: Aspose.Cells hide data labels for one series | C# disable data labels third series column chart | remove specific series labels Aspose.Cells | how to turn off chart labels for a single series in .NET | Aspose.Cells chart label settings example
// Developer Intent: The developer needs to suppress all data‑label displays for the third series of a column chart while keeping labels visible for the other series.
// Use Cases: Generate an Excel chart where only key series show values, keeping auxiliary series label‑free for a cleaner visual. | Create a report with a reference line (third series) that should not display any label information. | Produce a presentation‑ready chart that highlights primary metrics by showing labels only on selected series.
// AI Prompts: Write C# code using Aspose.Cells to hide value, category name, percentage, and series name labels for a specific chart series. | Explain how to toggle individual data‑label properties for a series in an Aspose.Cells chart. | Provide an example that adds a column chart with three series and enables labels only for the first two series using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with a column chart containing three series, enable data labels for the first two series, and completely hide all label types (value, category name, percentage, series name) for the third series before saving the file as an XLSX document.
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
                thirdSeries.DataLabels.ShowValue = false;
                thirdSeries.DataLabels.ShowCategoryName = false;
                thirdSeries.DataLabels.ShowPercentage = false;
                thirdSeries.DataLabels.ShowSeriesName = false;

                // Save the workbook
                string outputPath = "ChartWithThirdSeriesLabelsDisabled.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableThirdSeriesDataLabels.Run();
        }
    }
}
