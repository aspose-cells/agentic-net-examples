// Title: Aspose.Cells C# – Apply Scientific Notation to Data Labels of the Fourth Series in a Column Chart
// Description: Creates an Excel workbook with four high‑value series, adds a column chart, enables data labels only for the fourth series, and formats those labels with the "0.00E+00" pattern so numbers appear in scientific notation before saving the file.
// Keywords: Aspose.Cells C# chart data label format | scientific notation Excel chart | format fourth series data labels | column chart number format Aspose | .NET Excel chart customization | high magnitude values display
// Common Searches: Aspose.Cells set scientific notation for specific series | C# column chart data label format example | how to apply number format to chart series in Aspose.Cells | display large numbers as 0.00E+00 in Excel chart using .NET
// Developer Intent: Format the data labels of the fourth series in a column chart to use scientific notation.
// Use Cases: Present financial or engineering data with values in the millions using concise scientific notation for a single series. | Highlight a particular series in multi‑series charts while keeping other series in default formatting. | Automate generation of Excel reports where only one series requires exponential display for readability.
// AI Prompts: Generate C# code that applies the "0.00E+00" number format to the data labels of the fourth series in an Aspose.Cells column chart. | Show how to enable data labels for a specific series and format them as scientific notation using Aspose.Cells for .NET. | Explain the steps to create a chart, activate data labels for one series, set a scientific notation pattern, and save the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates an Excel workbook with four high‑value series, adds a column chart, enables data labels only for the fourth series, and formats those labels with the "0.00E+00" pattern so numbers appear in scientific notation before saving the file.
    public class ScientificNotationDataLabelsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with high magnitude values for four series
                // Category labels
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Series 1
                sheet.Cells["B2"].PutValue(1_200_000);
                sheet.Cells["B3"].PutValue(1_500_000);
                sheet.Cells["B4"].PutValue(1_800_000);
                sheet.Cells["B5"].PutValue(2_100_000);

                // Series 2
                sheet.Cells["C2"].PutValue(2_300_000);
                sheet.Cells["C3"].PutValue(2_600_000);
                sheet.Cells["C4"].PutValue(2_900_000);
                sheet.Cells["C5"].PutValue(3_200_000);

                // Series 3
                sheet.Cells["D2"].PutValue(3_400_000);
                sheet.Cells["D3"].PutValue(3_700_000);
                sheet.Cells["D4"].PutValue(4_000_000);
                sheet.Cells["D5"].PutValue(4_300_000);

                // Series 4 (the target series)
                sheet.Cells["E2"].PutValue(5_500_000);
                sheet.Cells["E3"].PutValue(5_800_000);
                sheet.Cells["E4"].PutValue(6_100_000);
                sheet.Cells["E5"].PutValue(6_400_000);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add each series to the chart
                chart.NSeries.Add("B2:B5", true); // Series 1
                chart.NSeries.Add("C2:C5", true); // Series 2
                chart.NSeries.Add("D2:D5", true); // Series 3
                chart.NSeries.Add("E2:E5", true); // Series 4

                // Set common category data for all series
                chart.NSeries.CategoryData = "A2:A5";

                // Enable data labels for the fourth series (index 3)
                Series fourthSeries = chart.NSeries[3];
                fourthSeries.DataLabels.ShowValue = true;

                // Apply scientific notation format to the data labels of the fourth series
                // Format string "0.00E+00" displays numbers like 5.50E+06
                fourthSeries.DataLabels.NumberFormat = "0.00E+00";

                // Save the workbook
                workbook.Save("ScientificNotationDataLabelsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ScientificNotationDataLabelsDemo.Run();
        }
    }
}
