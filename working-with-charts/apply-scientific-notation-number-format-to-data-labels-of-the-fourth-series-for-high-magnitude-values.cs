// Title: Aspose.Cells C# – Apply Scientific Notation to Data Labels of the Fourth Series in a Column Chart
// Description: Creates a workbook with four data series, adds a column chart, enables data labels only for the fourth series, and formats those labels with the scientific notation pattern "0.00E+00" (e.g., 1.23E+06) before saving the Excel file.
// Keywords: Aspose.Cells | C# | .NET | scientific notation | data labels | chart series formatting | NumberFormat property | column chart | high magnitude values | Excel automation
// Common Searches: Aspose.Cells set scientific notation for chart data labels | C# format specific series data labels in Aspose.Cells | NumberFormat for fourth series in column chart Aspose.Cells | apply custom number format to chart labels .NET | display large values as 1.23E+06 in Aspose.Cells chart
// Developer Intent: Format the data labels of the fourth chart series to show values in scientific notation while leaving other series unchanged.
// Use Cases: Present large financial or scientific figures in a column chart with scientific notation for a single series. | Generate Excel reports where only the last series uses a custom number format. | Automate workbook creation with mixed formatting for chart series using Aspose.Cells.
// AI Prompts: Give C# code that sets scientific notation (0.00E+00) for the data labels of the fourth series in an Aspose.Cells column chart. | Explain how to use Series.DataLabels.NumberFormat to display values like 1.23E+06 while keeping other series default. | Show an example of enabling data labels for a specific series and applying a custom format string in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook with four data series, adds a column chart, enables data labels only for the fourth series, and formats those labels with the scientific notation pattern "0.00E+00" (e.g., 1.23E+06) before saving the Excel file.
    public class ScientificNotationDataLabelsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for four series
                // Categories (X-axis)
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Q1");
                worksheet.Cells["A3"].PutValue("Q2");
                worksheet.Cells["A4"].PutValue("Q3");
                worksheet.Cells["A5"].PutValue("Q4");

                // Series 1
                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(500);
                worksheet.Cells["B3"].PutValue(600);
                worksheet.Cells["B4"].PutValue(700);
                worksheet.Cells["B5"].PutValue(800);

                // Series 2
                worksheet.Cells["C1"].PutValue("Series2");
                worksheet.Cells["C2"].PutValue(1500);
                worksheet.Cells["C3"].PutValue(1600);
                worksheet.Cells["C4"].PutValue(1700);
                worksheet.Cells["C5"].PutValue(1800);

                // Series 3
                worksheet.Cells["D1"].PutValue("Series3");
                worksheet.Cells["D2"].PutValue(2500);
                worksheet.Cells["D3"].PutValue(2600);
                worksheet.Cells["D4"].PutValue(2700);
                worksheet.Cells["D5"].PutValue(2800);

                // Series 4 (high magnitude values)
                worksheet.Cells["E1"].PutValue("Series4");
                worksheet.Cells["E2"].PutValue(1.23e6);
                worksheet.Cells["E3"].PutValue(2.34e6);
                worksheet.Cells["E4"].PutValue(3.45e6);
                worksheet.Cells["E5"].PutValue(4.56e6);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Add all four series to the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.Add("C2:C5", true);
                chart.NSeries.Add("D2:D5", true);
                chart.NSeries.Add("E2:E5", true);

                // Set category (X) data
                chart.NSeries.CategoryData = "A2:A5";

                // Enable data labels for the fourth series (index 3)
                Series fourthSeries = chart.NSeries[3];
                fourthSeries.DataLabels.ShowValue = true;

                // Apply scientific notation format to the data labels of the fourth series
                // Format string "0.00E+00" displays numbers like 1.23E+06
                fourthSeries.DataLabels.NumberFormat = "0.00E+00";

                // Save the workbook
                string outputPath = "ScientificNotationDataLabelsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
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
