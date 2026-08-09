// Title: C# – Apply Thousand‑Separator Format to Data Labels of the Fourth Series in an Aspose.Cells Column Chart
// Description: Shows how to build a workbook with four data series, insert a column chart, enable data labels for the fourth series, and format those labels with the "#,##0" thousand‑separator pattern using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | column chart | data labels | thousand separator | number format | .NET chart formatting | fourth series | Excel automation | chart customization
// Common Searches: Aspose.Cells set data label number format | C# chart thousand separator Aspose | format specific series labels Aspose.Cells | apply #,##0 to chart labels .NET | how to add commas to chart data labels C#
// Developer Intent: The developer needs the values of the fourth series in a column chart to appear with commas (thousand separators) in the data labels.
// Use Cases: Quarterly sales report where the last series shows revenue and requires comma separators for readability. | Financial dashboard highlighting the final data series with a clear thousand‑separator format. | Performance chart that emphasizes large numbers in the fourth series by applying a custom number pattern.
// AI Prompts: Generate C# code with Aspose.Cells to set a currency format on data labels of the second series in a line chart. | Explain how to change chart data label formats conditionally based on value thresholds using Aspose.Cells. | Provide steps to enable data labels for a specific series and apply the "#,##0" pattern in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook with four data series, insert a column chart, enable data labels for the fourth series, and format those labels with the "#,##0" thousand‑separator pattern using Aspose.Cells for .NET.
    public class ApplyThousandSeparatorToFourthSeries
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for four series (B, C, D, E) with categories in column A
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Series 1
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["B4"].PutValue(1800);
                sheet.Cells["B5"].PutValue(2100);

                // Series 2
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(800);
                sheet.Cells["C3"].PutValue(950);
                sheet.Cells["C4"].PutValue(1100);
                sheet.Cells["C5"].PutValue(1300);

                // Series 3
                sheet.Cells["D1"].PutValue("Series3");
                sheet.Cells["D2"].PutValue(500);
                sheet.Cells["D3"].PutValue(700);
                sheet.Cells["D4"].PutValue(900);
                sheet.Cells["D5"].PutValue(1100);

                // Series 4 (the target series)
                sheet.Cells["E1"].PutValue("Series4");
                sheet.Cells["E2"].PutValue(2500);
                sheet.Cells["E3"].PutValue(3000);
                sheet.Cells["E4"].PutValue(3500);
                sheet.Cells["E5"].PutValue(4000);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the four series to the chart
                chart.NSeries.Add("B2:B5", true); // Series 1
                chart.NSeries.Add("C2:C5", true); // Series 2
                chart.NSeries.Add("D2:D5", true); // Series 3
                chart.NSeries.Add("E2:E5", true); // Series 4

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A5";

                // Enable data labels for the fourth series
                Series fourthSeries = chart.NSeries[3]; // zero‑based index
                fourthSeries.DataLabels.ShowValue = true;

                // Apply thousand‑separator number format to the data labels of the fourth series
                fourthSeries.DataLabels.NumberFormat = "#,##0";

                // Save the workbook
                workbook.Save("ThousandSeparatorFourthSeries.xlsx");
                Console.WriteLine("Workbook saved as ThousandSeparatorFourthSeries.xlsx");
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
            ApplyThousandSeparatorToFourthSeries.Run();
        }
    }
}
