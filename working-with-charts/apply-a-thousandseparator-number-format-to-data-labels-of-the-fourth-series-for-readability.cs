// Title: C# – Apply Thousand‑Separator Format to Data Labels of the Fourth Series in an Aspose.Cells Column Chart
// Description: Creates a workbook, fills categories and four numeric series, adds a column chart, enables data labels only for the fourth series, applies the "#,##0" thousand‑separator format to those labels, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells for .NET | C# chart formatting | column chart data labels | thousand separator number format | fourth series label formatting | Aspose.Cells number format example | Excel chart custom labels | Aspose.Cells chart series styling
// Common Searches: Aspose.Cells set thousand separator on specific chart series | C# format data labels of fourth series in column chart | How to apply custom number format to chart labels using Aspose.Cells | Enable data labels for one series only Aspose.Cells .NET | Aspose.Cells column chart label formatting example
// Developer Intent: Format the data labels of the fourth series in a column chart with a thousand‑separator number format using Aspose.Cells for .NET.
// Use Cases: Present large values with commas in the last series of a sales chart for clearer financial reports. | Generate Excel dashboards where only the final data series requires a custom numeric display. | Create mixed‑format charts that keep default labels for earlier series while highlighting the fourth series with "#,##0" formatting.
// AI Prompts: Show C# code that enables data labels for the fourth series of an Aspose.Cells column chart and sets the number format to "#,##0". | Explain step‑by‑step how to apply a thousand‑separator format to chart series data labels with Aspose.Cells for .NET. | Provide a concise example of creating a workbook, adding a column chart, and formatting only the fourth series' data labels using a custom number format.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills categories and four numeric series, adds a column chart, enables data labels only for the fourth series, applies the "#,##0" thousand‑separator format to those labels, and saves the file as an Excel workbook.
    public class ApplyThousandSeparatorToFourthSeries
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for four series (columns B to E) with categories in column A
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
                sheet.Cells["C2"].PutValue(2200);
                sheet.Cells["C3"].PutValue(2500);
                sheet.Cells["C4"].PutValue(2800);
                sheet.Cells["C5"].PutValue(3100);

                // Series 3
                sheet.Cells["D1"].PutValue("Series3");
                sheet.Cells["D2"].PutValue(3200);
                sheet.Cells["D3"].PutValue(3500);
                sheet.Cells["D4"].PutValue(3800);
                sheet.Cells["D5"].PutValue(4100);

                // Series 4 (the target series)
                sheet.Cells["E1"].PutValue("Series4");
                sheet.Cells["E2"].PutValue(4200);
                sheet.Cells["E3"].PutValue(4500);
                sheet.Cells["E4"].PutValue(4800);
                sheet.Cells["E5"].PutValue(5100);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add all four series to the chart (range B2:E5) and set categories (A2:A5)
                chart.NSeries.Add("B2:E5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Enable data labels for the fourth series (index 3)
                Series fourthSeries = chart.NSeries[3];
                fourthSeries.DataLabels.ShowValue = true;

                // Apply thousand‑separator format to the data labels of the fourth series
                fourthSeries.DataLabels.NumberFormat = "#,##0";

                // Save the workbook
                workbook.Save("ThousandSeparatorFourthSeries.xlsx");
                Console.WriteLine("Workbook saved successfully.");
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
            ApplyThousandSeparatorToFourthSeries.Run();
        }
    }
}
