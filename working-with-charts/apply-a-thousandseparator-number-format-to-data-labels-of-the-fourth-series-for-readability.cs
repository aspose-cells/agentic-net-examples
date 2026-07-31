// Title: Aspose.Cells C# – Apply Thousand‑Separator Format to Data Labels of the Fourth Series in a Column Chart
// Description: Creates a workbook with four data series, adds a column chart, enables data labels only for the fourth series, and formats those labels with the "#,##0" pattern so values appear with commas as thousand separators. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells C# chart data label format | thousand separator column chart | format fourth series labels Aspose.Cells | custom number format chart series .NET | Excel chart label commas Aspose
// Common Searches: Aspose.Cells set comma separator for chart data labels | C# format specific series data labels in Excel chart | apply "#,##0" to fourth series in Aspose.Cells chart | how to show thousand separators on chart labels .NET | format chart series labels with commas using Aspose
// Developer Intent: Show large numbers in the fourth series of a column chart with commas for better readability.
// Use Cases: Generate financial reports where only one series needs comma‑separated values in the chart. | Create presentation‑ready Excel files with highlighted data labels for a specific series. | Automate Excel chart styling in .NET applications, applying custom number formats to selected series.
// AI Prompts: Write C# code with Aspose.Cells to add a column chart and format the fourth series data labels using "#,##0". | Explain how to enable data labels for a single series and apply a thousand‑separator number format in Aspose.Cells for .NET. | Provide step‑by‑step instructions to format chart series labels with commas in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook with four data series, adds a column chart, enables data labels only for the fourth series, and formats those labels with the "#,##0" pattern so values appear with commas as thousand separators. The workbook is then saved as an Excel file.
    public class ApplyThousandSeparatorToFourthSeries
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for four series (columns B to E) with categories in column A
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");

            // Series 1
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(2500);
            sheet.Cells["B4"].PutValue(3700);
            sheet.Cells["B5"].PutValue(4600);

            // Series 2
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(800);
            sheet.Cells["C3"].PutValue(1900);
            sheet.Cells["C4"].PutValue(3100);
            sheet.Cells["C5"].PutValue(4200);

            // Series 3
            sheet.Cells["D1"].PutValue("Series3");
            sheet.Cells["D2"].PutValue(1500);
            sheet.Cells["D3"].PutValue(2600);
            sheet.Cells["D4"].PutValue(3800);
            sheet.Cells["D5"].PutValue(4700);

            // Series 4 (the target series)
            sheet.Cells["E1"].PutValue("Series4");
            sheet.Cells["E2"].PutValue(2000);
            sheet.Cells["E3"].PutValue(3400);
            sheet.Cells["E4"].PutValue(4800);
            sheet.Cells["E5"].PutValue(5900);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add each series to the chart
            chart.NSeries.Add("B2:B5", true); // Series 1
            chart.NSeries.Add("C2:C5", true); // Series 2
            chart.NSeries.Add("D2:D5", true); // Series 3
            chart.NSeries.Add("E2:E5", true); // Series 4

            // Set category (X) data
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the fourth series
            Series fourthSeries = chart.NSeries[3]; // zero‑based index
            fourthSeries.DataLabels.ShowValue = true;

            // Apply thousand‑separator number format to the data labels of the fourth series
            fourthSeries.DataLabels.NumberFormat = "#,##0";

            // Save the workbook
            workbook.Save("ThousandSeparatorFourthSeries.xlsx");
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ApplyThousandSeparatorToFourthSeries.Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
