// Title: Aspose.Cells for .NET – C# – Show Pie Chart Data Labels as Percentages with One Decimal Place
// Description: This C# example demonstrates how to create a workbook with Aspose.Cells, add a pie chart, and configure its data labels to display percentages formatted to one decimal place (0.0%). The code enables ShowPercentage, hides raw values, applies the NumberFormat property, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# chart example | pie chart data labels | percentage format 0.0% | show percentage Aspose.Cells | Excel chart formatting .NET | chart number format | Aspose.Cells tutorial | Excel export C# | hide values in chart labels
// Common Searches: Aspose.Cells pie chart percentage label C# | format chart data labels one decimal Aspose.Cells | C# show only percentages on pie chart | set NumberFormat for chart labels Aspose.Cells | how to hide values in chart data labels .NET
// Developer Intent: Add a pie chart and format its data labels to show percentages with one decimal place.
// Use Cases: Financial dashboard showing market‑share percentages with one‑decimal precision. | Sales report where each region's contribution is displayed as a formatted percentage. | Project‑management workbook illustrating task distribution via a pie chart with precise percentages. | Academic research presenting survey results with exact percentage slices. | Marketing presentation exporting product‑mix percentages to Excel with formatted labels.
// AI Prompts: Generate C# code using Aspose.Cells to create a donut chart with data labels formatted to two decimal places. | Explain how to apply custom number formats to chart data labels for any chart type in Aspose.Cells for .NET. | Provide a step‑by‑step guide to hide raw values and display only formatted percentages on Excel chart labels using Aspose.Cells. | Show how to programmatically change the font style of pie chart data labels in Aspose.Cells. | Write a script that reads data from a CSV and creates a pie chart with percentage labels formatted to one decimal place.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to create a workbook with Aspose.Cells, add a pie chart, and configure its data labels to display percentages formatted to one decimal place (0.0%). The code enables ShowPercentage, hides raw values, applies the NumberFormat property, and saves the result as an XLSX file.
    public class PieChartDataLabelsPercentage
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
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and configure them to show percentages
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowPercentage = true;          // Show percentage values
            dataLabels.ShowValue = false;              // Hide raw values (optional)
            dataLabels.NumberFormat = "0.0%";          // One decimal place percentage format

            // Define output file path
            string outputPath = "PieChartWithPercentageLabels.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
        }
    }
}
