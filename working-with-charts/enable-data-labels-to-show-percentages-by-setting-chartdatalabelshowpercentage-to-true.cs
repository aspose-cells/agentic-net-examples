// Title: Show Percentage Data Labels on a Pie Chart with Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, adds category and value data, inserts a pie chart, binds the series, and sets ChartDataLabel.ShowPercentage to true while optionally hiding raw values. The workbook is then saved as an XLSX file, demonstrating how to display only percentages on chart data labels using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | pie chart | percentage data labels | ChartDataLabel.ShowPercentage | Excel chart percentages | hide values | data label customization | Aspose.Cells example
// Common Searches: Aspose.Cells show percentages on pie chart | C# set ChartDataLabel.ShowPercentage true | display only percentage labels in Excel chart using Aspose.Cells | hide raw values and show percentages Aspose.Cells .NET | how to enable percentage data labels in Aspose.Cells chart
// Developer Intent: Add a pie chart to an Excel workbook and configure its data labels to display percentages instead of raw numbers.
// Use Cases: Financial reports that illustrate expense distribution with percentage labels on a pie chart. | Sales dashboards where each product's market share is shown as a percentage without showing absolute values. | Automated Excel generation for presentations that require clean, percentage‑only chart annotations.
// AI Prompts: Generate C# code with Aspose.Cells that creates a donut chart and sets DataLabels.ShowPercentage = true while hiding values. | Provide an Aspose.Cells example that adds multiple series to a pie chart and configures each series to show percentage labels only. | Explain how to customize font style, size, and position of percentage data labels on an Aspose.Cells chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, adds category and value data, inserts a pie chart, binds the series, and sets ChartDataLabel.ShowPercentage to true while optionally hiding raw values. The workbook is then saved as an XLSX file, demonstrating how to display only percentages on chart data labels using Aspose.Cells for .NET.
    public class ShowPercentageDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a pie chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Define the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels and display percentages
                DataLabels dataLabels = chart.NSeries[0].DataLabels;
                dataLabels.ShowPercentage = true;
                dataLabels.ShowValue = false; // hide raw values if desired

                // Determine output file path
                string outputFile = "ShowPercentageDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ShowPercentageDemo.Run();
        }
    }
}
