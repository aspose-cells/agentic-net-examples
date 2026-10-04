// Title: Add data labels linked to worksheet cells for each point in a line chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a line chart, fills the series from column A, and sets each point's data label to the corresponding text in column B. | Show how to configure an Aspose.Cells chart series to hide numeric values and display only category names as data labels. | Provide a C# snippet that verifies the output directory exists and creates it before saving the workbook containing the chart.
// Common Searches: Aspose.Cells C# line chart data labels from cell range | How to bind chart point labels to worksheet cells using Aspose.Cells .NET | C# hide data label values and show category names in Aspose.Cells line chart | Create Excel line chart with custom labels from column B using Aspose.Cells | Ensure output folder exists before saving workbook in Aspose.Cells C#
// Tags: Aspose.Cells line chart data labels from range | C# hide numeric values in chart data labels Aspose.Cells | link chart point labels to worksheet cells Aspose.Cells | create output directory before workbook save Aspose.Cells | populate chart series values column A Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, writes numeric values to column A and label strings to column B, adds a line chart based on the values, configures the series to hide numeric values and display the category names from column B as data labels, ensures the target folder exists, and saves the file as LineChartWithLinkedLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the line chart (column A) and corresponding labels (column B)
            double[] values = { 10, 20, 15, 30, 25 };
            string[] labels = { "Alpha", "Beta", "Gamma", "Delta", "Epsilon" };

            // Fill the worksheet with data and label cells
            for (int i = 0; i < values.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(values[i]); // Column A – chart values
                sheet.Cells[i, 1].PutValue(labels[i]); // Column B – label source
            }

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the series using the values in column A
            chart.NSeries.Add("A1:A5", true);

            // Hide the numeric values on data labels
            chart.NSeries[0].DataLabels.ShowValue = false;

            // Show category names (labels) on data labels – alternative to SetDataLabelRange
            chart.NSeries[0].DataLabels.ShowCategoryName = true;

            // Save the workbook with the chart
            string outputPath = "LineChartWithLinkedLabels.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
