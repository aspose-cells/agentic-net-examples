// Title: Enable data labels and increase their font size in a column chart using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart, turns on data labels for the first series, and sets the label font size to 14 points. | Update an existing Aspose.Cells chart example to display data label values and enlarge the label text for better readability.
// Common Searches: Aspose.Cells set data label font size in chart C# | C# Aspose.Cells enable data labels for column chart | increase chart data label text size using Aspose.Cells .NET | how to show values as data labels in Aspose.Cells chart | Aspose.Cells chart data label formatting example
// Tags: Aspose.Cells chart data label font size | enable data labels Aspose.Cells column chart | C# Aspose.Cells set data label visibility | customize chart data label appearance .NET | Aspose.Cells increase data label text size

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds a column chart with sample data, enables data labels for the first series, sets the label font size to 14 points, and saves the workbook as DataLabelResize_Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["A5"].PutValue("D");
            worksheet.Cells["B5"].PutValue(40);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set series and category data
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the first series
            var series = chart.NSeries[0];
            // Show values as data labels
            series.DataLabels.ShowValue = true;
            // Increase font size of data labels
            series.DataLabels.Font.Size = 14;

            // Save the workbook
            string outputPath = "DataLabelResize_Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
