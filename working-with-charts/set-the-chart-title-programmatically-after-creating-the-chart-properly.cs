// Title: Programmatically set the title of a column chart in Aspose.Cells for .NET and save the workbook as XLSX
// AI Prompts: Write C# code that creates a new Workbook, adds a column chart, assigns a custom string to chart.Title.Text, ensures the title is visible, and saves the file as an XLSX using Aspose.Cells. | Update an existing Aspose.Cells chart object to change its Title.Text, toggle Title.IsVisible, and write the workbook to disk in .NET.
// Common Searches: asp.net set column chart title Aspose.Cells C# | how to change chart title text in Aspose.Cells workbook | Aspose.Cells make chart title visible programmatically | example of saving a chart with custom title using Aspose.Cells | C# Aspose.Cells chart title property usage
// Tags: Aspose.Cells chart title text property | Aspose.Cells set column chart title | Aspose.Cells chart title visibility | Aspose.Cells export workbook with chart | C# Aspose.Cells column chart creation

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example demonstrates creating a new Workbook, adding a column chart to the first worksheet, setting the chart's Title.Text to "Sales Report", explicitly enabling Title.IsVisible, and saving the workbook as ChartWithTitle.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet (create chart rule)
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart title programmatically
            chart.Title.Text = "Sales Report";

            // Optionally, make the title visible (default is true, but set explicitly)
            chart.Title.IsVisible = true;

            // Save the workbook (save rule)
            string outputPath = "ChartWithTitle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
