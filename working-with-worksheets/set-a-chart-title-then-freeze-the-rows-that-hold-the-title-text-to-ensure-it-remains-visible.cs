// Title: Add a column chart with a custom title and freeze the title row in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate an Excel workbook, populate sample data, insert a column chart, set its Title.Text to "Sales Comparison", write the same text into cell A1, and apply FreezePanes to keep the first row visible using Aspose.Cells for .NET. | Using C# and Aspose.Cells, create a worksheet, add a column chart with a custom title, copy the title into a worksheet cell, then freeze the top row so the title remains on screen.
// Common Searches: c# aspose.cells set chart title and freeze header row | how to keep chart title row visible with freeze panes in Aspose.Cells | example of column chart with title and frozen first row using Aspose.Cells for .NET | asp.net freeze first worksheet row after adding a chart title
// Tags: set chart title Aspose.Cells C# | freeze first row Aspose.Cells | column chart creation Aspose.Cells | write chart title to worksheet cell Aspose.Cells | apply FreezePanes after chart insertion Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The program creates a workbook, adds sample data, inserts a column chart titled "Sales Comparison", writes the title into cell A1, freezes the first row, and saves the file as ChartWithTitleAndFrozenRow.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(20);
            sheet.Cells["A4"].PutValue(30);
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["B4"].PutValue(35);

            // Insert a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series
            chart.NSeries.Add("A2:A4", true);
            // Category data (optional). If the API version does not support CategoryData, this line can be omitted.
            // chart.NSeries[0].CategoryData = "B2:B4";

            // Set the chart title
            chart.Title.Text = "Sales Comparison";

            // Optionally write the same title text into a cell (e.g., A1) for visual reference
            sheet.Cells["A1"].PutValue("Sales Comparison");

            // Freeze the first row (freeze rows above row index 1)
            sheet.FreezePanes(1, 0, 1, 0);

            // Determine output file path
            string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "ChartWithTitleAndFrozenRow.xlsx");

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
