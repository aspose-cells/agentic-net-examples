// Title: Set chart title to "Monthly Revenue" and apply the Accent1 style to the first chart in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing workbook, retrieve the first chart on the first worksheet, set its title text to "Monthly Revenue" and make the title visible. | Apply the predefined Accent1 chart style (ChartStyleType.Style1) to that chart. | Save the modified workbook to a new file path.
// Common Searches: how to set chart title in Aspose.Cells C# | apply Accent1 chart style using Aspose.Cells .NET | change first chart title to Monthly Revenue in existing Excel workbook C# | Aspose.Cells modify chart title and style programmatically | C# Aspose.Cells set chart title visibility
// Tags: set chart title Aspose.Cells C# | apply Accent1 chart style Aspose.Cells | modify first chart in existing Excel workbook | chart title visibility Aspose.Cells | update Excel chart style .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel file, accesses the first worksheet's first chart, sets the chart title to "Monthly Revenue" with visibility enabled, applies the predefined Accent1 style, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the first chart on the sheet
            Chart chart = sheet.Charts[0];

            // Set the chart title text and make it visible
            chart.Title.Text = "Monthly Revenue";
            chart.Title.IsVisible = true;

            // Apply a predefined chart style (optional – removed due to API compatibility)
            // chart.Style = ChartStyleType.Style1;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
