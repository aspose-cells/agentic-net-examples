// Title: How to set all chart titles to Arial 12‑point font in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, iterates over every worksheet and chart, and changes each chart title's font to Arial 12pt. | Create a reusable C# method using Aspose.Cells that updates the font name and size of all chart titles in a given workbook.
// Common Searches: Aspose.Cells C# change font of chart titles across all worksheets | loop through Excel charts and set title font to Arial 12 using Aspose.Cells | bulk update chart title style in an Excel workbook with Aspose.Cells .NET
// Tags: Aspose.Cells chart title font formatting | C# bulk update Excel chart titles | iterate worksheets and charts Aspose.Cells | set Arial font for Excel chart titles .NET | workbook chart title styling Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel file, walks through each worksheet and its charts, verifies the presence of a title, sets the title font to Arial size 12, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Ensure the chart has a title object
                    if (chart.Title != null)
                    {
                        // Set the title font to Arial, size 12
                        chart.Title.Font.Name = "Arial";
                        chart.Title.Font.Size = 12;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
