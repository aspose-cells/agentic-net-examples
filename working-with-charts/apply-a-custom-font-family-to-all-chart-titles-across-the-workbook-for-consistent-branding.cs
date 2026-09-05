// Title: Apply a custom font family to every chart title in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that opens an existing .xlsx file, loops through all worksheets and their charts, checks if a chart title exists, and sets the title's Font.Name to a specified brand font before saving. | Show how to safely handle missing chart titles while applying a consistent custom font to all chart titles in a workbook and then export the modified file.
// Common Searches: c# aspose.cells change font of all chart titles in a workbook | how to set a custom font for Excel chart titles programmatically using Aspose.Cells | loop through worksheets and charts to apply branding font to chart titles in .xlsx | Aspose.Cells bulk update chart title style across multiple sheets | apply corporate font to chart titles with Aspose.Cells C# example
// Tags: Aspose.Cells chart title font customization | C# loop through workbook charts | set corporate font for Excel chart titles | bulk update chart title appearance Aspose.Cells | customize chart title style in .xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel file, iterates over each worksheet and each chart, checks for a chart title, applies the custom font "MyCustomFont" to the title's Font.Name, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each chart on the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Apply custom font name to the chart title if it exists
                    if (chart.Title != null)
                    {
                        chart.Title.Font.Name = "MyCustomFont";
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
