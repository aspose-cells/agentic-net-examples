// Title: How to set every chart title to Arial 12‑pt font in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells change font of all chart titles to Arial 12 | Iterate over charts in a workbook and set title font size with Aspose.Cells | Bulk update Excel chart title style using Aspose.Cells .NET | Programmatically set chart title font name and size in Aspose.Cells | Aspose.Cells loop through worksheets to modify chart title formatting
// Tags: Aspose.Cells Chart.Title.Font property usage | C# enumerate worksheets and charts | Excel chart title font assignment Aspose | Bulk chart title font modification .NET | Programmatic chart title styling with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel file, iterates through each worksheet and its charts, and for every chart that has a title it sets the title's font to Arial with a size of 12 points, then saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each chart in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // If the chart has a title, set its font properties
                    if (chart.Title != null)
                    {
                        chart.Title.Font.Name = "Arial";
                        chart.Title.Font.Size = 12;
                    }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
