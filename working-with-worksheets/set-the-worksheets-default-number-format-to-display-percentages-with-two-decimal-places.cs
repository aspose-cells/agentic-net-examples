// Title: How to set a workbook’s default number format to a two‑decimal percentage in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set default cell style to show percentages with two decimal places | C# workbook default number format 0.00% using Aspose.Cells | Apply a custom percentage format to all cells in an Aspose.Cells workbook | Set default style custom format for percentages in Aspose.Cells .NET | How to display percentages with two decimals by default in Aspose.Cells
// Tags: default style custom number format Aspose.Cells | global percentage format for workbook cells | custom 0.00% number format Aspose.Cells | percentage display two decimal places .NET | apply default number format to entire workbook

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, assigns the custom number format "0.00%" to the workbook's DefaultStyle so every cell shows percentages with two decimal places, writes a sample value to A1, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Set the default number format for the entire workbook to display percentages with two decimal places
            // Using the workbook's DefaultStyle which is applied to all cells unless overridden
            workbook.DefaultStyle.Custom = "0.00%";

            // Get the first worksheet (optional, just to demonstrate access)
            var worksheet = workbook.Worksheets[0];

            // (Optional) Verify that the style is applied by setting a sample value
            worksheet.Cells["A1"].PutValue(0.1234); // Will display as 12.34%

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
