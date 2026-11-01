// Title: How to enable PreserveFormatting on a QueryTable to retain cell styles after refresh using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx workbook, locate the first QueryTable, set its PreserveFormatting property to true, and save the file with Aspose.Cells in C#. | Write C# code that checks for a QueryTable in a worksheet, enables formatting preservation, and handles missing files or tables gracefully. | Demonstrate preserving cell formatting during data refresh by configuring QueryTable.PreserveFormatting and saving the workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# set QueryTable PreserveFormatting to keep cell styles after refresh | How to prevent formatting loss when refreshing an Excel query table with Aspose.Cells | Enable PreserveFormatting on QueryTable in .NET workbook using Aspose.Cells | C# example for preserving Excel query table formatting after data update with Aspose.Cells
// Tags: Aspose.Cells QueryTable formatting preservation | C# retain cell styles after query refresh | enable QueryTable formatting retention .NET | prevent Excel cell style loss on data refresh | modify QueryTable settings using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads an existing Excel workbook, verifies that the first worksheet contains a QueryTable, activates the PreserveFormatting flag to keep cell styles during data refresh, and saves the modified workbook to a new file, handling missing files or tables and ensuring the output directory exists.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one QueryTable
            if (sheet.QueryTables.Count == 0)
            {
                Console.WriteLine("Error: No QueryTable found in the first worksheet.");
                return;
            }

            // Access the first QueryTable on the worksheet
            QueryTable queryTable = sheet.QueryTables[0];

            // Enable PreserveFormatting to keep cell styles after a refresh
            queryTable.PreserveFormatting = true;

            // NOTE: The RefreshData method is not available in the current Aspose.Cells version.
            // If needed, implement data refresh via alternative means (e.g., external data source).

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
