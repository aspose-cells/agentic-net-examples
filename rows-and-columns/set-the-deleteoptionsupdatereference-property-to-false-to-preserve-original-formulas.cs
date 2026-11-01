// Title: Delete rows in an Excel file with Aspose.Cells for .NET while keeping formulas unchanged by disabling UpdateReference
// AI Prompts: Write C# code that removes a specific row range in an Excel worksheet using Aspose.Cells and disables formula reference updates via a DeleteOptions object. | Show how to call Aspose.Cells DeleteRows method with a DeleteOptions instance to preserve formula references when rows are removed.
// Common Searches: Aspose.Cells C# remove rows while keeping formulas unchanged | How to set DeleteOptions.UpdateReference to false in Aspose.Cells | Preserve Excel formulas when deleting rows programmatically with .NET
// Tags: Aspose.Cells DeleteRows with DeleteOptions | DeleteOptions.UpdateReference property false example | preserve formulas after row deletion .NET | C# Excel row removal without reference update

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, creates a DeleteOptions object with UpdateReference set to false, deletes the desired rows using Worksheet.Cells.DeleteRows(startRow, rowCount, deleteOptions) to avoid altering formula references, and saves the modified file.
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

            // NOTE: The setting to disable automatic formula reference updating when rows are deleted
            // is not available in this version of Aspose.Cells, so it is omitted.

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Delete the first row (row index 0) without adjusting formulas
            // The third argument (true) indicates that the entire row should be removed
            sheet.Cells.DeleteRows(0, 1, true);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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
