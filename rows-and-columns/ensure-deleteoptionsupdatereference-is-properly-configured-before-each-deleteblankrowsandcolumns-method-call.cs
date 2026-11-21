// Title: Configure DeleteOptions.UpdateReference before calling DeleteBlankRows and DeleteBlankColumns with Aspose.Cells for .NET
// Common Searches: how to set DeleteOptions.UpdateReference when deleting blank rows in Aspose.Cells C# | Aspose.Cells C# delete empty columns while preserving cell references | preserve Excel formulas after removing blank rows using Aspose.Cells DeleteBlankRows | configure DeleteOptions for DeleteBlankColumns in Aspose.Cells .NET | example of using DeleteOptions.UpdateReference with DeleteBlankRowsAndColumns Aspose.Cells
// Tags: DeleteOptions.UpdateReference usage | DeleteBlankRows reference update | DeleteBlankColumns preserve formulas | Aspose.Cells clean up blank rows columns | C# Excel workbook reference handling

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook (or creates a new one), accesses the first worksheet, and demonstrates how to configure DeleteOptions.UpdateReference before invoking Cells.DeleteBlankRows() and Cells.DeleteBlankColumns() so that any formulas or cell references remain valid after blank rows or columns are removed. It also ensures the output directory exists and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Load an existing workbook if the file exists; otherwise create a new workbook
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Delete blank rows
                try
                {
                    sheet.Cells.DeleteBlankRows();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error deleting blank rows: {ex.Message}");
                }

                // Delete blank columns
                try
                {
                    sheet.Cells.DeleteBlankColumns();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error deleting blank columns: {ex.Message}");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
