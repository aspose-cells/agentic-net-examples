// Title: Recalculate formulas in all XLSX workbooks in a folder using Aspose.Cells for .NET (automatic mode)
// Common Searches: C# code to recalculate formulas in every .xlsx file in a directory with Aspose.Cells | How to batch process Excel workbooks and update formulas automatically using Aspose.Cells .NET | Load multiple XLSX files, set calculation mode to Automatic and save changes with Aspose.Cells | Aspose.Cells C# iterate over folder of Excel files and recalculate all formulas | Batch recalculate Excel formulas in .NET without opening Excel UI
// Tags: batch formula recalculation Aspose.Cells | load multiple xlsx workbooks C# | automatic calculation mode Aspose.Cells | process Excel files directory Aspose.Cells | overwrite original workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example enumerates every .xlsx file in a specified folder, loads each workbook with Aspose.Cells, triggers formula calculation in Automatic mode, and saves the workbook back to the original file, with error handling for missing files and processing exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Directory containing the XLSX files
            string folderPath = @"C:\Path\To\XlsxFolder";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Get all .xlsx files in the directory
            string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                try
                {
                    // Ensure the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Recalculate all formulas in the workbook (default mode is Automatic)
                    workbook.CalculateFormula();

                    // Save the workbook (overwrites the original file)
                    workbook.Save(filePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("All workbooks processed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
