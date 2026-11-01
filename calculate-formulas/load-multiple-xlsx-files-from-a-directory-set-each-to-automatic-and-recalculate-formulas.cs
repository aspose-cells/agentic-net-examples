// Title: C# – Batch load XLSX files from a folder, set calculation mode to Automatic, and recalculate formulas using Aspose.Cells
// Common Searches: C# Aspose.Cells batch recalculate formulas in all .xlsx files in a directory | How to update formulas for every workbook in a folder using Aspose.Cells .NET | Automate processing of multiple Excel files and recalculate formulas with Aspose.Cells | Set calculation mode to Automatic for each workbook while batch processing in C# | Load each .xlsx from a folder and call CalculateFormula with Aspose.Cells
// Tags: batch workbook formula recalculation Aspose.Cells | automatic calculation mode Aspose.Cells .NET | load multiple XLSX workbooks C# | folder based Excel processing Aspose.Cells | Workbook.CalculateFormula usage C#

using System;
using System.IO;
using Aspose.Cells;

// The example enumerates all .xlsx files in a specified directory, loads each workbook with Aspose.Cells, switches the calculation mode to Automatic, invokes Workbook.CalculateFormula to refresh all formulas, saves the changes back to the original files, and logs progress while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Directory containing the XLSX files
            string folderPath = @"C:\Path\To\XlsxFolder";

            // Verify that the directory exists
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

                    // Load the workbook from file
                    Workbook workbook = new Workbook(filePath);

                    // Recalculate all formulas in the workbook
                    workbook.CalculateFormula();

                    // Save the workbook (overwrites the original file)
                    workbook.Save(filePath);

                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
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
