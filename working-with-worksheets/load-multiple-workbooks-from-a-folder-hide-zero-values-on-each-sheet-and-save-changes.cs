// Title: Batch hide zero values in all worksheets of multiple Excel workbooks using Aspose.Cells for .NET
// Description: A C# console app that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, disables zero display on every worksheet via the DisplayZeros property, and overwrites the original files. Includes folder validation, file existence checks, and robust exception handling.
// Keywords: Aspose.Cells | C# hide zeros | batch process Excel workbooks | .NET Excel automation | DisplayZeros false | multiple workbook processing | Excel zero suppression | folder iteration Excel files
// Common Searches: How to hide zero values in all sheets of multiple Excel files using Aspose.Cells | C# batch hide zeros in Excel workbooks | Aspose.Cells set DisplayZeros false for many workbooks | Programmatically remove zero display from Excel worksheets .NET | Iterate through folder of .xlsx files and hide zeros
// Developer Intent: Load every .xlsx file in a specified directory, set DisplayZeros = false on each worksheet, and save the changes.
// Use Cases: Clean up generated reports by removing visible zeros across all sheets before distribution. | Prepare data packages for dashboards where zero values should be invisible. | Integrate into CI/CD pipelines to enforce zero‑value hiding on Excel artifacts automatically.
// AI Prompts: Generate C# code that uses Aspose.Cells to iterate over all Excel files in a folder, hide zero values on every worksheet, and overwrite the originals. | Explain best practices for exception handling and resource disposal when batch‑processing workbooks with Aspose.Cells in .NET. | Modify the sample to save the processed workbooks to a separate output directory while preserving the original files.

using System;
using System.IO;
using Aspose.Cells;

namespace HideZeroValuesInWorkbooks
{
    // A C# console app that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, disables zero display on every worksheet via the DisplayZeros property, and overwrites the original files. Includes folder validation, file existence checks, and robust exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the folder containing the Excel files.
                // Adjust this path as needed or pass it via command‑line arguments.
                string folderPath = @"C:\Path\To\Your\Folder";

                // Verify that the folder exists before attempting to enumerate files.
                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine($"Folder not found: {folderPath}");
                    return;
                }

                // Get all Excel files in the folder (you can adjust the pattern as needed).
                string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx");

                foreach (string filePath in excelFiles)
                {
                    // Ensure the file still exists before loading.
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    // Load the workbook inside a using block to guarantee disposal.
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Iterate through each worksheet and hide zero values.
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            sheet.DisplayZeros = false; // Hide zero values on this sheet.
                        }

                        // Save the workbook back to the same file (overwrites the original).
                        workbook.Save(filePath);
                    }
                }

                Console.WriteLine("Zero values hidden and workbooks saved successfully.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
