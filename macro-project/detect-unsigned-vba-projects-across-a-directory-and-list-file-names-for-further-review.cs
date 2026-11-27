// Title: Scan a folder for Excel workbooks with unsigned VBA projects using Aspose.Cells for .NET
// Common Searches: C# code to list Excel files that contain unsigned VBA macros with Aspose.Cells | how to detect unsigned VBA projects in .xlsm and .xlsb files using Aspose.Cells | enumerate Excel workbooks in a directory and check VbaProject.IsSigned in C# | Aspose.Cells find workbooks with unsigned macro projects across multiple folders | batch scan Excel files for unsigned VBA projects in .NET
// Tags: Aspose.Cells VBA project unsigned detection | Workbook.VbaProject.IsSigned verification | scan Excel directory for unsigned macros | enumerate .xls .xlsx .xlsm .xlsb files C# | batch processing of Excel workbooks with Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace VbaProjectScanner
{
    // The example recursively enumerates all Excel files (xls, xlsx, xlsm, xlsb) in a given folder, loads each workbook with Aspose.Cells, checks whether a VBA project exists and if it is unsigned via the Workbook.VbaProject.IsSigned property, collects the paths of those workbooks, and outputs the list for further review.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory to scan – change as needed
            string folderPath = @"C:\ExcelFiles";

            // Supported Excel extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

            // List to hold files with unsigned VBA projects
            List<string> unsignedVbaFiles = new List<string>();

            // Enumerate all files with the supported extensions
            foreach (string filePath in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories))
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue; // Skip non‑Excel files

                try
                {
                    // Load the workbook (uses Aspose.Cells lifecycle rule)
                    Workbook workbook = new Workbook(filePath);

                    // Check if a VBA project exists
                    if (workbook.VbaProject != null)
                    {
                        // If the project is not signed, record the file name
                        if (!workbook.VbaProject.IsSigned)
                        {
                            unsignedVbaFiles.Add(filePath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log or handle files that cannot be opened
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            // Output the results
            Console.WriteLine("Files with unsigned VBA projects:");
            foreach (string file in unsignedVbaFiles)
            {
                Console.WriteLine(file);
            }
        }
    }
}
