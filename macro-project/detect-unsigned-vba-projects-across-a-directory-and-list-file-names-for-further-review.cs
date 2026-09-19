// Title: Find Excel workbooks with unsigned VBA projects in a folder using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that recursively scans a given directory for .xls, .xlsx, .xlsm, and .xlsb files and prints the full paths of those whose VbaProject.IsSigned property is false, using Aspose.Cells. | Extend the scanner to create a CSV log that records the workbook name, file path, and a column indicating whether the VBA project is unsigned. | Add robust error handling so that any file that cannot be opened is written to an error log while the scan continues processing the remaining Excel files.
// Common Searches: asp.net detect unsigned VBA macros in Excel files with Aspose.Cells | c# code to list workbooks that have unsigned VBA projects in a directory | how to check VBA project signature status using Aspose.Cells VbaProject.IsSigned | scan folder for .xlsm files with unsigned macro projects in C# | list Excel files with unsigned VBA using Aspose.Cells API
// Tags: enumerate Excel files unsigned VBA detection | Aspose.Cells VbaProject.IsSigned check | C# directory scan for unsigned macro projects | list workbooks with unsigned VBA using Aspose.Cells | recursive Excel file processing Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject

// The example recursively enumerates .xls, .xlsx, .xlsm, and .xlsb files in a specified folder, loads each workbook with Aspose.Cells, checks the VbaProject.IsSigned flag, collects paths of files that contain an unsigned VBA project, and outputs the list (or logs it) for further review.
class UnsignedVbaDetector
{
    static void Main(string[] args)
    {
        // Determine directory to scan – use argument or default
        string targetDirectory = args.Length > 0 ? args[0] : @"C:\ExcelFiles";

        if (!Directory.Exists(targetDirectory))
        {
            Console.Error.WriteLine($"Directory not found: {targetDirectory}");
            return;
        }

        // Supported Excel extensions
        string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

        // List to hold files with unsigned VBA projects
        List<string> unsignedVbaFiles = new List<string>();

        // Enumerate all files with the supported extensions
        foreach (string filePath in Directory.EnumerateFiles(targetDirectory, "*.*", SearchOption.AllDirectories))
        {
            if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                continue; // Skip non‑Excel files

            if (!File.Exists(filePath))
                continue; // Safety check

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Access VBA project (may be null)
                VbaProject vba = workbook.VbaProject;
                if (vba != null && !vba.IsSigned)
                {
                    // Unsigned VBA project found – add to result list
                    unsignedVbaFiles.Add(filePath);
                }
            }
            catch (Exception ex)
            {
                // Log loading errors but continue processing other files
                Console.Error.WriteLine($"Error processing '{filePath}': {ex.Message}");
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
