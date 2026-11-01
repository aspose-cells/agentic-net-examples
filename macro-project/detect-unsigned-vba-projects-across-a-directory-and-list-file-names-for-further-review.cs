// Title: Find and list Excel files with unsigned VBA projects in a folder using Aspose.Cells for .NET
// Common Searches: C# code to scan a directory for Excel workbooks that have unsigned VBA macros using Aspose.Cells | how to identify unsigned VBA projects in .xlsm files with Aspose.Cells .NET | list file paths of workbooks with unsigned VBA projects in a folder C# Aspose.Cells | detect unsigned macro projects in multiple Excel files programmatically | recursive search for unsigned VBA projects in Excel files using Aspose.Cells API
// Tags: detect unsigned VBA projects Aspose.Cells | recursive scan of Excel workbooks for unsigned macros | Workbook.VbaProject.IsSigned property usage | list file paths of workbooks with unsigned VBA | process .xls .xlsm .xlsb .xlsx files Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// The example recursively walks through a specified directory, loads each Excel workbook (*.xls, *.xlsm, *.xlsb, *.xlsx) with Aspose.Cells, checks whether a VBA project exists and if its IsSigned flag is false, collects the paths of those files, and outputs the list while logging any files that cannot be processed.
class Program
{
    static void Main(string[] args)
    {
        // Directory to scan – can be supplied as a command‑line argument
        string folderPath = args.Length > 0 ? args[0] : @"C:\Spreadsheets";

        // List to hold file names that contain an unsigned VBA project
        List<string> unsignedVbaFiles = new List<string>();

        // Excel file extensions to examine
        string[] extensions = new[] { "*.xls", "*.xlsm", "*.xlsb", "*.xlsx" };

        foreach (string ext in extensions)
        {
            // Search recursively for matching files
            foreach (string filePath in Directory.GetFiles(folderPath, ext, SearchOption.AllDirectories))
            {
                try
                {
                    // Load the workbook (Aspose.Cells)
                    Workbook workbook = new Workbook(filePath);

                    // If a VBA project exists, check its signature status
                    if (workbook.VbaProject != null && !workbook.VbaProject.IsSigned)
                    {
                        unsignedVbaFiles.Add(filePath);
                    }
                }
                catch (Exception ex)
                {
                    // Log files that cannot be processed (corrupt, unsupported, etc.)
                    Console.Error.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
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
