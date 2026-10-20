// Title: Detect and log unsigned VBA projects in multiple Excel workbooks using Aspose.Cells for .NET (C#)
// Common Searches: C# code to scan a directory of .xlsx files and list workbooks with unsigned VBA macros using Aspose.Cells | How to identify unsigned VBA projects in Excel files with Aspose.Cells .NET library | Batch process Excel workbooks to find unsigned macro projects and generate a log file in C# | Skip password‑protected Excel files while checking VBA signature status with Aspose.Cells | Write a list of workbooks containing unsigned VBA projects to a log using Aspose.Cells
// Tags: detect unsigned VBA projects Aspose.Cells | batch scan Excel workbooks for unsigned macros C# | log unsigned VBA workbooks Aspose.Cells | skip password protected Excel files Aspose.Cells | enumerate .xlsx files directory C# Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// The program enumerates all .xlsx files in a given folder (including subfolders), loads each workbook with Aspose.Cells, checks whether a VBA project exists and if it is unsigned, collects the file names of such workbooks, writes the list to a log file, and outputs the results to the console while gracefully handling password‑protected files and other exceptions.
class Program
{
    static void Main(string[] args)
    {
        // Folder that contains the workbooks to be examined
        string folderPath = @"C:\Workbooks";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Retrieve all Excel files in the folder (including subfolders)
        string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.AllDirectories);

        // List to collect names of workbooks with unsigned VBA projects
        List<string> unsignedVbaWorkbooks = new List<string>();

        foreach (string filePath in workbookFiles)
        {
            // Ensure the file still exists before attempting to load
            if (!File.Exists(filePath))
                continue;

            try
            {
                // Load the workbook (no password supplied). If a password is required,
                // Aspose.Cells will throw a CellsException which we handle below.
                Workbook workbook = new Workbook(filePath);

                // If the workbook contains a VBA project, check its signature status
                if (workbook.VbaProject != null && !workbook.VbaProject.IsSigned)
                {
                    // Add the file name (without path) to the log list
                    unsignedVbaWorkbooks.Add(Path.GetFileName(filePath));
                }
            }
            catch (CellsException ex)
            {
                // Detect password‑protected files by inspecting the exception message
                if (ex.Message != null && ex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"Skipped password‑protected file: {Path.GetFileName(filePath)}");
                }
                else
                {
                    Console.WriteLine($"CellsException processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log any other unexpected errors and continue
                Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
            }
        }

        // Write the results to a log file in the same folder
        string logFilePath = Path.Combine(folderPath, "UnsignedVbaWorkbooks.log");
        try
        {
            File.WriteAllLines(logFilePath, unsignedVbaWorkbooks);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write log file: {ex.Message}");
        }

        // Also output the results to the console for immediate feedback
        Console.WriteLine("Workbooks with unsigned VBA projects:");
        foreach (string name in unsignedVbaWorkbooks)
        {
            Console.WriteLine(name);
        }
    }
}
