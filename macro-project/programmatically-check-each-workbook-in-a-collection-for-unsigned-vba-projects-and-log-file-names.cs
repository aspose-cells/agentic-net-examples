// Title: Scan a collection of Excel workbooks and list those with unsigned VBA projects using Aspose.Cells in C#
// AI Prompts: Write a C# method that accepts an IEnumerable of Excel file paths, loads each workbook with Aspose.Cells, checks the VbaProject.IsSigned flag, and returns the paths where a VBA project exists but is not signed. | Enhance the method to handle missing files gracefully, output a warning for each absent file, and log any exceptions thrown while loading a workbook without stopping the overall scan. | Adapt the scanner so that instead of returning a list, it appends each identified unsigned‑VBA workbook path to a log file with a timestamp, creating the log if it does not exist.
// Common Searches: how to use Aspose.Cells in C# to find Excel files with unsigned VBA macros | C# code to iterate over multiple .xlsm files and check VBA project signature with Aspose | detect unsigned VBA projects in a batch of workbooks using Aspose.Cells library | list Excel workbooks that contain unsigned macro projects in .NET | Aspose.Cells VBA IsSigned property example for scanning directories
// Tags: batch scan Excel workbooks for unsigned VBA using Aspose.Cells | C# Aspose.Cells VBA project signature check | detect unsigned macro projects in .xlsm files | enumerate Excel files and evaluate VbaProject.IsSigned | log unsigned VBA workbooks in .NET application

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace VbaProjectChecker
{
    // Scans a set of Excel file paths, loads each workbook with Aspose.Cells, and returns (or logs) the full paths of those containing a VBA project whose IsSigned property is false, while safely handling missing files and load errors.
    public class UnsignedVbaScanner
    {
        /// <param name="workbookFiles">Collection of full file paths to Excel workbooks.</param>
        /// <returns>List of file names (with full path) that have an unsigned VBA project.</returns>
        public static List<string> GetWorkbooksWithUnsignedVba(IEnumerable<string> workbookFiles)
        {
            var unsignedWorkbooks = new List<string>();

            foreach (var filePath in workbookFiles)
            {
                // Ensure the file exists before attempting to load.
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook using Aspose.Cells.
                    var workbook = new Workbook(filePath);

                    // Check if the workbook contains a VBA project.
                    var vbaProject = workbook.VbaProject;

                    // If a VBA project exists and it is not signed, record the file name.
                    if (vbaProject != null && !vbaProject.IsSigned)
                    {
                        unsignedWorkbooks.Add(filePath);
                    }
                }
                catch (Exception ex)
                {
                    // Log any loading errors but continue processing other files.
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            return unsignedWorkbooks;
        }

        // Example usage.
        public static void Main(string[] args)
        {
            // Example collection of workbook file paths.
            var workbookPaths = new List<string>
            {
                @"C:\Workbooks\Report1.xlsx",
                @"C:\Workbooks\Report2.xlsm",
                @"C:\Workbooks\Report3.xlsb"
            };

            var unsignedFiles = GetWorkbooksWithUnsignedVba(workbookPaths);

            Console.WriteLine("Workbooks with unsigned VBA projects:");
            foreach (var file in unsignedFiles)
            {
                Console.WriteLine(file);
            }
        }
    }
}
