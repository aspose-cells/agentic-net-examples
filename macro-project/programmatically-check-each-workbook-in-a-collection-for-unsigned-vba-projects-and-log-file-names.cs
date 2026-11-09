// Title: C# program to scan a folder of Excel workbooks and log files that contain unsigned VBA projects using Aspose.Cells
// Common Searches: C# Aspose.Cells find Excel files with unsigned VBA macros | How to detect unsigned VBA projects in multiple .xlsm files using Aspose.Cells | Log workbook names that have unsigned VBA code with Aspose.Cells .NET | Iterate through Excel workbooks and check VbaProject.IsSigned in C# | Batch process Excel files to identify unsigned macro projects using Aspose.Cells
// Tags: scan Excel workbooks for unsigned VBA projects | Aspose.Cells VbaProject.IsSigned verification | batch detect unsigned macros in .xlsm files | log unsigned VBA workbook names C# | enumerate Excel files with Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject

namespace VbaProjectChecker
{
    // The example loads every *.xls* file in a specified directory with Aspose.Cells, examines each workbook's VbaProject, and records the file names of workbooks whose VBA projects exist but are not signed (VbaProject.IsSigned = false). It outputs the list to the console and writes it to an UnsignedVbaLog.txt file.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to inspect
            string folderPath = @"C:\Workbooks";

            // Verify that the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Collect all Excel files (xls, xlsx, xlsm, etc.)
            string[] workbookFiles = Directory.GetFiles(folderPath, "*.xls*");

            // List to hold names of workbooks with unsigned VBA projects
            List<string> unsignedVbaWorkbooks = new List<string>();

            foreach (string filePath in workbookFiles)
            {
                try
                {
                    // Ensure the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    // Load the workbook using Aspose.Cells (lifecycle rule: load)
                    Workbook workbook = new Workbook(filePath);

                    // Access the VBA project, if any
                    VbaProject vbaProject = workbook.VbaProject;

                    // Check if a VBA project exists and is not signed
                    if (vbaProject != null && !vbaProject.IsSigned)
                    {
                        // Record the file name
                        unsignedVbaWorkbooks.Add(Path.GetFileName(filePath));
                    }
                }
                catch (Exception ex)
                {
                    // Log any errors and continue processing other files
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            // Output the results
            Console.WriteLine("Workbooks with unsigned VBA projects:");
            foreach (string fileName in unsignedVbaWorkbooks)
            {
                Console.WriteLine(fileName);
            }

            // Optionally, write the log to a text file
            try
            {
                string logPath = Path.Combine(folderPath, "UnsignedVbaLog.txt");
                File.WriteAllLines(logPath, unsignedVbaWorkbooks);
                Console.WriteLine($"Log written to: {logPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write log file: {ex.Message}");
            }
        }
    }
}
