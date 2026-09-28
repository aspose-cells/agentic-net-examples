// Title: Batch delete VBA modules named "Temp" from .xlsx workbooks and overwrite files using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, removes any VBA module whose name equals "Temp" (case‑insensitive), and saves the workbook back to its original location. | Enhance the program to create a log file that records the name of each workbook processed and indicates whether a "Temp" module was found and removed. | Add robust error handling that detects password‑protected Excel files, skips them, and continues processing the remaining workbooks without aborting.
// Common Searches: asp.net remove specific VBA module from multiple Excel files programmatically | how to batch delete a macro named Temp in .xlsx using Aspose.Cells C# | skip password protected Excel workbook when iterating with Aspose.Cells library | overwrite original Excel file after VBA cleanup with Aspose.Cells | C# code to remove VBA modules from workbooks in a folder
// Tags: batch delete VBA module Aspose.Cells | remove Temp macro .xlsx C# | ignore password‑protected workbooks Aspose.Cells | overwrite workbook after VBA removal | iterate workbook files Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookProcessor
{
    // Scans a directory for .xlsx files, loads each workbook with Aspose.Cells, deletes any VBA module named "Temp" (case‑insensitive), skips password‑protected files, and saves the modified workbook back to the original file.
    class Program
    {
        static void Main()
        {
            // Folder that contains the workbooks to process
            string folderPath = @"C:\Workbooks";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Get all Excel files in the folder
            string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsx");

            foreach (string filePath in workbookFiles)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (skip if password protected)
                    Workbook workbook = new Workbook(filePath);

                    // Remove VBA modules named "Temp" (case‑insensitive)
                    if (workbook.VbaProject != null)
                    {
                        var modules = workbook.VbaProject.Modules;
                        for (int i = modules.Count - 1; i >= 0; i--)
                        {
                            if (string.Equals(modules[i].Name, "Temp", StringComparison.OrdinalIgnoreCase))
                            {
                                modules.RemoveAt(i);
                            }
                        }
                    }

                    // Save the modified workbook, overwriting the original file
                    workbook.Save(filePath);
                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                }
                // Catch password‑protected files (Aspose.Cells throws CellsException with a message containing "Password")
                catch (CellsException ex) when (ex.Message != null && ex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"Skipping password‑protected file: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    // Log any other errors and continue processing remaining files
                    Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                }
            }
        }
    }
}
