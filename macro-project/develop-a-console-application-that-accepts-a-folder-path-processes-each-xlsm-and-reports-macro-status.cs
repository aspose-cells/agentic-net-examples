// Title: Create a C# console tool that scans a folder for .xlsm workbooks and reports whether each file contains VBA macros using Aspose.Cells
// AI Prompts: Write a .NET console program that accepts a directory path, loads every *.xlsm file with Aspose.Cells, and prints the file name with "Contains macros" or "No macros" based on the Workbook.VbaProject.Modules collection. | Add comprehensive exception handling so that any workbook that cannot be opened outputs the file name followed by the error message. | Enhance the utility to optionally traverse subfolders recursively and display the full file path together with the macro status.
// Common Searches: how to list macro status of multiple xlsm files using Aspose.Cells in C# | C# program to detect VBA macros in Excel workbooks in a specific folder | Aspose.Cells check workbook.VbaProject.Modules for macros in batch | console application to report macro presence for all .xlsm files in a directory
// Tags: scan folder for xlsm files with Aspose.Cells | detect VBA macros using Workbook.VbaProject | report macro presence in Excel workbooks C# | batch macro detection Aspose.Cells .NET | console utility for macro status enumeration

using System;
using System.IO;
using Aspose.Cells;

namespace MacroStatusReporter
{
    // A .NET console application that receives a folder path, iterates over each .xlsm file, loads it with Aspose.Cells, checks the VbaProject.Modules collection to determine if VBA macros exist, and writes the file name with a status message or error details to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a folder path was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: MacroStatusReporter <folderPath>");
                return;
            }

            string folderPath = args[0];

            // Validate folder existence
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Error: The folder \"{folderPath}\" does not exist.");
                return;
            }

            // Get all XLSM files in the folder (non‑recursive)
            string[] files = Directory.GetFiles(folderPath, "*.xlsm", SearchOption.TopDirectoryOnly);

            if (files.Length == 0)
            {
                Console.WriteLine("No XLSM files found in the specified folder.");
                return;
            }

            Console.WriteLine($"Processing {files.Length} XLSM file(s) in \"{folderPath}\":");
            Console.WriteLine();

            foreach (string filePath in files)
            {
                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Determine macro presence
                    bool hasMacros = workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0;

                    string status = hasMacros ? "Contains macros" : "No macros";

                    Console.WriteLine($"{Path.GetFileName(filePath)}: {status}");
                }
                catch (Exception ex)
                {
                    // Report any errors while processing the file
                    Console.WriteLine($"{Path.GetFileName(filePath)}: Error - {ex.Message}");
                }
            }
        }
    }
}
