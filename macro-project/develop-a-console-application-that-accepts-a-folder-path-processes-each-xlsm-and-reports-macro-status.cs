// Title: C# console app that scans a folder and reports macro presence in each XLSM workbook using Aspose.Cells
// Common Searches: how to list macro status of all .xlsm files in a directory with Aspose.Cells C# | C# program to detect VBA modules in multiple Excel macro‑enabled workbooks | batch process XLSM files to check for macros using Aspose.Cells library | command line tool for reporting macro presence in Excel files with Aspose.Cells | enumerate VBA projects in .xlsm files from a C# console application
// Tags: Aspose.Cells batch macro detection XLSM | C# console folder enumeration Excel workbooks | detect VBA modules Aspose.Cells | report macro presence command line | process multiple XLSM files Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace MacroStatusReporter
{
    // The console application receives a folder path as a command‑line argument, validates the directory, enumerates all .xlsm files (non‑recursive), loads each workbook with Aspose.Cells, checks whether the VbaProject contains any modules to determine macro presence, and writes the file name with either "Macros present" or "No macros" to the console. It also handles missing files and runtime exceptions gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a folder path was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide a folder path as the first argument.");
                return;
            }

            string folderPath = args[0];

            // Check if the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"The folder \"{folderPath}\" does not exist.");
                return;
            }

            // Get all .xlsm files in the folder (non‑recursive)
            string[] xlsmFiles = Directory.GetFiles(folderPath, "*.xlsm", SearchOption.TopDirectoryOnly);

            if (xlsmFiles.Length == 0)
            {
                Console.WriteLine("No .xlsm files found in the specified folder.");
                return;
            }

            Console.WriteLine($"Processing {xlsmFiles.Length} .xlsm file(s) in \"{folderPath}\":");
            Console.WriteLine();

            foreach (string filePath in xlsmFiles)
            {
                try
                {
                    // Ensure the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"{Path.GetFileName(filePath)}: File not found.");
                        continue;
                    }

                    // Load the workbook (auto‑detect format; .xlsm is supported)
                    Workbook workbook = new Workbook(filePath);

                    // Determine macro status
                    bool hasMacros = workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0;

                    // Output the result
                    Console.WriteLine($"{Path.GetFileName(filePath)}: {(hasMacros ? "Macros present" : "No macros")}");
                }
                catch (Exception ex)
                {
                    // Report any errors encountered while processing the file
                    Console.WriteLine($"{Path.GetFileName(filePath)}: Error - {ex.Message}");
                }
            }

            Console.WriteLine();
            Console.WriteLine("Processing complete.");
        }
    }
}
