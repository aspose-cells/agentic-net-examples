// Title: C# console program to scan a folder for .xlsm files and report VBA macro presence with Aspose.Cells
// Common Searches: how to list macro status of multiple XLSM workbooks in C# using Aspose.Cells | C# code to detect VBA modules in Excel macro-enabled files | batch process XLSM files to check for macros with Aspose.Cells | console application that reports if an .xlsm file contains macros in .NET | using Aspose.Cells Workbook.VbaProject to determine macro presence
// Tags: Aspose.Cells VBA macro detection | C# console batch XLSM processing | Workbook.VbaProject modules enumeration | report macro presence in Excel files | enumerate .xlsm files in folder .NET

using System;
using System.IO;
using Aspose.Cells;

namespace MacroStatusReporter
{
    // A console application that receives a folder path, enumerates all .xlsm files in that directory, loads each workbook with Aspose.Cells, checks the Workbook.VbaProject.Modules collection to determine if VBA macros exist, and outputs each file name with "Macros present", "No macros", or an error message.
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

            // Check if the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Error: The folder \"{folderPath}\" does not exist.");
                return;
            }

            // Get all .xlsm files in the folder (non‑recursive)
            string[] xlsmFiles = Directory.GetFiles(folderPath, "*.xlsm", SearchOption.TopDirectoryOnly);

            if (xlsmFiles.Length == 0)
            {
                Console.WriteLine("No XLSM files found in the specified folder.");
                return;
            }

            Console.WriteLine($"Processing {xlsmFiles.Length} XLSM file(s) in \"{folderPath}\":");
            Console.WriteLine();

            foreach (string filePath in xlsmFiles)
            {
                try
                {
                    // Load the workbook using Aspose.Cells
                    Workbook workbook = new Workbook(filePath);

                    // Determine macro presence
                    bool hasMacros = false;

                    // Aspose.Cells exposes the VBA project; if it exists and contains modules, macros are present
                    if (workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0)
                    {
                        hasMacros = true;
                    }

                    // Output the result for the current file
                    string fileName = Path.GetFileName(filePath);
                    Console.WriteLine($"{fileName}: {(hasMacros ? "Macros present" : "No macros")}");
                }
                catch (Exception ex)
                {
                    // Report any errors encountered while processing the file
                    string fileName = Path.GetFileName(filePath);
                    Console.WriteLine($"{fileName}: Error - {ex.Message}");
                }
            }

            Console.WriteLine();
            Console.WriteLine("Processing completed.");
        }
    }
}
