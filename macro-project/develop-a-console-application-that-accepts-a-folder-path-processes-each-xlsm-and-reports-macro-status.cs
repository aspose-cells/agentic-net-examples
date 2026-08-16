// Title: C# Console App to Scan a Folder for XLSM Files and Report Macro Presence with Aspose.Cells
// Description: A .NET console program that accepts a folder path, enumerates all *.xlsm files in that directory, loads each workbook using Aspose.Cells, checks the Workbook.HasMacro property, and prints the file name with its macro status or any loading error.
// Keywords: Aspose.Cells macro detection | C# XLSM HasMacro | detect Excel macros .NET | batch scan Excel files for macros | console application Aspose.Cells | Workbook.HasMacro usage | macro audit Excel files | C# folder scan XLSM | Excel macro status report
// Common Searches: how to check if an XLSM file has macros using Aspose.Cells C# | C# console program to list macro-enabled Excel files in a folder | Aspose.Cells Workbook.HasMacro example | batch process Excel workbooks for macro presence | error handling loading XLSM with Aspose.Cells
// Developer Intent: Build a lightweight console tool that scans a specified directory for .xlsm files and outputs whether each workbook contains macros.
// Use Cases: Perform a security audit of macro‑enabled spreadsheets on shared drives. | Generate a pre‑migration report of macro usage before moving to a macro‑free environment. | Identify corrupted or unsupported macro files by capturing load exceptions.
// AI Prompts: Create a C# method that takes a folder path and returns a dictionary of XLSM file names with their HasMacro values using Aspose.Cells. | Add robust error logging to the MacroStatusChecker program, writing stack traces to a log file. | Modify the console app to recursively search subfolders and export the results to a CSV file.

using System;
using System.IO;
using Aspose.Cells;

namespace MacroStatusChecker
{
    // A .NET console program that accepts a folder path, enumerates all *.xlsm files in that directory, loads each workbook using Aspose.Cells, checks the Workbook.HasMacro property, and prints the file name with its macro status or any loading error.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a folder path was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: MacroStatusChecker <folderPath>");
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
                Console.WriteLine("No macro‑enabled Excel files (*.xlsm) found in the specified folder.");
                return;
            }

            Console.WriteLine($"Processing {xlsmFiles.Length} file(s) in \"{folderPath}\":");

            foreach (string filePath in xlsmFiles)
            {
                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Use the HasMacro property to determine macro presence
                    bool hasMacro = workbook.HasMacro;

                    // Report the result
                    Console.WriteLine($"{Path.GetFileName(filePath)} : HasMacro = {hasMacro}");
                }
                catch (Exception ex)
                {
                    // Report any errors encountered while processing the file
                    Console.WriteLine($"{Path.GetFileName(filePath)} : Error - {ex.Message}");
                }
            }
        }
    }
}
