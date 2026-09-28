// Title: PowerShell wrapper for Aspose.Cells .NET to strip printer settings (print area, title rows, title columns) from every worksheet in an Excel workbook
// AI Prompts: Create a PowerShell wrapper that launches a C# console application built with Aspose.Cells to open an Excel file, clear its PrintArea, PrintTitleRows, and PrintTitleColumns across all worksheets, and write the cleaned file to a target location. | Write PowerShell code that validates input and output arguments, ensures the destination folder exists, and uses the Aspose.Cells .NET assembly to remove printer‑related page‑setup properties from each worksheet. | Develop a PowerShell function that accepts an Excel file path, calls Aspose.Cells to reset page orientation, margins, and all printer settings, then returns the path of the updated workbook.
// Common Searches: how to call a C# Aspose.Cells library from PowerShell to clear print settings | PowerShell script to remove print area and title rows from Excel using Aspose.Cells | batch process Excel files to delete printer settings with Aspose.Cells and PowerShell | invoke Aspose.Cells .NET from PowerShell to reset worksheet page setup
// Tags: PowerShell invoke Aspose.Cells .NET library | remove printer settings from Excel worksheets | clear print area and title rows with Aspose.Cells | batch page setup cleanup for .xlsx files | Aspose.Cells reset worksheet page orientation

using System;
using System.IO;
using Aspose.Cells;

namespace RemovePrinterSettings
{
    // // Loads an Excel workbook with Aspose.Cells, iterates through each worksheet, clears PrintArea, PrintTitleRows, and PrintTitleColumns via the PageSetup object, ensures the output directory exists, and saves the modified file to the specified path. The logic can be wrapped in a PowerShell script that calls the compiled .NET program.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: RemovePrinterSettings <InputFile> <OutputFile>");
                return;
            }

            string inputFile = args[0];
            string outputFile = args[1];

            // Ensure the input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets and clear printer‑related settings
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Clear print area and title settings
                    sheet.PageSetup.PrintArea = string.Empty;
                    sheet.PageSetup.PrintTitleRows = string.Empty;
                    sheet.PageSetup.PrintTitleColumns = string.Empty;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook to the output file
                workbook.Save(outputFile);
                Console.WriteLine($"Printer settings removed and file saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
