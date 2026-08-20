// Title: PowerShell Wrapper for Aspose.Cells .NET to Strip Printer Settings from Excel Files
// Description: A PowerShell‑compatible wrapper that invokes a compiled C# console app using Aspose.Cells for .NET. It loads an Excel workbook, clears the PageSetup.PrinterSettings on every worksheet, and saves the file to a new location. Includes argument validation, missing‑file checks, and automatic output‑folder creation.
// Keywords: PowerShell wrapper | Aspose.Cells | remove printer settings | Excel workbook | PageSetup.PrinterSettings | C# console utility | automate Excel sanitization | CI/CD Excel processing | strip printer configuration | .NET Excel library
// Common Searches: PowerShell script to clear printer settings in Excel using Aspose.Cells | How to remove stored printer configuration from .xlsx files | Aspose.Cells remove printer settings command line | Batch strip printer settings from Excel workbooks PowerShell | C# program to delete printer settings in Excel workbook
// Developer Intent: Provide a PowerShell‑friendly way to call an Aspose.Cells .NET executable that removes all printer settings from an Excel workbook.
// Use Cases: Sanitize confidential Excel reports before distribution by eliminating embedded printer configurations. | Integrate the wrapper into a CI/CD pipeline to ensure generated spreadsheets contain no printer metadata. | Run a scheduled PowerShell job that processes a folder of workbooks, stripping printer settings from each file.
// AI Prompts: Create a PowerShell script that runs RemovePrinterSettings.exe with input and output paths and captures any errors. | Write a PowerShell function Remove-ExcelPrinterSettings that loads the Aspose.Cells assembly, clears PageSetup.PrinterSettings for each worksheet, and returns a success flag. | Generate PowerShell code to batch process all .xlsx files in a directory using the compiled Aspose.Cells printer‑setting remover.

using System;
using System.IO;
using Aspose.Cells;

namespace RemovePrinterSettings
{
    // A PowerShell‑compatible wrapper that invokes a compiled C# console app using Aspose.Cells for .NET. It loads an Excel workbook, clears the PageSetup.PrinterSettings on every worksheet, and saves the file to a new location. Includes argument validation, missing‑file checks, and automatic output‑folder creation.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect exactly two arguments: input file path and output file path
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: RemovePrinterSettings <InputFile> <OutputFile>");
                return;
            }

            string inputFile = args[0];
            string outputFile = args[1];

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' does not exist.");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets and clear stored printer settings
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // PageSetup.PrinterSettings is a byte[]; setting it to null removes the stored settings
                    sheet.PageSetup.PrinterSettings = null;
                }

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Printer settings have been removed and the file has been saved to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
