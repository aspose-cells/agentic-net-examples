// Title: PowerShell Wrapper for Aspose.Cells .NET to Strip Printer Settings from Excel Workbooks
// Description: A PowerShell‑based solution that invokes a compiled C# console app (or directly loads Aspose.Cells.dll) to load an Excel workbook, clear the PageSetup.PrinterSettings on every worksheet, and save the sanitized file. Ideal for automation, CI/CD pipelines, and secure distribution of Excel reports.
// Keywords: Aspose.Cells PowerShell wrapper | remove printer settings Excel | clear worksheet printer configuration | Worksheet.PageSetup.PrinterSettings null | .NET Excel automation | PowerShell call .NET library | sanitize Excel files | CI/CD Excel cleanup | command‑line Excel printer removal | GitHub Aspose.Cells example
// Common Searches: PowerShell script to remove printer settings from XLSX using Aspose.Cells | How to clear Excel printer configuration with .NET and PowerShell | Aspose.Cells example for stripping printer settings | Command line tool to sanitize Excel workbooks printer data | Load Aspose.Cells.dll in PowerShell and modify PageSetup
// Developer Intent: Create a PowerShell utility that leverages the Aspose.Cells .NET library to delete all embedded printer settings from a specified Excel file, enabling easy command‑line or scripted execution.
// Use Cases: Sanitize confidential Excel reports before sharing by removing printer metadata. | Integrate the wrapper into build or release pipelines to guarantee printer‑free workbooks. | Provide IT administrators a simple command‑line tool for bulk cleaning of Excel files on Windows servers.
// AI Prompts: Generate a PowerShell script that uses Add-Type to load Aspose.Cells.dll and replicates the C# logic to set Worksheet.PageSetup.PrinterSettings to null for each sheet. | Write a PowerShell function Remove-ExcelPrinterSettings that accepts input and output paths, invokes the compiled RemovePrinterSettings.exe, and returns a success status. | Create a PowerShell one‑liner that calls the Aspose.Cells .NET API to open an XLSX file, clear printer settings on all worksheets, and save the result.

using System;
using System.IO;
using Aspose.Cells;

namespace RemovePrinterSettings
{
    // A PowerShell‑based solution that invokes a compiled C# console app (or directly loads Aspose.Cells.dll) to load an Excel workbook, clear the PageSetup.PrinterSettings on every worksheet, and save the sanitized file. Ideal for automation, CI/CD pipelines, and secure distribution of Excel reports.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect exactly two arguments: input file and output file paths
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: RemovePrinterSettings <InputFile> <OutputFile>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                var workbook = new Workbook(inputPath);

                // Iterate through all worksheets and clear printer settings
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Setting PrinterSettings to null removes stored printer configuration
                    sheet.PageSetup.PrinterSettings = null;
                }

                // Save the modified workbook to the output file
                workbook.Save(outputPath);

                // Release unmanaged resources
                workbook.Dispose();

                Console.WriteLine($"Printer settings removed successfully. Output saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
