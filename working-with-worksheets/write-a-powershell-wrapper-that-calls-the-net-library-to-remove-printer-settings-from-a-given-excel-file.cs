// Title: PowerShell wrapper that uses Aspose.Cells for .NET to remove printer settings from an Excel workbook
// Common Searches: how to use PowerShell to call Aspose.Cells .NET and clear printer settings in an Excel file | PowerShell script for removing worksheet printer configuration with Aspose.Cells | C# console application example that strips printer settings from all worksheets using Aspose.Cells | remove page setup printer settings from an Excel workbook via Aspose.Cells API | automate clearing of printer settings in Excel files using PowerShell and Aspose.Cells
// Tags: PowerShell wrapper Aspose.Cells printer settings removal | clear worksheet printer settings using Aspose.Cells | remove Excel printer configuration programmatically .NET | Aspose.Cells PageSetup PrinterSettings null assignment | batch clear printer settings in Excel workbook

using System;
using System.IO;
using Aspose.Cells;

namespace RemovePrinterSettings
{
    // The script loads an Excel file with Aspose.Cells, iterates through each worksheet, sets PageSetup.PrinterSettings to null to clear any printer configuration, ensures the output directory exists, and saves the cleaned workbook to the specified path.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input and output file paths
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: RemovePrinterSettings <inputFile> <outputFile>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Remove printer settings for each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Clear printer settings if present
                        sheet.PageSetup.PrinterSettings = null;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not clear printer settings for sheet '{sheet.Name}': {ex.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
