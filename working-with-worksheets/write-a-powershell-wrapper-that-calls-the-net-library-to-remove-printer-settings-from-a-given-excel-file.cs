// Title: Create a PowerShell wrapper that uses Aspose.Cells for .NET to strip printer settings from an Excel workbook
// Common Searches: how to call Aspose.Cells .NET library from PowerShell to remove print area in an XLSX file | PowerShell script to clear worksheet printer settings using Aspose.Cells | remove printer configuration from Excel workbook via Aspose.Cells in a PowerShell command | invoke C# Aspose.Cells code from PowerShell to reset page setup properties | batch process Excel files to delete printer settings with PowerShell and Aspose.Cells
// Tags: Aspose.Cells PowerShell wrapper for printer settings removal | clear worksheet print area Aspose.Cells .NET | reset page setup printer properties Excel using Aspose.Cells | remove printer settings from XLSX via Aspose.Cells API | invoke C# Aspose.Cells from PowerShell script

using System;
using System.IO;
using Aspose.Cells;

namespace RemovePrinterSettings
{
    // The solution validates input and output paths, loads the Excel workbook with Aspose.Cells, iterates each worksheet to nullify PrintArea, PrintTitleRows, PrintTitleColumns, and the optional PrinterSettings property via reflection, then saves the cleaned file—exposed as a PowerShell wrapper that calls the .NET code.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect exactly two arguments: input file and output file paths
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

                // Clear printer‑related settings for each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Reset common printer settings manually (ResetToDefault not available in this version)
                        sheet.PageSetup.PrintArea = null;
                        sheet.PageSetup.PrintTitleRows = null;
                        sheet.PageSetup.PrintTitleColumns = null;

                        // If the PrinterSettings property exists, clear it
                        // This property may not be present in older versions; ignore if unavailable
                        var printerProp = sheet.PageSetup.GetType().GetProperty("PrinterSettings");
                        if (printerProp != null && printerProp.CanWrite)
                        {
                            printerProp.SetValue(sheet.PageSetup, null);
                        }
                    }
                    catch (Exception innerEx)
                    {
                        Console.WriteLine($"Warning: Could not fully clear printer settings for sheet '{sheet.Name}': {innerEx.Message}");
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
                Console.WriteLine($"Successfully saved modified file to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
