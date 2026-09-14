// Title: Export a merged Aspose.Cells workbook to Excel 97‑2003 (XLS) with compatibility options in C#
// AI Prompts: Write C# code that loads a workbook, configures XlsSaveOptions (MatchColor, ValidateMergedAreas, MergeAreas, SortNames), and saves it as an XLS file. | Show how to use Aspose.Cells XlsSaveOptions to preserve the 56‑color palette and validate merged cells when exporting to Excel 97‑2003. | Create a console application that accepts input and output paths, loads a merged workbook, and exports it to legacy XLS format with full compatibility settings.
// Common Searches: aspnet export merged workbook to xls using Aspose.Cells | c# XlsSaveOptions preserve colors validate merged areas | how to save Excel 97-2003 file with merged cells using Aspose.Cells | legacy xls export settings Aspose.Cells C# | convert .xlsx to .xls preserving merged cells Aspose.Cells
// Tags: merged workbook export XLS Aspose.Cells | XlsSaveOptions MatchColor ValidateMergedAreas | preserve 56‑color palette Aspose.Cells | legacy Excel 97‑2003 compatibility C# | console application Aspose.Cells XLS export

using System;
using System.IO;
using Aspose.Cells;

namespace ExportExample
{
    // Uses Aspose.Cells XlsSaveOptions to save a merged workbook as an Excel 97‑2003 XLS file, preserving colors, validating merged areas, merging conditional formatting, and sorting defined names for legacy compatibility.
    public class ExportToXls
    {
        // Exports a merged workbook to the legacy XLS format.
        public void Export(Workbook mergedWorkbook, string outputFilePath)
        {
            try
            {
                // Create save options for Excel 97‑2003 (XLS) files.
                XlsSaveOptions saveOptions = new XlsSaveOptions
                {
                    MatchColor = true,               // Preserve original colors within the 56‑color palette.
                    ValidateMergedAreas = true,      // Validate merged cells before saving.
                    MergeAreas = true,               // Merge conditional formatting and validation areas.
                    SortNames = true                 // Sort defined names for compatibility.
                };

                // Save the workbook using the specified options.
                mergedWorkbook.Save(outputFilePath, saveOptions);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error exporting to XLS: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        // Entry point for the console application.
        static void Main(string[] args)
        {
            // Expect input and output file paths as arguments.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: ExportToXls <inputWorkbookPath> <outputXlsPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Verify that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the source workbook.
                Workbook mergedWorkbook = new Workbook(inputPath);

                // Export to XLS.
                ExportToXls exporter = new ExportToXls();
                exporter.Export(mergedWorkbook, outputPath);

                Console.WriteLine($"Workbook exported successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
