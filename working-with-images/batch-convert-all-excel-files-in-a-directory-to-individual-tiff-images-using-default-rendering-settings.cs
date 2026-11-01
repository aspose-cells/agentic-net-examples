// Title: Batch convert every Excel workbook in a directory to a multi‑page TIFF image using Aspose.Cells for .NET
// AI Prompts: Write C# code that scans a folder for .xls, .xlsx, .xlsm, and .xlsb files and saves each workbook as a multi‑page TIFF using Aspose.Cells default settings. | Modify the script to walk through subdirectories recursively and let the user define TIFF resolution and compression parameters. | Implement try‑catch logic that logs files failing to load and continues processing the remaining Excel workbooks.
// Common Searches: Aspose.Cells C# convert all Excel files in a folder to TIFF images | export each worksheet of an Excel workbook as a page in a TIFF file using .NET | how to save Excel workbook as multi‑page TIFF with default rendering in C# | C# code example for converting directory of .xlsx files to TIFF with Aspose.Cells
// Tags: batch Excel to TIFF conversion Aspose.Cells | save workbook as multi‑page TIFF C# | enumerate Excel files in a directory programmatically | default rendering settings TIFF export Aspose.Cells | Aspose.Cells SaveFormat.Tiff example | non‑recursive file processing C#

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToTiffBatch
{
    // The program scans a specified input folder for .xls, .xlsx, .xlsm, and .xlsb files, loads each workbook with Aspose.Cells, and saves it as a multi‑page TIFF image in an output folder using the library’s default rendering settings.
    class Program
    {
        static void Main(string[] args)
        {
            // Input directory containing Excel files
            string inputFolder = @"C:\ExcelFiles";

            // Output directory for generated TIFF images
            string outputFolder = @"C:\TiffOutput";

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Supported Excel extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

            // Get all Excel files in the input folder (non‑recursive)
            var excelFiles = Directory.GetFiles(inputFolder)
                                      .Where(f => extensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));

            foreach (var filePath in excelFiles)
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(filePath);

                // Build the output TIFF file name (same base name, .tiff extension)
                string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".tiff";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the workbook as a TIFF image using default rendering settings
                // Each worksheet will be rendered as a separate page in the multi‑page TIFF
                workbook.Save(outputPath, SaveFormat.Tiff);

                Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to '{outputFileName}'.");
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
