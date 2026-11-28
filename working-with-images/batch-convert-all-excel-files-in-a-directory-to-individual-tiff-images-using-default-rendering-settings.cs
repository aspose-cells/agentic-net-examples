// Title: Convert multiple Excel workbooks in a folder to individual multi‑page TIFF files using Aspose.Cells for .NET (C#)
// Common Searches: C# code to batch convert all .xls and .xlsx files in a directory to TIFF with Aspose.Cells | How to render each worksheet of an Excel workbook to a multi‑page TIFF using Aspose.Cells in .NET | Aspose.Cells example for converting a folder of Excel files to TIFF images with default settings | Automate Excel to TIFF conversion for multiple files using WorkbookRender in C# | Command line tool to convert Excel spreadsheets to TIFF using Aspose.Cells library
// Tags: Aspose.Cells batch conversion Excel to TIFF | WorkbookRender multi-page TIFF output | ImageOrPrintOptions TIFF rendering defaults | C# iterate Excel files in directory | convert Excel workbook to TIFF programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace BatchExcelToTiff
{
    // The C# console application reads input and output folder paths (from arguments or defaults), validates them, and scans the input directory for supported Excel extensions. For each workbook it creates a Workbook object, configures ImageOrPrintOptions with ImageType set to TIFF (using default settings), and uses WorkbookRender to generate a single multi‑page TIFF file named after the source workbook. The TIFF files are saved to the output folder, with progress logging and error handling for missing files or directory issues.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output directories.
            string inputDirectory;
            string outputDirectory;

            if (args.Length >= 2)
            {
                inputDirectory = args[0];
                outputDirectory = args[1];
            }
            else
            {
                // Default folders relative to the executable location.
                string basePath = AppDomain.CurrentDomain.BaseDirectory;
                inputDirectory = Path.Combine(basePath, "Input");
                outputDirectory = Path.Combine(basePath, "Output");
                Console.WriteLine("Usage: BatchExcelToTiff.exe <inputFolder> <outputFolder>");
                Console.WriteLine("No arguments supplied. Using default folders:");
                Console.WriteLine($"  Input : {inputDirectory}");
                Console.WriteLine($"  Output: {outputDirectory}");
            }

            // Validate input directory.
            if (!Directory.Exists(inputDirectory))
            {
                Console.WriteLine($"Error: Input directory does not exist: {inputDirectory}");
                return;
            }

            // Ensure output directory exists.
            if (!Directory.Exists(outputDirectory))
            {
                try
                {
                    Directory.CreateDirectory(outputDirectory);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create output directory '{outputDirectory}': {ex.Message}");
                    return;
                }
            }

            // Supported Excel extensions.
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };
            string[] files = Directory.GetFiles(inputDirectory);

            int processedCount = 0;
            foreach (string filePath in files)
            {
                string ext = Path.GetExtension(filePath);
                // Case‑insensitive check for supported extensions.
                if (Array.FindIndex(extensions, e => e.Equals(ext, StringComparison.OrdinalIgnoreCase)) < 0)
                {
                    // Skip non‑Excel files.
                    continue;
                }

                try
                {
                    Console.WriteLine($"Processing: {Path.GetFileName(filePath)}");

                    // Verify the file exists before loading.
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook.
                    Workbook workbook = new Workbook(filePath);

                    // Set up default image options for TIFF.
                    ImageOrPrintOptions options = new ImageOrPrintOptions
                    {
                        ImageType = Aspose.Cells.Drawing.ImageType.Tiff
                        // Other settings remain default.
                    };

                    // Render the entire workbook to a single multi‑page TIFF file.
                    WorkbookRender renderer = new WorkbookRender(workbook, options);

                    string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".tiff";
                    string outputPath = Path.Combine(outputDirectory, outputFileName);

                    // The ToImage(string) overload writes the whole workbook as a TIFF.
                    renderer.ToImage(outputPath);

                    Console.WriteLine($"Saved TIFF: {outputFileName}");
                    processedCount++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to convert '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine($"Conversion completed. {processedCount} file(s) processed.");
        }
    }
}
