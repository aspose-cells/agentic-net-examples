// Title: Convert a folder of XLSX workbooks to high‑resolution 150 DPI PNG images with Aspose.Cells for .NET
// Common Searches: C# code to batch convert all .xlsx files in a directory to 150 DPI PNG using Aspose.Cells | How to render each worksheet page of multiple Excel files to separate PNG images in .NET | Aspose.Cells example for exporting Excel worksheets to high‑resolution PNG files in a loop | Programmatically save Excel sheets as PNG with specific DPI in C# | Batch process Excel workbooks to PNG images with custom resolution using Aspose.Cells
// Tags: Aspose.Cells batch XLSX to PNG conversion .NET | render Excel worksheet to PNG at custom DPI | export each worksheet page as separate PNG | high‑resolution Excel to image conversion C# | directory based Excel to PNG processing Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample scans a given input folder for *.xlsx files, creates an output folder if needed, loads each workbook with Aspose.Cells, sets ImageOrPrintOptions to 150 dpi, iterates through every worksheet and its pages, and saves each page as a PNG file named with the original workbook, sheet name, and page number.
class BatchXlsxToPngConverter
{
    static void Main()
    {
        try
        {
            // Folder containing the source XLSX files
            string sourceFolder = @"C:\InputXlsx";

            // Folder where the PNG images will be saved
            string outputFolder = @"C:\OutputPng";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Get all XLSX files in the source folder
            string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                // Verify the file still exists before loading
                if (!File.Exists(xlsxPath))
                {
                    Console.WriteLine($"File not found (skipped): {xlsxPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(xlsxPath);

                    // Configure image rendering options: 150 DPI
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        HorizontalResolution = 150,
                        VerticalResolution = 150
                    };

                    // Process each worksheet in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Create a renderer for the current worksheet
                        SheetRender sheetRender = new SheetRender(sheet, imgOptions);

                        // Render each page of the worksheet (most sheets have a single page)
                        for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                        {
                            // Build the output file name: OriginalFileName_SheetName_PageIndex.png
                            string outputFileName = $"{Path.GetFileNameWithoutExtension(xlsxPath)}_{sheet.Name}_Page{pageIndex + 1}.png";
                            string outputPath = Path.Combine(outputFolder, outputFileName);

                            // Save the rendered page as PNG (format inferred from file extension)
                            sheetRender.ToImage(pageIndex, outputPath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
