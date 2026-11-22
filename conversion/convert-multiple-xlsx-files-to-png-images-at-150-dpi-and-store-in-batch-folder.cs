// Title: Batch convert every worksheet of multiple XLSX files to 150 DPI PNG images using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export each worksheet of an Excel file to PNG at 150 DPI | How to batch convert a directory of .xlsx files to high‑resolution PNG images in .NET | Aspose.Cells render Excel sheets to separate PNG files with custom DPI | Loop through multiple workbooks and save sheets as PNG using Aspose.Cells C# | Set image resolution to 150 DPI when converting Excel to PNG with Aspose.Cells
// Tags: Aspose.Cells batch XLSX to PNG conversion | ImageOrPrintOptions 150 DPI setting | SheetRender PNG rendering | individual PNG per worksheet | process multiple workbooks in folder C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The C# program scans a specified input folder for .xlsx files, loads each workbook with Aspose.Cells, configures ImageOrPrintOptions to 150 DPI, and uses SheetRender to export every worksheet as an individual PNG file into an output directory, handling missing files and rendering errors.
class Program
{
    static void Main()
    {
        // Folder containing source XLSX files
        string inputFolder = @"C:\InputXlsx";

        // Folder where PNG images will be saved
        string outputFolder = @"C:\OutputPng";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify the input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Process each XLSX file in the input folder
        foreach (string xlsxPath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            try
            {
                // Ensure the file still exists before loading
                if (!File.Exists(xlsxPath))
                {
                    Console.WriteLine($"File not found: {xlsxPath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(xlsxPath);

                // Configure image export options: PNG format at 150 DPI
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Default format is PNG; explicit setting omitted to avoid API mismatch
                    HorizontalResolution = 150,
                    VerticalResolution = 150,
                    OnePagePerSheet = true // Export each sheet as a single image
                };

                // Export each worksheet to a separate PNG file
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    try
                    {
                        Worksheet sheet = workbook.Worksheets[i];
                        SheetRender renderer = new SheetRender(sheet, imgOptions);

                        // Build output file name: OriginalFileName_SheetN.png
                        string pngPath = Path.Combine(
                            outputFolder,
                            $"{Path.GetFileNameWithoutExtension(xlsxPath)}_Sheet{i + 1}.png");

                        // Render the first (and only) page of the sheet to PNG
                        renderer.ToImage(0, pngPath);
                    }
                    catch (Exception exSheet)
                    {
                        Console.WriteLine($"Error rendering sheet {i + 1} of '{xlsxPath}': {exSheet.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{xlsxPath}': {ex.Message}");
            }
        }
    }
}
