// Title: C# Parallel Worksheet‑to‑PNG Export with Aspose.Cells and TPL
// Description: This example demonstrates how to load an Excel workbook with Aspose.Cells, then use the Task Parallel Library (Parallel.ForEach) to render each worksheet and its pages to PNG files concurrently. It creates a safe file name, ensures the output directory exists, and includes per‑sheet error handling for reliable multi‑core processing.
// Keywords: Aspose.Cells | C# | Parallel.ForEach | Task Parallel Library | worksheet to PNG | Excel image export | SheetRender | multi‑core conversion | batch Excel to PNG | GitHub example
// Common Searches: export Excel worksheets to PNG in parallel C# | Aspose.Cells batch image conversion using TPL | how to render multiple Excel sheets as PNG concurrently | C# code for parallel worksheet image generation | safe file naming for Excel sheet PNG output
// Developer Intent: Generate PNG images for all worksheets of a workbook simultaneously to reduce conversion time on multi‑core machines.
// Use Cases: Create preview thumbnails for each sheet of large reports on a web server. | Automate bulk export of Excel dashboards to PNG for CI/CD pipelines. | Produce page‑by‑page PNG assets for e‑learning material from multi‑sheet workbooks.
// AI Prompts: Write C# code that uses Aspose.Cells to export every worksheet of an Excel file to separate PNG files with Parallel.ForEach. | Explain how to sanitize worksheet names for file system paths when saving PNG images with Aspose.Cells. | Suggest robust error‑handling patterns for parallel rendering of Excel sheets to PNG using the Task Parallel Library.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example demonstrates how to load an Excel workbook with Aspose.Cells, then use the Task Parallel Library (Parallel.ForEach) to render each worksheet and its pages to PNG files concurrently. It creates a safe file name, ensures the output directory exists, and includes per‑sheet error handling for reliable multi‑core processing.
public static class WorksheetToPngParallelizer
{
    // Converts each worksheet (and each of its pages) of an Excel file to PNG images in parallel.
    public static void ConvertWorksheetsToPng(string excelFilePath, string outputDirectory)
    {
        // Verify the source file exists to avoid FileNotFoundException.
        if (!File.Exists(excelFilePath))
        {
            Console.Error.WriteLine($"Error: The file \"{excelFilePath}\" does not exist.");
            return;
        }

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDirectory);

        try
        {
            // Load the workbook from the specified file.
            using (Workbook workbook = new Workbook(excelFilePath))
            {
                // Parallelize over the collection of worksheets.
                Parallel.ForEach(workbook.Worksheets, worksheet =>
                {
                    try
                    {
                        // Configure image rendering options: PNG format (default), one page per sheet.
                        ImageOrPrintOptions options = new ImageOrPrintOptions
                        {
                            // The default image format is PNG; explicit setting omitted to avoid API mismatch.
                            OnePagePerSheet = true
                        };

                        // Create a SheetRender for the current worksheet.
                        SheetRender sheetRender = new SheetRender(worksheet, options);

                        // Render each page of the worksheet to a separate PNG file.
                        for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                        {
                            // Build a safe file name using the worksheet name and page index.
                            string safeSheetName = string.Concat(worksheet.Name.Split(Path.GetInvalidFileNameChars()));
                            string outputPath = Path.Combine(outputDirectory, $"{safeSheetName}_page{pageIndex}.png");

                            // Save the rendered page to the PNG file.
                            sheetRender.ToImage(pageIndex, outputPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Failed to render worksheet \"{worksheet.Name}\": {ex.Message}");
                    }
                });
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to process workbook \"{excelFilePath}\": {ex.Message}");
        }
    }

    // Example entry point.
    public static void Main()
    {
        string sourceExcel = "input.xlsx";               // Path to the source Excel file.
        string pngOutputFolder = "RenderedPages";        // Folder where PNG files will be saved.

        ConvertWorksheetsToPng(sourceExcel, pngOutputFolder);

        Console.WriteLine("All worksheets have been rendered to PNG images (if no errors were reported).");
    }
}
