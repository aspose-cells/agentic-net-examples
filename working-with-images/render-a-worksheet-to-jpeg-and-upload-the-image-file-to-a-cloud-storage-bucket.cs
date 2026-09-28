// Title: Render the first worksheet of an Excel file to a high‑resolution JPEG using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, selects the first worksheet, configures ImageOrPrintOptions for 150 dpi and one page per sheet, and saves the rendered output as a JPEG file. | Demonstrate how to modify Aspose.Cells rendering options to produce a single‑page JPEG with custom horizontal and vertical resolution and write the image to a target directory.
// Common Searches: how to convert first sheet of Excel to JPEG with Aspose.Cells C# | Aspose.Cells render worksheet as high resolution JPEG 150 dpi | C# export Excel worksheet to image using SheetRender ToImage | set ImageOrPrintOptions DPI when saving Excel as JPEG in .NET
// Tags: Aspose.Cells worksheet JPEG export | C# ImageOrPrintOptions DPI setting | SheetRender ToImage single page output | export Excel sheet as high‑resolution image .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorksheetToJpegUploader
{
    // The example loads a workbook from a specified .xlsx file, selects the first worksheet, configures ImageOrPrintOptions with 150 dpi and OnePagePerSheet enabled, uses SheetRender to render the sheet to a JPEG file named Sheet1.jpg, and saves the image in the same folder as the source workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be processed
            string excelFilePath = @"C:\Data\SampleWorkbook.xlsx";

            // Verify that the Excel file exists before attempting to load it
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine($"Error: The file '{excelFilePath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelFilePath);

                // Select the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare rendering options for JPEG output
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    HorizontalResolution = 150,
                    VerticalResolution = 150,
                    OnePagePerSheet = true
                };

                // Determine output JPEG path
                string outputPath = Path.Combine(
                    Path.GetDirectoryName(excelFilePath) ?? string.Empty,
                    "Sheet1.jpg");

                // Render the worksheet directly to a JPEG file
                SheetRender sheetRender = new SheetRender(worksheet, renderOptions);
                sheetRender.ToImage(0, outputPath);

                Console.WriteLine($"Worksheet rendered to JPEG and saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
