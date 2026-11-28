// Title: Export every worksheet in an Excel workbook to separate PNG images with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export each worksheet as individual PNG files | How to render Excel sheet pages to PNG using Aspose.Cells .NET | Iterate through workbook worksheets and save them as PNG images in C# | Generate PNG files for all sheets in a .xlsx file with Aspose.Cells | Sanitize worksheet names when exporting PNG images with Aspose.Cells
// Tags: Aspose.Cells SheetRender PNG per worksheet | C# export Excel worksheets as PNG images | ImageOrPrintOptions PNG configuration Aspose.Cells | safe file name generation for exported images | batch export workbook sheets to PNG

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorksheetToPngExporter
{
    // The example loads an Excel workbook, configures ImageOrPrintOptions for PNG, loops through each worksheet, creates a SheetRender for the sheet, and saves every rendered page as a PNG file using a sanitized worksheet name and page index. A helper method ensures file names are safe for the operating system.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the input workbook path.
            string inputPath;
            if (args.Length > 0)
            {
                inputPath = args[0];
            }
            else
            {
                // Default path – change as needed.
                inputPath = "input.xlsx";
            }

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook loaded successfully. Worksheets count: {workbook.Worksheets.Count}");

                // Prepare image options for PNG output.
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    SaveFormat = SaveFormat.Png,
                    // Optional: set image quality, resolution, etc.
                    // HorizontalResolution = 300,
                    // VerticalResolution = 300
                };

                // Iterate through each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Ensure the worksheet has a valid name for file naming.
                    string safeSheetName = GetSafeFileName(sheet.Name);
                    Console.WriteLine($"Processing worksheet: {sheet.Name}");

                    // Create a SheetRender for the current worksheet.
                    SheetRender sheetRender = new SheetRender(sheet, imgOptions);

                    // Render each page of the worksheet to a separate PNG file.
                    for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                    {
                        string outputFileName = $"{safeSheetName}_page{pageIndex + 1}.png";
                        sheetRender.ToImage(pageIndex, outputFileName);
                        Console.WriteLine($"Saved page {pageIndex + 1} of \"{sheet.Name}\" as \"{outputFileName}\"");
                    }
                }

                Console.WriteLine("All worksheets have been exported to PNG files.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static string GetSafeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            // Trim whitespace and limit length to avoid OS limits.
            return name.Trim().Substring(0, Math.Min(name.Length, 100));
        }
    }
}
