using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace BatchExcelToTiff
{
    class Program
    {
        static void Main()
        {
            // Directory containing source Excel files
            string sourceDir = @"C:\ExcelFiles";
            // Directory where TIFF images will be saved
            string outputDir = @"C:\TiffOutput";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDir);

            // Get all Excel files (xls, xlsx, xlsm, etc.) in the source directory
            string[] excelFiles = Directory.GetFiles(sourceDir, "*.xls*");

            foreach (string excelPath in excelFiles)
            {
                try
                {
                    // Load the workbook from the Excel file
                    Workbook workbook = new Workbook(excelPath);

                    // Iterate through each worksheet in the workbook
                    for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
                    {
                        Worksheet sheet = workbook.Worksheets[sheetIndex];

                        // Configure image rendering options for TIFF output
                        ImageOrPrintOptions options = new ImageOrPrintOptions
                        {
                            ImageType = ImageType.Tiff,          // Ensure TIFF format
                            OnePagePerSheet = true               // Render the whole sheet as one page
                        };

                        // Create a SheetRender instance for the current worksheet
                        SheetRender renderer = new SheetRender(sheet, options);

                        // Build the output file name: OriginalName_Sheet0.tiff, etc.
                        string tiffFileName = $"{Path.GetFileNameWithoutExtension(excelPath)}_Sheet{sheetIndex}.tiff";
                        string tiffPath = Path.Combine(outputDir, tiffFileName);

                        // Render the worksheet to a TIFF file
                        renderer.ToTiff(tiffPath);
                    }

                    Console.WriteLine($"Successfully converted '{Path.GetFileName(excelPath)}' to TIFF images.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }
        }
    }
}