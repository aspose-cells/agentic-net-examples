// Title: Export an Excel worksheet to a BMP file and display it as a thumbnail in a C# application using Aspose.Cells
// Common Searches: how to export an Excel worksheet as a BMP image with Aspose.Cells in C# | generate a bitmap thumbnail from a worksheet for a .NET UI using Aspose.Cells | render an Excel sheet to BMP with custom DPI settings in C# | save the first worksheet as a BMP file for thumbnail preview in a C# program | use SheetRender to create a BMP thumbnail of an Excel worksheet
// Tags: worksheet BMP export Aspose.Cells | ImageOrPrintOptions BMP configuration C# | SheetRender bitmap thumbnail generation | OnePagePerSheet BMP rendering option | load or create workbook Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorksheetBmpThumbnail
{
    // The example loads an existing Excel file or creates a sample workbook, selects the first worksheet, configures ImageOrPrintOptions for BMP output with 150 DPI and OnePagePerSheet, uses SheetRender to render the sheet to a BMP file (sheet_thumbnail.bmp), saves the image, and optionally opens it with the default viewer for use as a UI thumbnail.
    class Program
    {
        // Path to the source Excel file (you can change this to any existing file)
        private const string SourceExcelPath = "input.xlsx";

        // Path where the generated BMP thumbnail will be saved
        private const string ThumbnailBmpPath = "sheet_thumbnail.bmp";

        [STAThread]
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Load or create a workbook
                // ------------------------------------------------------------
                Workbook workbook;
                if (File.Exists(SourceExcelPath))
                {
                    Console.WriteLine($"Loading workbook from \"{SourceExcelPath}\"...");
                    workbook = new Workbook(SourceExcelPath);
                }
                else
                {
                    Console.WriteLine($"Source file \"{SourceExcelPath}\" not found. Creating a sample workbook...");
                    workbook = CreateSampleWorkbook();
                    workbook.Save(SourceExcelPath);
                    Console.WriteLine($"Sample workbook saved to \"{SourceExcelPath}\".");
                }

                // ------------------------------------------------------------
                // 2. Get the first worksheet (you can change the index or name)
                // ------------------------------------------------------------
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DemoSheet";

                // ------------------------------------------------------------
                // 3. Configure image rendering options for BMP output
                // ------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Bmp, // BMP format
                    OnePagePerSheet = true,                       // Render the whole sheet on one page
                    HorizontalResolution = 150,                  // Optional: set DPI
                    VerticalResolution = 150
                };

                // ------------------------------------------------------------
                // 4. Render the worksheet directly to a BMP file
                // ------------------------------------------------------------
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);
                if (sheetRender.PageCount == 0)
                {
                    throw new InvalidOperationException("The worksheet produced no renderable pages.");
                }

                // Render the first (and only) page to a BMP file
                sheetRender.ToImage(0, ThumbnailBmpPath);
                Console.WriteLine($"Worksheet rendered and saved as BMP to \"{ThumbnailBmpPath}\".");

                // ------------------------------------------------------------
                // 5. Open the generated BMP with the default image viewer (optional UI)
                // ------------------------------------------------------------
                try
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = ThumbnailBmpPath,
                        UseShellExecute = true
                    };
                    Process.Start(startInfo);
                }
                catch (Exception uiEx)
                {
                    Console.Error.WriteLine($"Unable to open the thumbnail image automatically: {uiEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }

        private static Workbook CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate a small table
            ws.Cells["A1"].PutValue("Product");
            ws.Cells["B1"].PutValue("Quantity");
            ws.Cells["C1"].PutValue("Price");
            ws.Cells["A2"].PutValue("Apple");
            ws.Cells["B2"].PutValue(120);
            ws.Cells["C2"].PutValue(0.5);
            ws.Cells["A3"].PutValue("Banana");
            ws.Cells["B3"].PutValue(85);
            ws.Cells["C3"].PutValue(0.3);
            ws.Cells["A4"].PutValue("Cherry");
            ws.Cells["B4"].PutValue(60);
            ws.Cells["C4"].PutValue(1.2);

            // Add a total formula
            ws.Cells["A5"].PutValue("Total");
            ws.Cells["B5"].Formula = "SUM(B2:B4)";
            ws.Cells["C5"].Formula = "SUM(C2:C4)";

            // Adjust column widths for better appearance
            ws.AutoFitColumns();

            return wb;
        }
    }
}
