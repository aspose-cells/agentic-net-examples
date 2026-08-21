// Title: Render an Excel worksheet to JPEG without gridlines using Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, disables worksheet gridlines (IsGridlinesVisible = false), sets JPEG quality via ImageOrPrintOptions, and exports the sheet to a high‑quality JPEG with SheetRender. The workbook can be saved afterward.
// Keywords: Aspose.Cells | C# | render worksheet to JPEG | hide gridlines | ImageOrPrintOptions | SheetRender | JPEG quality | export Excel as image | .NET | save workbook
// Common Searches: Aspose.Cells render worksheet to JPEG without gridlines | hide gridlines when exporting Excel to image using Aspose.Cells | set JPEG quality in Aspose.Cells ImageOrPrintOptions | C# export Excel sheet as JPEG | save workbook after rendering image Aspose.Cells
// Developer Intent: Generate a JPEG image of a worksheet while keeping gridlines invisible.
// Use Cases: Create a clean JPEG snapshot of a report for web publishing without visible gridlines. | Email a worksheet as a high‑quality JPEG attachment, ensuring no gridlines appear. | Archive a worksheet as an image for documentation while preserving the hidden‑gridlines setting in the original workbook.
// AI Prompts: Write C# code that uses Aspose.Cells to render a worksheet to PNG with hidden gridlines and a custom DPI. | Explain how to adjust ImageOrPrintOptions to control JPEG compression level when rendering an Excel sheet with Aspose.Cells. | Show how to loop through all worksheets in a workbook and export each to a separate JPEG file while keeping gridlines hidden.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, disables worksheet gridlines (IsGridlinesVisible = false), sets JPEG quality via ImageOrPrintOptions, and exports the sheet to a high‑quality JPEG with SheetRender. The workbook can be saved afterward.
    public class RenderWorksheetToJpegWithoutGridlines
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for visible content
                worksheet.Cells["A1"].PutValue("Aspose.Cells");
                worksheet.Cells["A2"].PutValue("Rendering to JPEG");
                worksheet.Cells["B1"].PutValue(DateTime.Now);

                // Hide gridlines on the worksheet
                worksheet.IsGridlinesVisible = false;

                // Set image rendering options for JPEG output
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    // JPEG quality (0-100)
                    Quality = 90
                    // Note: ImageFormat property is optional; the file extension determines the format.
                };

                // Render the worksheet to a JPEG file
                SheetRender sheetRender = new SheetRender(worksheet, options);
                string jpegPath = "RenderedWorksheet.jpg";
                sheetRender.ToImage(0, jpegPath);
                Console.WriteLine($"Worksheet rendered to JPEG: {Path.GetFullPath(jpegPath)}");

                // Save the workbook (optional)
                string workbookPath = "WorkbookWithHiddenGridlines.xlsx";
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
            }
        }
    }
}
