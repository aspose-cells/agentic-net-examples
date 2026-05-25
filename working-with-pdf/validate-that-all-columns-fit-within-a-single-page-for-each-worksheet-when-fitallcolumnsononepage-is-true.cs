using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class FitAllColumnsValidator
{
    static void Main()
    {
        try
        {
            // Create a new workbook and populate it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add enough columns to demonstrate fitting
            for (int col = 0; col < 50; col++)
            {
                sheet.Cells[0, col].PutValue($"Header {col + 1}");
                sheet.Cells[1, col].PutValue($"Data {col + 1}");
            }

            // Set the page setup to fit all columns on one page (optional, reinforces the option)
            sheet.PageSetup.FitToPagesWide = 1;   // 1 page wide
            sheet.PageSetup.FitToPagesTall = 0;   // height adjusts automatically

            // Create PDF save options with the required property
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = true,
                OnePagePerSheet = true   // ensures whole sheet is rendered as a single page
            };

            // Options used for rendering to evaluate page count
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                AllColumnsInOnePagePerSheet = true,
                OnePagePerSheet = true
            };

            // Validate each worksheet: after applying the option, the rendered page count must be 1
            foreach (Worksheet ws in workbook.Worksheets)
            {
                SheetRender renderer = new SheetRender(ws, renderOptions);
                int pageCount = renderer.PageCount;

                if (pageCount != 1)
                {
                    throw new InvalidOperationException(
                        $"Worksheet \"{ws.Name}\" does not fit all columns on a single page (PageCount={pageCount}).");
                }
            }

            // Save the workbook as PDF using the configured options
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, pdfOptions);
                // Write the PDF to disk
                string outputPath = "FitAllColumnsOnePage.pdf";
                File.WriteAllBytes(outputPath, ms.ToArray());
                Console.WriteLine($"PDF saved to \"{outputPath}\".");
            }

            Console.WriteLine("Workbook saved successfully. All worksheets fit columns on one page.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}