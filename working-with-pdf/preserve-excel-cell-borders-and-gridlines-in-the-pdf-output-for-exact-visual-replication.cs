using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    class PreserveBordersAndGridlines
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Header 1");
                sheet.Cells["B1"].PutValue("Header 2");
                sheet.Cells["A2"].PutValue("Data A");
                sheet.Cells["B2"].PutValue("Data B");

                // Apply a thick border to the used range to demonstrate border preservation
                Style borderStyle = workbook.CreateStyle();
                borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
                borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
                borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
                borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

                // Resolve ambiguity with System.Range by using fully‑qualified Aspose.Cells.Range
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                usedRange.SetStyle(borderStyle);

                // Ensure gridlines are visible in the worksheet
                sheet.IsGridlinesVisible = true;

                // Instruct the printer to print gridlines on the page
                sheet.PageSetup.PrintGridlines = true;

                // Configure PDF save options to include gridlines
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Use solid gridlines for exact visual match (default is solid, so this line can be omitted if not supported)
                    // GridlineType = GridlineType.Solid,
                    GridlineColor = Color.Black
                };

                // Define output file path
                string outputPath = "PreservedBordersAndGridlines.pdf";

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}