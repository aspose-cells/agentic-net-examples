// Title: C# – Set Print Area and Export Worksheet as High‑Resolution PNG with Aspose.Cells
// Description: Creates a workbook, fills a 20 × 5 range, adds a ListObject, builds an A1‑style address, assigns it to Worksheet.PageSetup.PrintArea, configures ImageOrPrintOptions for 300 DPI PNG, and renders the first page to a PNG file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# print area | Worksheet export PNG | high resolution DPI | ImageOrPrintOptions | PageSetup.PrintArea | A1 notation | SheetRender | Aspose.Cells .NET | Excel to PNG
// Common Searches: set print area programmatically Aspose.Cells C# | export Excel sheet to PNG 300 DPI Aspose.Cells | convert zero based indices to A1 address C# | render worksheet as image Aspose.Cells .NET | how to define print area before image export
// Developer Intent: Define the worksheet's print area and generate a 300 DPI PNG snapshot of the sheet.
// Use Cases: Prepare a printable region that matches a data table before creating an image. | Produce publication‑quality PNG files for reports, presentations, or web assets. | Automate image generation after programmatic page‑setup adjustments.
// AI Prompts: Generate C# code with Aspose.Cells that sets a custom print area and saves the worksheet as a 300 DPI PNG. | Provide a utility method to convert zero‑based row/column indices to an A1‑style address for PageSetup.PrintArea. | Explain how to configure ImageOrPrintOptions for high‑resolution PNG export and handle multi‑page worksheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Tables;

// Creates a workbook, fills a 20 × 5 range, adds a ListObject, builds an A1‑style address, assigns it to Worksheet.PageSetup.PrintArea, configures ImageOrPrintOptions for 300 DPI PNG, and renders the first page to a PNG file using Aspose.Cells for .NET.
class ExportSlicerPrintArea
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (20 rows x 5 columns)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create a table (ListObject) covering the data range
            int firstRow = 0, firstCol = 0, lastRow = 19, lastCol = 4;
            int tableIdx = worksheet.ListObjects.Add(firstRow, firstCol, lastRow, lastCol, true);
            ListObject table = worksheet.ListObjects[tableIdx];
            table.DisplayName = "DataTable";

            // Build the print area string in A1 style (e.g., "A1:E20")
            string printArea = $"{CellAddress(firstCol, firstRow)}:{CellAddress(lastCol, lastRow)}";
            worksheet.PageSetup.PrintArea = printArea;

            // Configure high‑resolution image options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true,
                HorizontalResolution = 300, // DPI
                VerticalResolution = 300    // DPI
            };

            // Render the worksheet (first page) to a PNG file
            string outputPath = "SlicerPrintArea.png";
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Print area set to {printArea} and exported to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper: converts zero‑based column/row indices to an A1‑style cell address
    static string CellAddress(int columnIndex, int rowIndex)
    {
        // Convert column number to letters (0 => A)
        string columnName = "";
        int dividend = columnIndex + 1;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        // Row index is zero‑based; add 1 for A1 notation
        int rowNumber = rowIndex + 1;
        return $"{columnName}{rowNumber}";
    }
}
