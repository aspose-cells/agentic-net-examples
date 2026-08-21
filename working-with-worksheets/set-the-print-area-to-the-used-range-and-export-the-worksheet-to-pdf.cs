// Title: Set Print Area to Used Range and Export to PDF using Aspose.Cells for .NET
// Description: Demonstrates how to determine a worksheet's used range with MaxDisplayRange, assign that range to PageSetup.PrintArea, and save the workbook as a PDF. The example includes error handling for empty sheets and works with C# and Aspose.Cells.
// Keywords: Aspose.Cells print area | Aspose.Cells used range | Aspose.Cells export PDF | C# set print area | Worksheet PageSetup | MaxDisplayRange C# | Aspose.Cells PDF conversion | .NET Aspose.Cells example | programmatic print area | save workbook as PDF
// Common Searches: Aspose.Cells set print area programmatically | Export worksheet to PDF with specific print area Aspose.Cells | Get used range of a worksheet Aspose.Cells .NET | PageSetup.PrintArea C# example | Save workbook as PDF Aspose.Cells | Set print area to MaxDisplayRange Aspose.Cells
// Developer Intent: Define the worksheet's print area from its used range and generate a PDF file.
// Use Cases: Create PDF reports that include only populated cells by automatically adjusting the print area. | Batch‑process multiple sheets where each sheet's print area adapts to dynamic data before PDF conversion. | Integrate into automated reporting pipelines that require precise page layout without manual configuration.
// AI Prompts: Generate C# code with Aspose.Cells that sets the print area to the worksheet's used range and saves the file as PDF. | Show how to retrieve MaxDisplayRange, assign it to PageSetup.PrintArea, and export to PDF using Aspose.Cells. | Explain safe handling of empty worksheets when configuring the print area prior to PDF export in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to determine a worksheet's used range with MaxDisplayRange, assign that range to PageSetup.PrintArea, and save the workbook as a PDF. The example includes error handling for empty sheets and works with C# and Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data (optional, just to have a used range)
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["A2"].PutValue(10);
            worksheet.Cells["B2"].PutValue(20);
            worksheet.Cells["A3"].PutValue(30);
            worksheet.Cells["B3"].PutValue(40);

            // Determine the used range of the worksheet
            AsposeRange usedRange = worksheet.Cells.MaxDisplayRange;

            // Calculate start and end cell indices
            int startRow = usedRange.FirstRow;
            int startColumn = usedRange.FirstColumn;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int endColumn = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // Convert the used range to an address string (e.g., "A1:B3")
            string startCell = CellsHelper.CellIndexToName(startRow, startColumn);
            string endCell = CellsHelper.CellIndexToName(endRow, endColumn);
            string printArea = $"{startCell}:{endCell}";

            // Set the print area to the used range
            worksheet.PageSetup.PrintArea = printArea;

            // Export the worksheet to PDF (the print area will be respected)
            workbook.Save("WorksheetPrintArea.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
