// Title: C# – Set Print Area to Used Range and Export Worksheet to PDF with Aspose.Cells
// Description: Shows how to programmatically define the print area from a worksheet's used range using Aspose.Cells for .NET, then save the sheet as a PDF. The sample creates data, obtains MaxDisplayRange, builds an A1‑style address, assigns it to PageSetup.PrintArea, and calls Workbook.Save with SaveFormat.Pdf.
// Keywords: Aspose.Cells | C# | .NET | set print area | used range | MaxDisplayRange | export to PDF | PageSetup.PrintArea | worksheet PDF conversion | Aspose.Cells PDF export | programmatic print area
// Common Searches: Aspose.Cells set print area to used range | C# export worksheet to PDF with custom print area | How to use MaxDisplayRange for printing in Aspose.Cells | Define print area before PDF conversion Aspose.Cells .NET | Save Aspose.Cells workbook as PDF with automatic print area | Aspose.Cells PageSetup.PrintArea example
// Developer Intent: Define the worksheet's print area based on its populated cells and generate a PDF file.
// Use Cases: Automatically adjust the printable region for dynamically sized reports before PDF creation. | Create PDF invoices that include only filled cells, eliminating blank pages. | Batch‑convert worksheets with varying data volumes to PDFs without manual range specification.
// AI Prompts: Provide C# code that sets the print area to the worksheet's used range and saves the workbook as a PDF using Aspose.Cells. | Explain the difference between MaxDisplayRange and MaxCell in Aspose.Cells and when each should be used for print area configuration. | Show how to set custom print areas for multiple worksheets and export each to a separate PDF file with Aspose.Cells. | Generate a script that iterates through all worksheets, applies the used‑range print area, and creates individual PDF files. | Describe how merged cells affect the calculation of the print area and how to handle them in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to programmatically define the print area from a worksheet's used range using Aspose.Cells for .NET, then save the sheet as a PDF. The sample creates data, obtains MaxDisplayRange, builds an A1‑style address, assigns it to PageSetup.PrintArea, and calls Workbook.Save with SaveFormat.Pdf.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // (Optional) Populate some sample data so that the used range is not empty
            // ------------------------------------------------------------
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["A2"].PutValue("Data1");
            worksheet.Cells["B2"].PutValue(123);
            worksheet.Cells["A3"].PutValue("Data2");
            worksheet.Cells["B3"].PutValue(456);

            // ------------------------------------------------------------
            // Set the print area to the used range of the worksheet
            // ------------------------------------------------------------
            // Get the used range as an Aspose.Cells.Range object
            Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

            // Determine start and end cell coordinates
            int startRow = usedRange.FirstRow;
            int startColumn = usedRange.FirstColumn;
            int endRow = startRow + usedRange.RowCount - 1;
            int endColumn = startColumn + usedRange.ColumnCount - 1;

            // Convert the start and end cells to their A1 style names
            string startCell = worksheet.Cells[startRow, startColumn].Name;
            string endCell = worksheet.Cells[endRow, endColumn].Name;

            // Assign the print area using the A1 style address (e.g., "A1:B3")
            worksheet.PageSetup.PrintArea = $"{startCell}:{endCell}";

            // ------------------------------------------------------------
            // Export the worksheet (with the defined print area) to PDF
            // ------------------------------------------------------------
            workbook.Save("Worksheet.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
