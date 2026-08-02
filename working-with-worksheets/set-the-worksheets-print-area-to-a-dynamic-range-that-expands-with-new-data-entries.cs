// Title: C# – Set a Dynamic Print Area in Aspose.Cells Using the Worksheet’s Used Range
// Description: This example creates a workbook, populates the first worksheet with sample data, retrieves the worksheet’s MaxDisplayRange, converts the range to A1 notation, assigns it to PageSetup.PrintArea, and saves the file. The print area automatically expands whenever new rows or columns are added.
// Keywords: Aspose.Cells | C# | .NET | dynamic print area | worksheet print area | MaxDisplayRange | PageSetup.PrintArea | used range | auto‑expand print area | Excel export sample | GitHub example
// Common Searches: Aspose.Cells set dynamic print area C# | How to use MaxDisplayRange for print area in .NET | Automatically expand Excel print area with Aspose.Cells | PageSetup.PrintArea from used range C# | Sample code for dynamic print area Aspose.Cells
// Developer Intent: Configure the worksheet’s print area so it automatically includes all existing and future data cells.
// Use Cases: Monthly sales reports that grow as new rows are appended. | Invoices where only populated cells should be printed, avoiding blank pages. | Exported data tables that need precise pagination without manual range adjustments.
// AI Prompts: Generate C# code that sets PageSetup.PrintArea to the worksheet’s MaxDisplayRange using Aspose.Cells. | Explain how to recalculate and update the print area after adding rows to an existing Aspose.Cells workbook. | Show a step‑by‑step example of retrieving a worksheet’s used range and assigning it to PrintArea in .NET.

using System;
using Aspose.Cells;

// This example creates a workbook, populates the first worksheet with sample data, retrieves the worksheet’s MaxDisplayRange, converts the range to A1 notation, assigns it to PageSetup.PrintArea, and saves the file. The print area automatically expands whenever new rows or columns are added.
class DynamicPrintAreaDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (replace with any dynamic data source)
            for (int row = 0; row < 30; row++)
            {
                worksheet.Cells[row, 0].PutValue($"Item {row + 1}");
                worksheet.Cells[row, 1].PutValue((row + 1) * 10);
            }

            // Determine the current used range (including data, merged cells and shapes)
            Aspose.Cells.Range maxDisplayRange = worksheet.Cells.MaxDisplayRange;

            if (maxDisplayRange != null)
            {
                // Calculate the last row/column indices
                int lastRow = maxDisplayRange.FirstRow + maxDisplayRange.RowCount - 1;
                int lastColumn = maxDisplayRange.FirstColumn + maxDisplayRange.ColumnCount - 1;

                // Convert the range indices to A1 style addresses
                string startAddress = CellsHelper.CellIndexToName(maxDisplayRange.FirstRow, maxDisplayRange.FirstColumn);
                string endAddress   = CellsHelper.CellIndexToName(lastRow, lastColumn);

                // Set the print area to the determined range so it expands automatically when new data is added
                worksheet.PageSetup.PrintArea = $"{startAddress}:{endAddress}";
            }

            // Save the workbook with the dynamic print area applied
            workbook.Save("DynamicPrintArea.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
