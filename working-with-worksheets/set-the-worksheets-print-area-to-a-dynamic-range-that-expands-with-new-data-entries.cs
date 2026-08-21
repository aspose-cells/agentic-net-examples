// Title: C# – Dynamically Set Excel Print Area with Aspose.Cells Using MaxDisplayRange
// Description: Demonstrates how to programmatically set and refresh a worksheet's print area to its current MaxDisplayRange in Aspose.Cells for .NET, handling growing data and empty sheets.
// Keywords: Aspose.Cells | C# | .NET | Excel print area | MaxDisplayRange | dynamic range | worksheet print area | set print area programmatically | update print area after adding rows | Excel automation
// Common Searches: Aspose.Cells set print area to used range C# | How to update Excel print area when rows are added using Aspose.Cells | MaxDisplayRange example for print area in .NET | C# code to automatically adjust worksheet print area | Aspose.Cells dynamic print area tutorial
// Developer Intent: Programmatically define and keep the worksheet's print area aligned with the populated cells as data grows.
// Use Cases: Create periodic reports that append rows and need accurate printing without manual area changes. | Build an Excel template where users can add entries and the print layout automatically expands. | Automate batch generation of invoices where each sheet's print area must reflect the final row count. | Integrate into a data‑export service that produces printable Excel files with variable row counts.
// AI Prompts: Generate C# code that sets the print area to include only visible rows while ignoring hidden ones. | Show how to limit the print area to columns A‑D based on MaxDisplayRange. | Provide error‑handling patterns for an empty worksheet when assigning PrintArea. | Explain how to combine MaxDisplayRange with custom margins for printing. | Suggest a way to store the calculated print range in a named range for later reuse.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to programmatically set and refresh a worksheet's print area to its current MaxDisplayRange in Aspose.Cells for .NET, handling growing data and empty sheets.
class DynamicPrintAreaDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Initial data population (simulating existing data)
            for (int i = 0; i < 10; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Item {i + 1}");
                worksheet.Cells[i, 1].PutValue((i + 1) * 10);
            }

            // Set the print area based on the current used range
            SetPrintAreaToMaxDisplayRange(worksheet);

            // Add more rows later (simulating dynamic data entry)
            for (int i = 10; i < 20; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Item {i + 1}");
                worksheet.Cells[i, 1].PutValue((i + 1) * 10);
            }

            // Update the print area to include the newly added rows
            SetPrintAreaToMaxDisplayRange(worksheet);

            // Save the workbook
            string outputPath = "DynamicPrintArea.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method: sets the worksheet's print area to its MaxDisplayRange
    static void SetPrintAreaToMaxDisplayRange(Worksheet ws)
    {
        // MaxDisplayRange returns null for an empty sheet (Aspose.Cells 21.5.2+)
        var maxRange = ws.Cells.MaxDisplayRange;
        if (maxRange != null)
        {
            int startRow = maxRange.FirstRow;
            int startCol = maxRange.FirstColumn;
            int endRow = maxRange.FirstRow + maxRange.RowCount - 1;
            int endCol = maxRange.FirstColumn + maxRange.ColumnCount - 1;

            string startAddress = CellIndexToAddress(startRow, startCol);
            string endAddress   = CellIndexToAddress(endRow,   endCol);
            ws.PageSetup.PrintArea = $"{startAddress}:{endAddress}";
        }
    }

    // Converts zero‑based row/column indices to an Excel cell address (e.g., A1)
    static string CellIndexToAddress(int rowIndex, int columnIndex)
    {
        // Convert column index to letters (A, B, ..., Z, AA, AB, ...)
        string columnName = "";
        int dividend = columnIndex + 1;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        // Excel rows are 1‑based
        int rowNumber = rowIndex + 1;
        return $"{columnName}{rowNumber}";
    }
}
