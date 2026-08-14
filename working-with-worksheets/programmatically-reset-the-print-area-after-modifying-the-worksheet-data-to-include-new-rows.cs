// Title: Reset worksheet print area programmatically after adding rows using Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, insert rows, determine the last used row and column with MaxDataRow/MaxDataColumn, build a range string such as A1:C20, and assign it to Worksheet.PageSetup.PrintArea so the printed area automatically expands to include the new data. Includes complete C# source and file saving.
// Keywords: Aspose.Cells | C# | .NET | reset print area | dynamic print area | Worksheet.PageSetup.PrintArea | MaxDataRow | MaxDataColumn | CellsHelper.ColumnIndexToName | update print area programmatically | Excel export | report generation
// Common Searches: Aspose.Cells set print area C# | How to update print area after inserting rows Aspose.Cells | Dynamic print area based on used range Aspose.Cells .NET | Worksheet.PageSetup.PrintArea example | Calculate last used row Aspose.Cells
// Developer Intent: The developer needs to recalculate and set the worksheet's print area so it encompasses rows added after the initial data load.
// Use Cases: Automated report generation where the number of data rows varies and the print area must cover the entire report. | Creating invoices or purchase orders with a dynamic list of line items, requiring the print area to adjust before saving or printing. | Batch exporting data tables to Excel where each sheet’s used range is unknown ahead of time and must be set as the printable area.
// AI Prompts: Provide a reusable C# method that accepts a Worksheet object and automatically resets its print area to the used range. | Explain how to construct the print area string using MaxDataRow, MaxDataColumn, and CellsHelper.ColumnIndexToName, including handling of merged cells. | Show how to update the print area after adding rows, then save the workbook as PDF or XPS with Aspose.Cells.

using System;
using Aspose.Cells;

namespace ResetPrintAreaDemo
{
    // Shows how to create a workbook, insert rows, determine the last used row and column with MaxDataRow/MaxDataColumn, build a range string such as A1:C20, and assign it to Worksheet.PageSetup.PrintArea so the printed area automatically expands to include the new data. Includes complete C# source and file saving.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add initial sample data (header row)
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["C1"].PutValue("Value");

            // Insert additional rows with data
            for (int i = 2; i <= 20; i++) // rows 2..20 (1‑based indexing)
            {
                worksheet.Cells[i - 1, 0].PutValue(i - 1);               // Column A (ID)
                worksheet.Cells[i - 1, 1].PutValue($"Item {i - 1}");    // Column B (Name)
                worksheet.Cells[i - 1, 2].PutValue((i - 1) * 10);      // Column C (Value)
            }

            // Determine the last used row and column after data insertion
            int lastRow = worksheet.Cells.MaxDataRow;       // zero‑based index
            int lastColumn = worksheet.Cells.MaxDataColumn; // zero‑based index

            // Convert column index to column name (e.g., 2 -> "C")
            string lastColumnName = CellsHelper.ColumnIndexToName(lastColumn);

            // Build the new print area string (e.g., "A1:C20")
            string newPrintArea = $"A1:{lastColumnName}{lastRow + 1}";

            // Reset the print area to cover the updated range
            worksheet.PageSetup.PrintArea = newPrintArea;

            // Save the workbook
            workbook.Save("ResetPrintAreaDemo.xlsx");

            Console.WriteLine($"Print area reset to \"{newPrintArea}\" and workbook saved.");
        }
    }
}
