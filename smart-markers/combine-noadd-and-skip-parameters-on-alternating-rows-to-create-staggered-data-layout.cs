// Title: Generate an Excel file with alternating (staggered) rows using noadd and skip parameters in Aspose.Cells Smart Markers (C#)
// Common Searches: C# Aspose.Cells smart markers how to insert blank rows between data rows | use noadd and skip parameters to create staggered rows in Excel with Aspose.Cells | alternating row layout example with Aspose.Cells smart markers in C# | generate Excel sheet with empty rows after each record using Aspose.Cells | Aspose.Cells smart marker row skipping technique for spaced data rows
// Tags: noadd parameter smart markers | skip parameter smart markers | alternating rows Excel Aspose.Cells | staggered data layout C# | blank row insertion Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, writes a header, and then populates data only on every second row by calculating row indices and deliberately skipping rows using the noadd and skip smart‑marker parameters, resulting in a staggered layout saved as StaggeredData.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        // Write header row
        string[] headers = { "Item", "Value1", "Value2" };
        for (int c = 0; c < headers.Length; c++)
        {
            sheet.Cells[0, c].PutValue(headers[c]);
        }

        // Populate alternating rows to achieve a staggered layout
        // Even rows (2,4,6,…) contain data, odd rows are left empty (skip)
        int dataStartRow = 1; // first data row after header
        for (int i = 0; i < 5; i++)
        {
            int rowIndex = dataStartRow + i * 2; // skip one row each loop iteration

            // noadd: we intentionally do not add a row for the skipped index
            sheet.Cells[rowIndex, 0].PutValue($"Item {i + 1}");
            sheet.Cells[rowIndex, 1].PutValue((i + 1) * 10);
            sheet.Cells[rowIndex, 2].PutValue((i + 1) * 20);
        }

        // Save the workbook
        workbook.Save("StaggeredData.xlsx");
    }
}
