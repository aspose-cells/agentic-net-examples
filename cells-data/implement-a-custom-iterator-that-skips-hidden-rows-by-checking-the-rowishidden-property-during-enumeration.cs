using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in column A
        for (int i = 0; i < 10; i++)
        {
            cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Hide a few rows (zero‑based indices)
        sheet.Cells.HideRow(2); // Hide row 3
        sheet.Cells.HideRow(6); // Hide row 7

        // Use the custom iterator to enumerate only visible rows
        foreach (Row row in GetVisibleRows(sheet))
        {
            // Output the index and first cell value of each visible row
            Console.WriteLine($"Visible Row Index: {row.Index}, Value: {row[0].StringValue}");
        }

        // Save the workbook (optional)
        workbook.Save("HiddenRowsSkipped.xlsx");
    }

    // Custom iterator that yields rows whose IsHidden property is false
    static IEnumerable<Row> GetVisibleRows(Worksheet worksheet)
    {
        // Obtain the row collection enumerator
        IEnumerator enumerator = worksheet.Cells.Rows.GetEnumerator();

        // Iterate through all rows in the collection
        while (enumerator.MoveNext())
        {
            Row row = (Row)enumerator.Current;

            // Skip hidden rows
            if (row.IsHidden)
                continue;

            // Return visible row
            yield return row;
        }
    }
}