using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

class VisibleRowEnumerable : IEnumerable<Row>
{
    private readonly RowCollection _rows;

    public VisibleRowEnumerable(RowCollection rows)
    {
        _rows = rows;
    }

    public IEnumerator<Row> GetEnumerator()
    {
        // Get the default row enumerator
        IEnumerator enumerator = _rows.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Row row = (Row)enumerator.Current;
            // Skip hidden rows
            if (!row.IsHidden)
            {
                yield return row;
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in column A
        for (int i = 0; i < 10; i++)
        {
            cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Hide some rows (zero‑based indices)
        cells.HideRow(1); // Hide row 2
        cells.HideRow(4); // Hide row 5
        cells.HideRow(7); // Hide row 8

        // Iterate only over visible rows using the custom iterator
        var visibleRows = new VisibleRowEnumerable(worksheet.Cells.Rows);
        foreach (Row row in visibleRows)
        {
            // Output the index (1‑based for readability) and first cell value
            Console.WriteLine($"Visible Row {row.Index + 1}: {row.FirstCell.StringValue}");
        }

        // Save the workbook (optional)
        workbook.Save("VisibleRowsDemo.xlsx");
    }
}