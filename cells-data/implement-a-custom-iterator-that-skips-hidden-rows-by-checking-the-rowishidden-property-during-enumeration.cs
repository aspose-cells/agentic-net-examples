// Title: Create a C# IEnumerable<Row> to enumerate only visible rows in an Aspose.Cells worksheet
// AI Prompts: Write a C# class that implements IEnumerable<Row> and yields rows from a RowCollection only when Row.IsHidden is false. | Demonstrate using the custom visible‑row enumerable to loop through a worksheet and print the first cell value of each non‑hidden row.
// Common Searches: C# iterate visible rows Aspose.Cells worksheet | filter hidden rows RowCollection Aspose.Cells .NET | custom iterator for non‑hidden rows in Aspose.Cells | skip hidden rows when reading Excel with Aspose.Cells C#
// Tags: enumerable visible rows Aspose.Cells | filter RowCollection by IsHidden C# | skip hidden rows Aspose.Cells .NET | custom iterator RowCollection Aspose.Cells | enumerate non‑hidden rows Excel Aspose.Cells

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

// Defines a VisibleRowEnumerable class that implements IEnumerable<Row> and yields only rows where Row.IsHidden is false, hides specific rows in a worksheet, then uses the enumerable to output values from visible rows before saving the workbook.
class VisibleRowEnumerable : IEnumerable<Row>
{
    private readonly RowCollection _rows;

    public VisibleRowEnumerable(RowCollection rows)
    {
        _rows = rows;
    }

    public IEnumerator<Row> GetEnumerator()
    {
        foreach (Row row in _rows)
        {
            if (!row.IsHidden)
                yield return row;
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

        // Hide rows 3, 6 and 9 (zero‑based indices)
        worksheet.Cells.HideRow(2);
        worksheet.Cells.HideRow(5);
        worksheet.Cells.HideRow(8);

        // Iterate only over visible rows using the custom iterator
        var visibleRows = new VisibleRowEnumerable(worksheet.Cells.Rows);
        foreach (Row row in visibleRows)
        {
            // Output the value of the first cell in each visible row
            Console.WriteLine($"Visible Row {row.Index + 1}: {row[0].StringValue}");
        }

        // Save the workbook (optional)
        workbook.Save("VisibleRowsDemo.xlsx");
    }
}
