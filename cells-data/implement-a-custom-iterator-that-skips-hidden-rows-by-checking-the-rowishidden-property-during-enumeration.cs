// Title: C# extension to enumerate only visible rows in Aspose.Cells worksheets
// Description: Provides a RowExtensions.GetVisibleRows extension method that iterates a worksheet's RowCollection, checks Row.IsHidden, and yields only rows that are not hidden, enabling simple foreach loops over visible rows.
// Keywords: Aspose.Cells | C# extension method | visible rows iterator | skip hidden rows | Row.IsHidden | Worksheet row enumeration | .NET spreadsheet API
// Common Searches: Aspose.Cells iterate visible rows C# | filter hidden rows Aspose.Cells worksheet | extension method to get visible rows Aspose.Cells | skip hidden rows during row enumeration .NET | how to loop only visible rows in Aspose.Cells
// Developer Intent: Iterate through a worksheet while automatically ignoring rows that are hidden.
// Use Cases: Create reports that include only rows the user left visible after hiding data in Excel. | Apply calculations or formatting exclusively to rows that are not hidden. | Export or copy visible rows to another workbook or external data source without manual filtering.
// AI Prompts: Generate a C# extension method for Aspose.Cells that returns IEnumerable<Row> of visible rows, skipping hidden rows using Row.IsHidden. | Show how to use GetVisibleRows to calculate the sum of numeric values in the first column of visible rows only. | Modify the iterator to accept a custom predicate so rows can be filtered by additional conditions such as an empty first cell.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomIterator
{
    // Custom iterator that skips hidden rows
    // Provides a RowExtensions.GetVisibleRows extension method that iterates a worksheet's RowCollection, checks Row.IsHidden, and yields only rows that are not hidden, enabling simple foreach loops over visible rows.
    public static class RowExtensions
    {
        // Returns an enumerable of only visible rows in the given worksheet
        public static IEnumerable<Row> GetVisibleRows(this Worksheet sheet)
        {
            // Get the row collection from the worksheet
            RowCollection rows = sheet.Cells.Rows;

            // Obtain the default enumerator (iterates all existing rows)
            IEnumerator enumerator = rows.GetEnumerator();

            // Iterate through all rows
            while (enumerator.MoveNext())
            {
                Row row = (Row)enumerator.Current;

                // Skip the row if it is hidden
                if (row.IsHidden)
                    continue;

                // Yield the visible row
                yield return row;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in rows 0..5
            for (int i = 0; i < 6; i++)
            {
                cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Hide rows 2 and 4 (zero‑based indices)
            cells.HideRow(1); // Row 2
            cells.HideRow(3); // Row 4

            // Use the custom iterator to process only visible rows
            Console.WriteLine("Visible rows:");
            foreach (Row visibleRow in sheet.GetVisibleRows())
            {
                // Retrieve the first cell value of the row for demonstration
                Cell firstCell = visibleRow.FirstCell;
                string value = firstCell != null ? firstCell.StringValue : "(empty)";
                Console.WriteLine($"Row {visibleRow.Index + 1}: {value}");
            }

            // No need to save the workbook for this demonstration
        }
    }
}
