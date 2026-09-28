// Title: How to sort column L in descending order while skipping hidden rows with Aspose.Cells DataSorter in C#
// AI Prompts: Write C# code that uses Aspose.Cells DataSorter to sort column L descending, ignoring rows where IsHidden = true. | Show a step‑by‑step example of collecting visible rows, sorting them, and writing them back while preserving hidden rows in an Excel workbook. | Provide a reusable method that accepts a worksheet and a column index, sorts visible rows in descending order, and leaves hidden rows untouched.
// Common Searches: Aspose.Cells C# sort column descending ignoring hidden rows | DataSorter skip hidden rows when sorting Excel sheet in .NET | How to preserve hidden rows while sorting data with Aspose.Cells | C# example sorting visible rows only using Aspose.Cells DataSorter | Sort Excel column L descending and keep hidden rows unchanged Aspose.Cells
// Tags: Aspose.Cells DataSorter sort visible rows | C# Aspose.Cells ignore hidden rows during sort | Excel column descending sort with hidden row preservation | DataSorter descending order column L | Aspose.Cells workbook hidden row handling

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDataSorterIgnoreHidden
{
    // Demonstrates creating a workbook, adding data to column L, hiding a row, extracting only visible rows, sorting them in descending order with DataSorter, copying whole rows back while preserving hidden rows, and saving the result as SortedIgnoreHidden.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column L (index 11) and some other columns
            // Row 0 (header)
            cells["L1"].PutValue("Score");
            // Visible rows
            cells["L2"].PutValue(85);
            cells["L3"].PutValue(92);
            // Hidden row (will be ignored during sorting)
            cells["L4"].PutValue(70);
            cells.Rows[3].IsHidden = true; // Row index 3 corresponds to Excel row 4
            // More visible rows
            cells["L5"].PutValue(78);
            cells["L6"].PutValue(95);

            // Add some additional data in other columns to demonstrate whole‑row movement
            for (int r = 0; r <= 5; r++)
            {
                cells[r, 0].PutValue($"Item{r}");
            }

            // -----------------------------------------------------------------
            // Configure DataSorter: sort by column L (index 11) in descending order
            // but ignore hidden rows.
            // -----------------------------------------------------------------
            DataSorter sorter = workbook.DataSorter;
            sorter.Key1 = 11;                     // Column L (zero‑based index)
            sorter.Order1 = SortOrder.Descending; // Descending order
            sorter.HasHeaders = true;             // First row is a header

            // Determine the range of rows that contain data (excluding the header)
            int startRow = 1; // data starts after header
            int endRow = cells.MaxDataRow; // last row with data

            // Collect visible rows and their values in column L
            List<(int RowIndex, object Value)> visibleRows = new List<(int, object)>();
            for (int r = startRow; r <= endRow; r++)
            {
                if (!cells.Rows[r].IsHidden)
                {
                    object val = cells[r, sorter.Key1].Value;
                    visibleRows.Add((r, val));
                }
            }

            // Sort the collected rows by the column L value in descending order
            visibleRows.Sort((a, b) =>
            {
                // Handle nulls gracefully
                if (a.Value == null && b.Value == null) return 0;
                if (a.Value == null) return 1;
                if (b.Value == null) return -1;

                // Compare as double if possible, otherwise as string
                if (double.TryParse(a.Value.ToString(), out double da) &&
                    double.TryParse(b.Value.ToString(), out double db))
                {
                    return db.CompareTo(da); // descending
                }
                return string.Compare(b.Value.ToString(), a.Value.ToString(), StringComparison.Ordinal);
            });

            // Write back the sorted visible rows to their original visible positions
            int writeRow = startRow;
            foreach (var (originalRow, _) in visibleRows)
            {
                // Skip hidden rows while writing
                while (cells.Rows[writeRow].IsHidden)
                {
                    writeRow++;
                }

                // Copy the entire row from originalRow to writeRow
                for (int c = 0; c <= cells.MaxDataColumn; c++)
                {
                    cells[writeRow, c].Copy(cells[originalRow, c]);
                }

                writeRow++;
            }

            // Save the workbook
            workbook.Save("SortedIgnoreHidden.xlsx");
        }
    }
}
