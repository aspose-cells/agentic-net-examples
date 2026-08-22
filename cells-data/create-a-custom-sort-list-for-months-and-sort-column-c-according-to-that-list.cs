// Title: Sorting Excel rows by a custom month sequence with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells DataSorter to define a user‑specified month sequence and sort rows based on the values in column C. | Implement a sort that respects the first row as headers and uses a predefined month order in a worksheet, then save the workbook as an .xlsx file using C#.
// Common Searches: Aspose.Cells C# how to apply a custom list sort for month names in a worksheet | C# sort Excel rows by custom month order using DataSorter | example of sorting column with user defined month sequence in Aspose.Cells | sorting Excel data with headers based on a predefined month list in .NET
// Tags: DataSorter custom order sorting Aspose.Cells | header‑aware column sorting C# | Excel month order sorting .NET | user‑defined list sort for worksheet | sorting rows by month names Aspose.Cells

using System;
using Aspose.Cells;

namespace CustomMonthSortExample
{
    // The example creates a workbook, fills columns A‑C with sample data that includes month names, defines a comma‑separated month order string, configures Aspose.Cells DataSorter with header awareness, adds a sort key for column C using the custom order, sorts the populated range, and saves the result as SortedByMonth.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data with a header row
            // Column A: ID, Column B: Value, Column C: Month
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Amount");
            cells["C1"].PutValue("Month");

            // Populate some rows (unsorted months)
            cells["A2"].PutValue(1);
            cells["B2"].PutValue(100);
            cells["C2"].PutValue("March");

            cells["A3"].PutValue(2);
            cells["B3"].PutValue(150);
            cells["C3"].PutValue("January");

            cells["A4"].PutValue(3);
            cells["B4"].PutValue(120);
            cells["C4"].PutValue("December");

            cells["A5"].PutValue(4);
            cells["B5"].PutValue(130);
            cells["C5"].PutValue("July");

            // Define the custom month order list
            string monthOrder = "January,February,March,April,May,June,July,August,September,October,November,December";

            // Configure the DataSorter
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;                     // First row contains headers
            sorter.AddKey(2, SortOrder.Ascending, monthOrder); // Column C (index 2) with custom list

            // Determine the range to sort (from first row to the last used row, columns A to C)
            int lastRow = cells.MaxDataRow;
            CellArea sortArea = CellArea.CreateCellArea(0, 0, lastRow, 2);

            // Perform the sort
            sorter.Sort(cells, sortArea);

            // Save the workbook
            workbook.Save("SortedByMonth.xlsx");
        }
    }
}
