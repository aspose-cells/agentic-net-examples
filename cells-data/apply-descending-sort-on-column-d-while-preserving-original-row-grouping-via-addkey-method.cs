// Title: C# – Sort Column D Descending with Aspose.Cells DataSorter (Preserve Row Groups)
// Description: Loads a workbook, defines the A‑D range, sets HasHeaders, adds a descending sort key for column D via DataSorter.AddKey, sorts the range while keeping each row’s other columns aligned, and saves the sorted file.
// Keywords: Aspose.Cells | C# | DataSorter | AddKey | sort descending | column D | preserve row grouping | Excel sorting example | header row | workbook
// Common Searches: Aspose.Cells sort column D descending C# | DataSorter AddKey preserve row grouping | C# sort Excel range with header Aspose | How to sort Excel column descending using Aspose.Cells | Sort Excel sheet while keeping rows together C#
// Developer Intent: Sort column D in descending order without breaking the association of the other columns in each record.
// Use Cases: Rank sales figures by amount (column D) while keeping product details (columns A‑C) together. | Order a product catalog by price descending without separating SKU, name, and description. | Generate a leaderboard where scores (column D) are sorted high‑to‑low while preserving each participant’s full record.
// AI Prompts: Create C# code that uses Aspose.Cells DataSorter to sort column D descending and keep the rest of the row intact, with an optional header row. | Explain the role of AddKey in Aspose.Cells sorting and show how to extend the sample to sort multiple columns with mixed orders. | Show how to detect the last data row dynamically, perform the descending sort on column D, and write the workbook to a memory stream instead of a file.

using System;
using Aspose.Cells;

// Loads a workbook, defines the A‑D range, sets HasHeaders, adds a descending sort key for column D via DataSorter.AddKey, sorts the range while keeping each row’s other columns aligned, and saves the sorted file.
class SortColumnDDescending
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define the range to sort.
        // Assuming data starts at row 1 (index 0) and spans columns A to D (indices 0‑3).
        // Adjust EndRow and EndColumn as needed for your data size.
        int startRow = 0;          // first row (including header if present)
        int startColumn = 0;       // column A
        int endRow = cells.MaxDataRow;   // last row with data
        int endColumn = 3;         // column D (zero‑based index)

        // Get the DataSorter object from the workbook
        DataSorter sorter = workbook.DataSorter;

        // If the range has a header row, set this property accordingly
        sorter.HasHeaders = true; // set to false if there is no header

        // Add a sort key for column D (index 3) with descending order.
        // This preserves the original grouping of rows because only the key column is used for sorting.
        sorter.AddKey(3, SortOrder.Descending);

        // Perform the sort on the defined range
        sorter.Sort(cells, startRow, startColumn, endRow, endColumn);

        // Save the sorted workbook (replace with your desired output path)
        workbook.Save("sorted.xlsx");
    }
}
