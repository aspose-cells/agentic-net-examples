// Title: C# – Sort Excel Sheet by Date (Asc) & Amount (Desc) with Aspose.Cells DataSorter
// Description: Shows how to create a workbook, add Date and Amount columns, and apply Aspose.Cells DataSorter to sort rows first by the Date column in ascending order and then by the Amount column in descending order while keeping the header row intact.
// Keywords: Aspose.Cells | C# DataSorter | Excel multi‑column sort | sort by date ascending | sort by amount descending | .NET Excel sorting | multiple key sort Aspose | Excel workbook sorting C#
// Common Searches: Aspose.Cells sort by multiple columns C# | C# sort Excel by date then amount | DataSorter example with headers | How to sort Excel range using Aspose.Cells .NET | Sort Excel worksheet ascending date descending amount
// Developer Intent: Apply Aspose.Cells to order worksheet rows by Date ascending and Amount descending, preserving a header row.
// Use Cases: Prepare a chronological financial ledger where each day's transactions are listed with the largest amounts first. | Create a sales export file that requires date‑first ordering and amount priority for downstream analytics. | Generate a sorted dataset for pivot‑table reporting, ensuring dates are in order and amounts are ranked within each date.
// AI Prompts: Write a C# snippet using Aspose.Cells to sort a worksheet by multiple columns with custom directions and skip empty rows. | Explain how to add a third sort key (e.g., Category) to the existing DataSorter configuration. | Provide XML documentation comments for each DataSorter property used in the example. | Show how to modify the code to sort a specific named range instead of the entire sheet.

using System;
using Aspose.Cells;

// Shows how to create a workbook, add Date and Amount columns, and apply Aspose.Cells DataSorter to sort rows first by the Date column in ascending order and then by the Amount column in descending order while keeping the header row intact.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet and its cells collection
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (first row is header)
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Amount");
        cells["A2"].PutValue(new DateTime(2023, 5, 10));
        cells["B2"].PutValue(1500);
        cells["A3"].PutValue(new DateTime(2023, 4, 22));
        cells["B3"].PutValue(2000);
        cells["A4"].PutValue(new DateTime(2023, 5, 10));
        cells["B4"].PutValue(1200);
        cells["A5"].PutValue(new DateTime(2023, 3, 15));
        cells["B5"].PutValue(1800);

        // Configure the DataSorter
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true;               // First row contains column names
        sorter.Key1 = 0;                         // First key: Date column (A)
        sorter.Order1 = SortOrder.Ascending;     // Ascending order for dates
        sorter.Key2 = 1;                         // Second key: Amount column (B)
        sorter.Order2 = SortOrder.Descending;    // Descending order for amounts

        // Define the range to sort (including header row)
        int startRow = 0;
        int startColumn = 0;
        int endRow = cells.MaxDataRow;           // Last used row
        int endColumn = cells.MaxDataColumn;     // Last used column

        // Perform the sort
        sorter.Sort(cells, startRow, startColumn, endRow, endColumn);

        // Save the sorted workbook
        workbook.Save("SortedByDateAndAmount.xlsx");
    }
}
