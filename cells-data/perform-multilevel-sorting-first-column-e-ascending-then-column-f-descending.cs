// Title: C# – Multi‑Level Sort of Columns E & F with Aspose.Cells DataSorter
// Description: Shows how to build a workbook, insert headers and sample rows into columns E and F, set up Aspose.Cells DataSorter for an ascending primary key on column E and a descending secondary key on column F, define the target range, run the sort, and write the result to MultiLevelSorted.xlsx.
// Keywords: Aspose.Cells | C# DataSorter | multi‑level sort | Excel column sorting | ascending then descending | sort with headers | column E sort | column F sort | .NET Excel example | GitHub sample code
// Common Searches: Aspose.Cells sort multiple columns C# | DataSorter sort by two columns .NET | C# sort Excel column E ascending then F descending | How to use Aspose.Cells DataSorter with headers | Multi‑key sorting example Aspose.Cells
// Developer Intent: Apply a two‑level sort on a worksheet: first order rows by column E in ascending order, then order rows with the same E value by column F in descending order.
// Use Cases: Generate a category report where items are grouped alphabetically (column E) and the highest sales figures appear first within each group (column F). | Create a ranked employee list sorted by department (column E) and then by performance score from highest to lowest (column F). | Prepare data for downstream analytics where a string key must be sorted alphabetically and a numeric key must be sorted in reverse order. | Export a cleanly ordered CSV from Excel where primary and secondary keys are defined by different columns.
// AI Prompts: Extend the example to add a third sort key on column G in ascending order and show the required property settings. | Write code that automatically detects the last used row in columns E‑F and builds the CellArea range without hard‑coding the end row. | Provide a version of the snippet that sorts a range without headers, using Aspose.Cells DataSorter.

using System;
using Aspose.Cells;

namespace MultiLevelSortExample
{
    // Shows how to build a workbook, insert headers and sample rows into columns E and F, set up Aspose.Cells DataSorter for an ascending primary key on column E and a descending secondary key on column F, define the target range, run the sort, and write the result to MultiLevelSorted.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data (including headers in row 0)
            cells["E1"].PutValue("Category");   // Header for column E
            cells["F1"].PutValue("Value");      // Header for column F

            cells["E2"].PutValue("B");
            cells["F2"].PutValue(30);
            cells["E3"].PutValue("A");
            cells["F3"].PutValue(20);
            cells["E4"].PutValue("B");
            cells["F4"].PutValue(10);
            cells["E5"].PutValue("A");
            cells["F5"].PutValue(40);

            // Configure the DataSorter for multi‑level sorting
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;                 // First row contains headers
            sorter.Key1 = 4;                          // Column E (index 4) – first sort key
            sorter.Order1 = SortOrder.Ascending;      // Ascending order for column E
            sorter.Key2 = 5;                          // Column F (index 5) – second sort key
            sorter.Order2 = SortOrder.Descending;     // Descending order for column F

            // Define the range to sort (including headers)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,          // Header row
                StartColumn = 4,       // Column E
                EndRow = 5,            // Last data row (row 5, zero‑based index)
                EndColumn = 5          // Column F
            };

            // Perform the sort
            sorter.Sort(cells, sortArea);

            // Save the workbook to verify the result
            workbook.Save("MultiLevelSorted.xlsx");
        }
    }
}
