// Title: C# – Sort Excel rows by column U cell background color using Aspose.Cells (.NET) – empty cells last
// Description: A concise C# example that loads an XLSX file, creates a DataSorter, and orders rows by the fill color of cells in column U (index 20). Cells without a background are treated as lowest priority and appear after colored rows. The code defines the sort range, executes the sort, and saves the workbook.
// Keywords: Aspose.Cells background color sort | C# Excel sort by cell fill | DataSorter SortOnType.CellColor | column U color sorting .NET | empty cells last Excel sort | Aspose.Cells sort example
// Common Searches: Aspose.Cells sort rows by cell color C# | How to sort Excel column by background color using .NET | Place empty cells at the bottom when sorting by color Aspose | DataSorter sort on cell fill color example | C# sort column U by fill color Aspose.Cells
// Developer Intent: Order worksheet rows according to the background color of column U, ensuring that cells with no fill are positioned after all colored rows.
// Use Cases: Display status‑coded tasks where colored rows appear first and uncolored tasks are listed at the bottom. | Generate a priority report that groups items by highlight color in column U while keeping non‑highlighted entries last. | Prepare a color‑driven export where rows are pre‑sorted for easier visual scanning in downstream tools.
// AI Prompts: Generate C# code with Aspose.Cells to sort rows by column U background color, moving empty cells to the end. | Show how to configure DataSorter.AddKey for cell‑color sorting with empty cells treated as lowest priority. | Explain how to modify the example to sort descending and place cells without fill at the top.

using System;
using System.Drawing;
using Aspose.Cells;

// A concise C# example that loads an XLSX file, creates a DataSorter, and orders rows by the fill color of cells in column U (index 20). Cells without a background are treated as lowest priority and appear after colored rows. The code defines the sort range, executes the sort, and saves the workbook.
class BackgroundColorSortExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Column U index (A=0, B=1, ..., U=20)
        const int columnU = 20;

        // Create a DataSorter instance
        DataSorter sorter = workbook.DataSorter;

        // Assume the first row contains headers
        sorter.HasHeaders = true;

        // Add a sort key to sort by cell background color in column U (ascending)
        // Empty cells (no fill) will be placed after colored cells.
        sorter.AddKey(columnU, SortOnType.CellColor, SortOrder.Ascending, null);

        // Determine the range to sort: from the first data row to the last used row in column U
        int startRow = 0; // includes header row
        int endRow = worksheet.Cells.MaxDataRow; // last row with any data in the sheet
        int startColumn = columnU;
        int endColumn = columnU;

        // Perform the sort
        sorter.Sort(worksheet.Cells, startRow, startColumn, endRow, endColumn);

        // Save the sorted workbook
        workbook.Save("output.xlsx");
    }
}
