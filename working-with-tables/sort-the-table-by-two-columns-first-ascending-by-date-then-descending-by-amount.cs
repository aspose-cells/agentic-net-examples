// Title: C# – Sort Excel Table by Date (asc) and Amount (desc) with Aspose.Cells
// Description: Shows how to load a workbook, configure Aspose.Cells DataSorter with header awareness, set column A (Date) to ascending order, column B (Amount) to descending order, define the data range, execute the sort, and save the sorted file.
// Keywords: Aspose.Cells | C# Excel sort | DataSorter | multi‑column sort | date ascending | amount descending | .NET Excel sorting | header aware sort | CellArea range
// Common Searches: Aspose.Cells sort by date then amount C# | C# DataSorter multiple columns example | How to sort Excel worksheet with headers using Aspose.Cells | Sort Excel range ascending descending Aspose | Multi‑key sort Aspose.Cells .NET
// Developer Intent: Implement a multi‑key sort on an Excel worksheet where rows are ordered first by the Date column in ascending order and then by the Amount column in descending order.
// Use Cases: Generate chronological transaction reports with the highest amounts listed first for each day. | Prepare sales data for pivot‑table analysis by ordering dates earliest‑to‑latest and amounts highest‑to‑lowest within each date. | Reorder imported financial records before feeding them to an ERP system to satisfy business sorting rules. | Create sorted export files for regulatory filing where date order is mandatory and amounts need descending priority.
// AI Prompts: Write C# code using Aspose.Cells to sort a worksheet by column A ascending and column B descending, keeping the header row intact. | Explain how to add a third sort key with a custom order in Aspose.Cells DataSorter. | Show how to limit the sort to a specific range such as B2:D150 while sorting by date and amount. | Provide a step‑by‑step guide to sort large Excel files efficiently with Aspose.Cells. | Demonstrate how to sort dates stored as text using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to load a workbook, configure Aspose.Cells DataSorter with header awareness, set column A (Date) to ascending order, column B (Amount) to descending order, define the data range, execute the sort, and save the sorted file.
class SortTableByDateAndAmount
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("Input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the DataSorter object
        DataSorter sorter = workbook.DataSorter;

        // Assume the first row contains headers
        sorter.HasHeaders = true;

        // First sort key: Date column (A) ascending
        sorter.Key1 = 0; // Column A index
        sorter.Order1 = SortOrder.Ascending;

        // Second sort key: Amount column (B) descending
        sorter.Key2 = 1; // Column B index
        sorter.Order2 = SortOrder.Descending;

        // Define the range to sort (including headers)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = worksheet.Cells.MaxDataRow,
            EndColumn = worksheet.Cells.MaxDataColumn
        };

        // Perform the sort
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the sorted workbook
        workbook.Save("SortedOutput.xlsx");
    }
}
