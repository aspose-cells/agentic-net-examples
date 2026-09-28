// Title: How to sort an Excel worksheet by column E (ascending) and column F (descending) using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, enables headers, and sorts the data first by column E in ascending order then by column F in descending order using DataSorter. | Create a reusable C# method that receives a worksheet, a collection of column indices with their sort directions, and applies multi‑level sorting via Aspose.Cells DataSorter. | Adjust the sorting range to cover all populated rows in columns E‑F, then save the sorted workbook to a new file using Aspose.Cells.
// Common Searches: asp.net c# Aspose.Cells example for multi column sorting with headers | how to sort Excel data by column E ascending and column F descending using Aspose.Cells DataSorter | setting DataSorter.Key1 and Key2 for multi‑level sort in C# | Aspose.Cells sort range based on dynamic row count in .xlsx
// Tags: Aspose.Cells DataSorter multi‑level sorting | C# sort Excel columns ascending descending | Excel worksheet multi‑column sort Aspose.Cells | DataSorter sort range with headers | Aspose.Cells sort by column index

using System;
using Aspose.Cells;

// The example loads an Excel workbook, configures Aspose.Cells.DataSorter with HasHeaders=true, sets Key1 to column E (ascending) and Key2 to column F (descending), defines the sort area covering all data rows in columns E‑F, performs the sort, and saves the result to a new file.
class MultiLevelSortExample
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Configure the DataSorter for multi‑level sorting
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true;                 // Assume the first row contains headers
        sorter.Key1 = 4;                          // Column E (0‑based index)
        sorter.Order1 = SortOrder.Ascending;      // First sort: ascending
        sorter.Key2 = 5;                          // Column F (0‑based index)
        sorter.Order2 = SortOrder.Descending;     // Second sort: descending

        // Define the range to sort (all rows that contain data, columns E and F)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 4,                       // Column E
            EndRow = cells.MaxDataRow,
            EndColumn = 5                          // Column F
        };

        // Perform the sort
        sorter.Sort(cells, sortArea);

        // Save the sorted workbook (replace with your desired output path)
        workbook.Save("sorted.xlsx");
    }
}
