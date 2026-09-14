// Title: How to sort an Excel worksheet by column D in descending order while keeping existing row groups using Aspose.Cells for .NET
// AI Prompts: Sort the full used range of a workbook by column D in descending order with Aspose.Cells DataSorter, ensuring any collapsed row groups remain intact. | Add a secondary sort key to the DataSorter so the data is first ordered by column D descending then by column A ascending, without breaking the original grouping layout. | Modify the example to treat the first row as a header and perform a descending sort on column D while preserving the row‑group hierarchy.
// Common Searches: Aspose.Cells C# sort worksheet column D descending keep row groups | DataSorter AddKey descending column example without breaking grouped rows | How to maintain Excel row grouping when sorting data with Aspose.Cells .NET
// Tags: Aspose.Cells DataSorter descending column sort | preserve Excel row grouping Aspose.Cells | C# sort worksheet by column D using DataSorter | sort entire used range without header Aspose.Cells | add secondary sort key DataSorter .NET

using System;
using Aspose.Cells;

// The sample loads an Excel file, creates a DataSorter with HasHeaders set to false, adds a descending sort key for column D (index 3), defines a CellArea covering the used range, sorts the worksheet while keeping any existing row groups intact, and saves the result.
class SortColumnDDescending
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a DataSorter instance
        DataSorter sorter = workbook.DataSorter;

        // No header row in this example (set to true if the first row contains headers)
        sorter.HasHeaders = false;

        // Add a sort key for column D (zero‑based index 3) with descending order
        sorter.AddKey(3, SortOrder.Descending);

        // Define the range to be sorted: from the first used row/column to the last used row/column
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = worksheet.Cells.MaxDataRow,
            EndColumn = worksheet.Cells.MaxDataColumn
        };

        // Perform the sort; original row grouping (if any) is preserved because the whole range is sorted based on column D
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the sorted workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
