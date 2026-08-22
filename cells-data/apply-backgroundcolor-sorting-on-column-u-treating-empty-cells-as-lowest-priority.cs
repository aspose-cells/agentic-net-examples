// Title: Sort Excel rows by background color in column U with Aspose.Cells for .NET, placing empty cells at the lowest priority
// AI Prompts: Generate C# code that uses Aspose.Cells DataSorter to order rows based on the fill color of column U, ensuring rows with no fill appear first. | Configure a DataSorter in Aspose.Cells to sort a worksheet range by cell color in column index 20, treating blank cells as the lowest sort priority, then save the workbook.
// Common Searches: asp.net sort excel rows by background color column U using Aspose.Cells | c# Aspose.Cells DataSorter sort on cell fill color with empty cells first | how to sort a worksheet by cell color in Aspose.Cells C# | sorting Excel sheet by column U color ascending treating no fill as lowest | Aspose.Cells sort range by cell color blank cells lowest priority
// Tags: Aspose.Cells DataSorter cell color sorting | C# sort rows by column U background color | Excel column U color based sorting .NET | blank cells lowest priority Aspose.Cells sort | DataSorter sort on cell fill color range

using System;
using System.Drawing;
using Aspose.Cells;

// The example loads an Excel workbook, creates a DataSorter with headers, adds a sort key on column U (index 20) using SortOnType.CellColor in ascending order so that cells without a fill are considered lowest, defines the full data range, performs the sort, and saves the sorted workbook.
class BackgroundColorSortExample
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a DataSorter instance
        DataSorter sorter = workbook.DataSorter;

        // Assume the first row contains headers
        sorter.HasHeaders = true;

        // Add a sort key for column U (index 20) based on cell background color.
        // Using AddKey with SortOnType.CellColor sorts by any cell color.
        // Empty cells (no fill) will be treated as the lowest priority in ascending order.
        sorter.AddKey(20, SortOnType.CellColor, SortOrder.Ascending, null);

        // Define the range to sort.
        // The range should cover all rows that need to be reordered and all columns that belong to the dataset.
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = worksheet.Cells.MaxDataRow,
            EndColumn = worksheet.Cells.MaxDataColumn
        };

        // Perform the sort operation.
        sorter.Sort(worksheet.Cells, area);

        // Save the sorted workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
