// Title: Sort Excel column M by custom priority list (High, Medium, Low) using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an XLSX file, applies a custom list sort (High → Medium → Low) to column M with Aspose.Cells DataSorter, and saves the result. | Create a reusable method in C# that accepts a workbook path and sorts any worksheet by a user‑defined priority list for a specified column using Aspose.Cells. | Extend the example to sort multiple columns while preserving the custom priority order for column M with Aspose.Cells.
// Common Searches: aspocells c# sort column with custom list high medium low | how to use DataSorter for custom text order in Excel using Aspose.Cells | c# example sorting Excel range by priority values with Aspose.Cells | custom list sorting Excel column M Aspose.Cells .NET tutorial
// Tags: custom list sorting Aspose.Cells | priority column DataSorter | Excel column M custom sort C# | Aspose.Cells custom priority order | DataSorter custom list sort

using System;
using Aspose.Cells;

// Loads input.xlsx, configures Aspose.Cells DataSorter with a custom list (High, Medium, Low) for column M, sorts the defined range, and saves the sorted workbook as output.xlsx.
class CustomPrioritySort
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the DataSorter object
        DataSorter sorter = workbook.DataSorter;

        // The data has a header row
        sorter.HasHeaders = true;

        // Column M (zero‑based index 12) will be sorted using a custom list:
        // High → Medium → Low
        sorter.AddKey(12, SortOrder.Ascending, "High,Medium,Low");

        // Determine the last used row in the worksheet
        int lastRow = worksheet.Cells.MaxDataRow;

        // Define the range to sort (from A1 to column M of the last row)
        CellArea sortArea = CellArea.CreateCellArea(0, 0, lastRow, 12);

        // Perform the sort
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the sorted workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
