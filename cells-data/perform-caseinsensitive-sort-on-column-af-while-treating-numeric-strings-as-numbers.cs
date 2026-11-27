// Title: Case‑insensitive ascending sort of column AF treating numeric strings as numbers with Aspose.Cells for .NET (C#)
// Common Searches: how to sort column AF in Excel case‑insensitively using Aspose.Cells C# | Aspose.Cells DataSorter sort numeric strings as numbers in C# | C# example for case‑insensitive sort of a specific Excel column with Aspose.Cells | sorting an entire used range by column AF while ignoring case in Aspose.Cells | Aspose.Cells ascending sort on column AF treating string numbers as numeric values
// Tags: Aspose.Cells DataSorter case-insensitive sorting | DataSorter SortAsNumber property for numeric strings | C# sort Excel column AF with Aspose.Cells | Excel worksheet range sort using Aspose.Cells API | ascending sort of Excel column treating strings as numbers

using System;
using Aspose.Cells;

// Loads an Excel workbook, configures Aspose.Cells' DataSorter for case‑insensitive sorting and numeric‑string handling, sets column AF (index 31) as the primary ascending sort key, sorts the entire used range, and saves the sorted workbook.
class Program
{
    static void Main()
    {
        // Paths to the input and output Excel files.
        // Ensure that the input file exists in the same folder as the executable
        // or provide a full path.
        string inputPath = "input.xlsx";
        string outputPath = "sorted.xlsx";

        // Load the workbook.
        Workbook workbook;
        try
        {
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook \"{inputPath}\": {ex.Message}");
            return;
        }

        // Work with the first worksheet (adjust if needed).
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the used range of the sheet.
        int maxRow = worksheet.Cells.MaxDataRow;
        int maxColumn = worksheet.Cells.MaxDataColumn;

        // Column AF is the 32nd column (zero‑based index 31).
        const int columnAfIndex = 31;

        // Verify that column AF is within the used range.
        if (columnAfIndex > maxColumn)
        {
            Console.WriteLine($"Column AF (index {columnAfIndex}) is outside the used range of the sheet.");
            return;
        }

        // Configure the DataSorter.
        DataSorter sorter = workbook.DataSorter;

        // 1. Case‑insensitive sort.
        sorter.CaseSensitive = false;

        // 2. Treat numeric strings as numbers.
        // The DataSorter class exposes a boolean property that forces numeric‑string
        // values to be interpreted as numbers during sorting.
        sorter.SortAsNumber = true;

        // 3. Sort by column AF in ascending order.
        sorter.Key1 = columnAfIndex;          // column to sort by
        sorter.Order1 = SortOrder.Ascending; // sort direction

        // Define the area to be sorted (entire used range).
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            EndRow = maxRow,
            StartColumn = 0,
            EndColumn = maxColumn
        };

        // Perform the sort.
        try
        {
            sorter.Sort(worksheet.Cells, sortArea);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Sorting failed: {ex.Message}");
            return;
        }

        // Save the sorted workbook.
        try
        {
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook sorted successfully. Saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
