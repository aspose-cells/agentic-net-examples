// Title: Aspose.Cells C# DataSorter – Enable Stable Sorting to Preserve Duplicate Key Order
// Description: Shows how to create a Workbook, add sample data with duplicate keys, configure DataSorter with a primary column and a secondary Natural order key for stable sorting, sort the specified range, and save the result as StableSorted.xlsx.
// Keywords: Aspose.Cells | DataSorter | stable sorting | C# | .NET | preserve row order | duplicate keys | SortOrder.Natural | Excel sorting | cell range
// Common Searches: Aspose.Cells stable sort C# | DataSorter preserve original order | SortOrder.Natural example | keep duplicate rows order Aspose.Cells | C# Excel stable sorting with Aspose
// Developer Intent: Configure a DataSorter to perform a stable sort so rows with identical key values retain their original sequence.
// Use Cases: Generate reports where categories are sorted but entries within each category stay in entry order. | Export data to downstream systems that rely on the original row sequence for duplicate keys. | Create multi‑level sorts where the secondary key acts only as a stability mechanism, not a true sort. | Prepare data for pivot tables while maintaining source order for identical keys.
// AI Prompts: Show C# code that uses Aspose.Cells DataSorter with SortOrder.Natural for stable sorting. | Explain how SortOrder.Natural differs from Ascending in Aspose.Cells. | Give a step‑by‑step guide to sort an Excel range while preserving the order of rows that share the same key. | What parameters are required to enable stable sorting in Aspose.Cells DataSorter?

using Aspose.Cells;
using System;

// Shows how to create a Workbook, add sample data with duplicate keys, configure DataSorter with a primary column and a secondary Natural order key for stable sorting, sort the specified range, and save the result as StableSorted.xlsx.
class StableSortingExample
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data with duplicate keys to demonstrate stable sorting
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue("A");
        cells["B2"].PutValue(10);
        cells["A3"].PutValue("B");
        cells["B3"].PutValue(20);
        cells["A4"].PutValue("A");
        cells["B4"].PutValue(15);
        cells["A5"].PutValue("B");
        cells["B5"].PutValue(5);
        cells["A6"].PutValue("A");
        cells["B6"].PutValue(12);

        // Get the workbook's DataSorter instance
        DataSorter sorter = workbook.DataSorter;

        // Primary key: sort by Category column (ascending)
        sorter.Key1 = 0;               // Column A (zero‑based index)
        sorter.Order1 = SortOrder.Ascending;

        // Secondary key: use Natural order to keep original relative order of equal keys
        // This enables stable sorting for rows with the same Category value
        sorter.AddKey(1, SortOrder.Natural); // Column B, Natural order preserves original sequence

        // Define the range to sort (including header row)
        CellArea range = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 5,
            EndColumn = 1
        };

        // Perform the sort
        sorter.Sort(cells, range);

        // Save the workbook (saving rule)
        workbook.Save("StableSorted.xlsx");
    }
}
