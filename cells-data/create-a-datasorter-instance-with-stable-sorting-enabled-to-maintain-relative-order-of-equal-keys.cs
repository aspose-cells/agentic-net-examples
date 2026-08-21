// Title: Stable Sorting with Aspose.Cells DataSorter (C#) – Natural Sort Order
// Description: Creates a workbook, fills it with duplicate‑key data, obtains the DataSorter, sets Order1 to SortOrder.Natural for stable sorting, defines a CellArea range, sorts the rows, and saves the file.
// Keywords: Aspose.Cells | DataSorter | stable sort | Natural sort order | C# | Excel sorting | duplicate keys | CellArea | preserve row order | example code
// Common Searches: Aspose.Cells stable sort example | DataSorter Natural order C# | how to keep duplicate key order in Excel with Aspose | stable sorting rows using Aspose.Cells | C# sort worksheet while preserving original sequence
// Developer Intent: Instantiate DataSorter, enable natural sort for stability, and sort a worksheet range in C#.
// Use Cases: Sorting a report by category while retaining the entry sequence of identical categories. | Preparing data for downstream processing where row order must remain consistent across runs. | Generating Excel exports that require duplicate keys to stay in their original order after sorting.
// AI Prompts: Show how to configure Aspose.Cells DataSorter for stable sorting with multiple keys in C#. | Provide a C# example that uses Natural sort order to sort a range that includes headers. | Explain how to verify that stable sorting preserved the original order of duplicate keys after the sort.

using System;
using Aspose.Cells;

namespace AsposeCellsStableSorterDemo
{
    // Creates a workbook, fills it with duplicate‑key data, obtains the DataSorter, sets Order1 to SortOrder.Natural for stable sorting, defines a CellArea range, sorts the rows, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including duplicate keys to demonstrate stability)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("A"); // Duplicate key
            cells["B4"].PutValue(15);
            cells["A5"].PutValue("B"); // Duplicate key
            cells["B5"].PutValue(25);
            cells["A6"].PutValue("A"); // Duplicate key
            cells["B6"].PutValue(5);

            // Obtain the DataSorter instance from the workbook
            DataSorter sorter = workbook.DataSorter;

            // Enable stable sorting by using the Natural sort order.
            // This keeps the original relative order of rows that have equal key values.
            sorter.Order1 = SortOrder.Natural;
            sorter.Key1 = 0; // Sort based on the first column (Category)

            // Define the range to sort (including headers)
            CellArea range = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 5,
                EndColumn = 1
            };

            // Perform the sort
            sorter.Sort(cells, range);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("StableSortedOutput.xlsx");
        }
    }
}
