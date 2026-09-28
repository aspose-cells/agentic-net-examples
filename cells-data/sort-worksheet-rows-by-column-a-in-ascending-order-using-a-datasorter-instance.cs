// Title: How to sort Excel rows by the first column (A) in ascending order using Aspose.Cells DataSorter in C#
// AI Prompts: Create a DataSorter, enable HasHeaders, set Key1 to column 0 with Ascending order, define a CellArea covering A1:A5, and call sorter.Sort to reorder the rows. | Implement row sorting for an Excel worksheet by the first column using Aspose.Cells, then save the workbook to a new file.
// Common Searches: C# Aspose.Cells sort worksheet rows by column A with header row | DataSorter sort single column range ascending example in .NET | Define CellArea for sorting specific rows using Aspose.Cells
// Tags: Aspose.Cells DataSorter ascending sort | C# sort Excel column A with DataSorter | Define CellArea range for Aspose.Cells sorting | Excel worksheet row sorting using Aspose.Cells | Save sorted workbook Aspose.Cells C#

using System;
using Aspose.Cells;

namespace AsposeCellsDataSorterExample
{
    // The example creates a new workbook, fills column A with sample names, configures Aspose.Cells.DataSorter to recognize a header row, sets the first sort key to column A in ascending order, defines the sort range A1:A5 via CellArea, performs the sort, and saves the sorted workbook as SortedByColumnA.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data (including a header in the first row)
            cells["A1"].PutValue("Name");
            cells["A2"].PutValue("John");
            cells["A3"].PutValue("Alice");
            cells["A4"].PutValue("Bob");
            cells["A5"].PutValue("Charlie");

            // Get the DataSorter object from the workbook
            DataSorter sorter = workbook.DataSorter;

            // Indicate that the first row contains headers (optional)
            sorter.HasHeaders = true;

            // Set the first sort key to column A (index 0) and order to ascending
            sorter.Key1 = 0;               // Column A
            sorter.Order1 = SortOrder.Ascending;

            // Define the range to sort (from A1 to A5)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,      // Row 0 (A1)
                StartColumn = 0,   // Column 0 (A)
                EndRow = 4,        // Row 4 (A5)
                EndColumn = 0      // Only column A
            };

            // Perform the sort
            sorter.Sort(cells, sortArea);

            // Save the workbook to a file
            workbook.Save("SortedByColumnA.xlsx");
        }
    }
}
