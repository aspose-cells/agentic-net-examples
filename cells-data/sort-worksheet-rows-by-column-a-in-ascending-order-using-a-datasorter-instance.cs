// Title: C# – Sort Worksheet Rows by Column A (Ascending) Using AspNet.Cells DataSorter
// Description: Demonstrates how to create a workbook with a header row, configure Aspose.Cells.DataSorter to treat the first row as headers, set column A as the primary sort key in ascending order, define the data range dynamically, execute the sort, and save the sorted file.
// Keywords: Aspose.Cells | DataSorter | C# sort worksheet | sort by column A | ascending Excel sort | Excel range sorting | header row preservation | CellArea | Save workbook
// Common Searches: Aspose.Cells sort rows by first column C# | C# DataSorter example with headers | How to sort an Excel sheet using Aspose.Cells | Sort worksheet range ascending column A Aspose
// Developer Intent: Sort all rows in a worksheet by the values in column A while keeping the header row intact.
// Use Cases: Generate a ranked list of names and scores before exporting a report. | Standardize record order when merging data from multiple sources. | Prepare data for downstream processing that requires alphabetical ordering.
// AI Prompts: Create C# code that sorts a worksheet by column B in descending order using Aspose.Cells DataSorter with headers. | Show how to sort on multiple columns (e.g., column A ascending then column B descending) with Aspose.Cells. | Provide an example that sorts a dynamic range where the row count is unknown at compile time.

using System;
using Aspose.Cells;

namespace AsposeCellsDataSorterExample
{
    // Demonstrates how to create a workbook with a header row, configure Aspose.Cells.DataSorter to treat the first row as headers, set column A as the primary sort key in ascending order, define the data range dynamically, execute the sort, and save the sorted file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (including a header in the first row)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Score");
            cells["A2"].PutValue("John");
            cells["B2"].PutValue(85);
            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(92);
            cells["A4"].PutValue("Bob");
            cells["B4"].PutValue(78);

            // Configure the DataSorter to sort by column A (index 0) in ascending order
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;          // First row contains headers
            sorter.Key1 = 0;                    // Column A
            sorter.Order1 = SortOrder.Ascending;

            // Define the range to sort (including headers)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = cells.MaxDataRow,       // Last row with data
                EndColumn = cells.MaxDataColumn  // Last column with data
            };

            // Perform the sort
            sorter.Sort(cells, area);

            // Save the sorted workbook
            workbook.Save("SortedByColumnA.xlsx");
        }
    }
}
