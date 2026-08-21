// Title: C# – Sort Excel rows by full name using a helper column with Aspose.Cells
// Description: Creates a workbook, adds FirstName and LastName columns, builds a temporary FullNameHelper column, sorts rows ascending with DataSorter, deletes the helper column, and saves the result.
// Keywords: Aspose.Cells sort by concatenated column | C# Excel sort full name | DataSorter helper column | temporary column sort Aspose | remove helper column after sorting | Excel row ordering Aspose.Cells | C# workbook sorting example
// Common Searches: how to sort Excel rows by full name using Aspose.Cells C# | Aspose.Cells sort by combined first and last name | C# create and delete helper column for sorting | DataSorter sort with headers Aspose.Cells | temporary column sorting Aspose.Cells .NET
// Developer Intent: Sort worksheet rows based on a concatenated first‑name/last‑name column and then clean up the temporary column.
// Use Cases: Generate an alphabetically ordered employee list without altering the original column layout. | Prepare a contact sheet where sorting must consider both first and last names while keeping the file tidy. | Apply a transient helper column for complex sorting criteria and remove it before downstream processing.
// AI Prompts: Show how to change the sort order to descending for the full‑name helper column. | Provide code that hides the helper column instead of deleting it after sorting. | Explain how to sort multiple separate ranges in the same worksheet using different helper columns with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsSortingExample
{
    // Creates a workbook, adds FirstName and LastName columns, builds a temporary FullNameHelper column, sorts rows ascending with DataSorter, deletes the helper column, and saves the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add header row
            cells["A1"].PutValue("FirstName");
            cells["B1"].PutValue("LastName");
            cells["C1"].PutValue("FullNameHelper"); // helper column header

            // Sample data (FirstName in column A, LastName in column B)
            string[,] data = {
                { "John", "Doe" },
                { "Alice", "Smith" },
                { "Bob", "Brown" },
                { "Charlie", "Adams" }
            };

            // Populate the worksheet and compute the helper column (FullNameHelper)
            for (int i = 0; i < data.GetLength(0); i++)
            {
                int row = i + 1; // data starts from row 2 (index 1)
                cells[row, 0].PutValue(data[i, 0]); // FirstName
                cells[row, 1].PutValue(data[i, 1]); // LastName

                // Concatenate first and last name and store in helper column (C)
                string fullName = $"{data[i, 0]} {data[i, 1]}";
                cells[row, 2].PutValue(fullName);
            }

            // Configure the DataSorter
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;                     // First row contains headers
            sorter.AddKey(2, SortOrder.Ascending);        // Sort by helper column (index 2)

            // Define the range to sort (including the helper column)
            int startRow = 0;
            int startColumn = 0;
            int endRow = data.GetLength(0); // last data row index (since rows are 0‑based)
            int endColumn = 2;               // include helper column

            // Perform the sort
            sorter.Sort(cells, startRow, startColumn, endRow, endColumn);

            // Optional: remove the helper column after sorting
            worksheet.Cells.DeleteColumn(2);

            // Save the workbook
            workbook.Save("SortedByFullName.xlsx");
        }
    }
}
