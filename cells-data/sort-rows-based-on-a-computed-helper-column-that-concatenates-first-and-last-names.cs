// Title: Sort worksheet rows by a concatenated FullName helper column with Aspose.Cells for .NET
// AI Prompts: Create a helper column that joins FirstName and LastName, then sort the worksheet rows by that column using Aspose.Cells DataSorter. | Adjust the example to perform a case‑insensitive sort on the concatenated FullName column. | Refactor the code to sort without a helper column by applying a custom sort expression that combines first and last names.
// Common Searches: aspocells c# sort rows by concatenated full name column | how to use DataSorter with a helper column in Aspose.Cells | sort Excel data by combined first and last name using Aspose.Cells .NET | sorting worksheet area with headers Aspose.Cells example
// Tags: Aspose.Cells DataSorter sort by helper column | concatenate first and last name for Excel sorting C# | sort worksheet area with headers Aspose.Cells | case‑insensitive full name sorting Aspose.Cells | custom sort expression without helper column Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSortingExample
{
    // The program builds a workbook with FirstName and LastName columns, adds a FullName helper column by concatenating them, and uses Aspose.Cells DataSorter to sort all rows ascending by the FullName column before saving the file.
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
            cells["C1"].PutValue("FullName"); // helper column

            // Sample data (FirstName in column A, LastName in column B)
            string[,] data = {
                { "John", "Doe" },
                { "Alice", "Smith" },
                { "Bob", "Brown" },
                { "Charlie", "Adams" }
            };

            // Populate the worksheet and compute the helper column (FullName)
            for (int i = 0; i < data.GetLength(0); i++)
            {
                int row = i + 1; // data starts at row 2 (index 1)
                cells[row, 0].PutValue(data[i, 0]); // FirstName
                cells[row, 1].PutValue(data[i, 1]); // LastName

                // Concatenate first and last names and store in helper column (C)
                string fullName = $"{data[i, 0]} {data[i, 1]}";
                cells[row, 2].PutValue(fullName);
            }

            // Configure the DataSorter
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;                 // first row contains headers
            sorter.AddKey(2, SortOrder.Ascending);    // sort by helper column (index 2)

            // Define the area to sort (including all three columns)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = data.GetLength(0), // last data row index
                EndColumn = 2               // up to helper column
            };

            // Perform the sort
            sorter.Sort(cells, sortArea);

            // Save the workbook
            workbook.Save("SortedByFullName.xlsx");
        }
    }
}
