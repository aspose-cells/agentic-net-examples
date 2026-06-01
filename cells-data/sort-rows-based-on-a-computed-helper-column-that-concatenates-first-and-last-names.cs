using System;
using Aspose.Cells;

namespace AsposeCellsSortingExample
{
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
            cells["C1"].PutValue("FullName"); // Helper column header

            // Sample data
            string[,] data = {
                { "John", "Doe" },
                { "Alice", "Smith" },
                { "Bob", "Brown" },
                { "Charlie", "Adams" }
            };

            // Populate first name and last name columns and compute the helper column
            for (int i = 0; i < data.GetLength(0); i++)
            {
                int row = i + 1; // Data starts from row 2 (index 1)
                cells[row, 0].PutValue(data[i, 0]); // FirstName
                cells[row, 1].PutValue(data[i, 1]); // LastName

                // Concatenate first and last names into the helper column (FullName)
                string fullName = $"{data[i, 0]} {data[i, 1]}";
                cells[row, 2].PutValue(fullName);
            }

            // Configure the DataSorter
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;                     // First row contains headers
            sorter.AddKey(2, SortOrder.Ascending);        // Sort by the helper column (index 2)

            // Define the range to sort (including the helper column)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = data.GetLength(0), // Last data row index
                EndColumn = 2               // Up to the helper column
            };

            // Perform the sort
            sorter.Sort(cells, sortArea);

            // (Optional) Remove the helper column after sorting
            // Shift columns left to delete column C
            cells.DeleteColumn(2);

            // Save the workbook
            workbook.Save("SortedByFullName.xlsx");
        }
    }
}