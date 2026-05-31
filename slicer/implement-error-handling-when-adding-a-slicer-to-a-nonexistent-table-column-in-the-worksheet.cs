using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerErrorHandling
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for a table (2 columns)
            worksheet.Cells["A1"].PutValue("Column1");
            worksheet.Cells["B1"].PutValue("Column2");
            worksheet.Cells["A2"].PutValue("Data1");
            worksheet.Cells["B2"].PutValue("Data2");
            worksheet.Cells["A3"].PutValue("Data3");
            worksheet.Cells["B3"].PutValue("Data4");

            // Add a ListObject (table) covering the data range
            int tableIndex = worksheet.ListObjects.Add("A1", "B3", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Attempt to add a slicer for a column that does NOT exist (e.g., index 5)
            int nonExistentColumnIndex = 5; // zero‑based index, out of range for this table

            try
            {
                // Validate that the requested column index exists
                if (nonExistentColumnIndex < 0 || nonExistentColumnIndex >= table.ListColumns.Count)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(nonExistentColumnIndex),
                        $"Column index {nonExistentColumnIndex} is out of range. Table has {table.ListColumns.Count} columns.");
                }

                // If validation passes, retrieve the ListColumn and add the slicer
                ListColumn listColumn = table.ListColumns[nonExistentColumnIndex];
                SlicerCollection slicers = worksheet.Slicers;
                // Add slicer at position row 1, column 3 (cell D1)
                slicers.Add(table, listColumn, 1, 3);
                Console.WriteLine("Slicer added successfully.");
            }
            catch (ArgumentOutOfRangeException ex)
            {
                // Handle the specific case where the column does not exist
                Console.WriteLine($"Error adding slicer: {ex.Message}");
            }
            catch (Exception ex)
            {
                // General exception handling for any other unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("SlicerErrorHandlingOutput.xlsx");
        }
    }
}