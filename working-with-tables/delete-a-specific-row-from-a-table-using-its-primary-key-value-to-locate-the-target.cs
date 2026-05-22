using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsExamples
{
    public class DeleteRowByPrimaryKey
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Add a header row (ID as primary key, Name as data)
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Name");

                // Populate sample data
                cells["A2"].PutValue(1);
                cells["B2"].PutValue("Alice");
                cells["A3"].PutValue(2);
                cells["B3"].PutValue("Bob");
                cells["A4"].PutValue(3);
                cells["B4"].PutValue("Charlie");

                // Create a table (ListObject) that uses the data range
                int firstRow = 0;   // zero‑based index for header
                int firstColumn = 0;
                int totalRows = 4;  // header + 3 data rows
                int totalColumns = 2;

                int tableIndex = sheet.ListObjects.Add(
                    firstRow,
                    firstColumn,
                    firstRow + totalRows - 1,
                    firstColumn + totalColumns - 1,
                    true);

                ListObject table = sheet.ListObjects[tableIndex];
                table.ShowHeaderRow = true;

                // Primary key value to delete
                int targetId = 2;

                // Locate the row index (zero‑based) that contains the target primary key
                int rowToDelete = -1;
                for (int row = 1; row <= sheet.Cells.MaxDataRow; row++) // start after header
                {
                    // Ensure the cell contains a numeric value before comparing
                    if (cells[row, 0].Type == CellValueType.IsNumeric &&
                        cells[row, 0].IntValue == targetId)
                    {
                        rowToDelete = row;
                        break;
                    }
                }

                if (rowToDelete != -1)
                {
                    // Delete the identified row and shift cells up
                    cells.DeleteRow(rowToDelete, true);
                    Console.WriteLine($"Row with ID {targetId} deleted (zero‑based index {rowToDelete}).");
                }
                else
                {
                    Console.WriteLine($"Row with ID {targetId} not found.");
                }

                // Save the workbook to verify the result
                string outputPath = "DeleteRowByPrimaryKey.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteRowByPrimaryKey.Run();
        }
    }
}