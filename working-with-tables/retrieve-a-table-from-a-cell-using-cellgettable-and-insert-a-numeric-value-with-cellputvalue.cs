using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class GetTableAndPutValueDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate some data that will become a table
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Quantity");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(20);

                // Create a ListObject (table) covering the data range A1:B3
                int tableIndex = worksheet.ListObjects.Add("A1", "B3", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Choose a cell that lies inside the table (e.g., B2)
                Cell cellInTable = cells["B2"];

                // Retrieve the table that contains this cell using GetTable()
                ListObject retrievedTable = cellInTable.GetTable();

                // Verify that the table was retrieved
                if (retrievedTable != null)
                {
                    // Update a cell inside the table (row offset 1, column offset 1 => B3)
                    retrievedTable.PutCellValue(1, 1, 999);

                    // Put a value into a cell outside the original table
                    Cell targetCell = cells["C2"];
                    targetCell.PutValue(12345);
                }

                // Save the workbook to a file
                workbook.Save("GetTableAndPutValueDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            GetTableAndPutValueDemo.Run();
        }
    }
}