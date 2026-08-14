// Title: Aspose.Cells .NET: Retrieve a ListObject with Cell.GetTable and modify numeric values using PutValue
// Description: This example creates a workbook, defines a ListObject (table) over range A1:B4, obtains the table that contains cell B2 via Cell.GetTable, updates a row value with ListObject.PutCellValue, changes the cell's numeric content with Cell.PutValue, and saves the file.
// Keywords: Aspose.Cells | Cell.GetTable | ListObject | PutCellValue | Cell.PutValue | .NET table manipulation | retrieve table from cell | update numeric cell value | C# Aspose.Cells example
// Common Searches: Aspose.Cells get ListObject from a cell | Cell.GetTable example C# | How to change a table row value with PutCellValue | Update numeric cell value using Aspose.Cells | Save workbook after modifying table Aspose
// Developer Intent: Find code that extracts the table containing a specific cell and updates numeric values both inside the table and directly in the cell.
// Use Cases: Programmatically locate the ListObject that encloses a target cell to adjust related data. | Change a column value of a specific data row within a table using ListObject.PutCellValue. | Overwrite a cell's numeric content after confirming its table membership with Cell.PutValue. | Persist modifications by saving the workbook after table and cell updates.
// AI Prompts: Generate C# code using Aspose.Cells to get the ListObject of a given cell with Cell.GetTable and then update a row's column value using ListObject.PutCellValue. | Show an Aspose.Cells .NET example that retrieves a table via Cell.GetTable, modifies a numeric cell with Cell.PutValue, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This example creates a workbook, defines a ListObject (table) over range A1:B4, obtains the table that contains cell B2 via Cell.GetTable, updates a row value with ListObject.PutCellValue, changes the cell's numeric content with Cell.PutValue, and saves the file.
    public class GetTableAndPutValueDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data that will become a table
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue(1);
            cells["B2"].PutValue(100);
            cells["A3"].PutValue(2);
            cells["B3"].PutValue(200);
            cells["A4"].PutValue(3);
            cells["B4"].PutValue(300);

            // Create a ListObject (table) covering the data range A1:B4
            int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Choose a cell that lies inside the table (e.g., B2)
            Cell cellInTable = cells["B2"];

            // Retrieve the table that contains this cell using GetTable()
            ListObject retrievedTable = cellInTable.GetTable();

            // Verify that the table was retrieved
            if (retrievedTable != null)
            {
                // Update the value of the cell at row offset 1, column offset 0 within the table
                // Row offset 1 corresponds to the second data row (row index 2 in the worksheet)
                // Column offset 0 corresponds to the first column of the table ("ID")
                retrievedTable.PutCellValue(1, 0, 999); // Set ID of second row to 999

                // Directly put a numeric value into the original cell
                cellInTable.PutValue(555); // Change B2 (Amount) to 555
            }

            // Save the workbook to verify the changes
            workbook.Save("GetTableAndPutValueDemo.xlsx");
        }
    }
}
