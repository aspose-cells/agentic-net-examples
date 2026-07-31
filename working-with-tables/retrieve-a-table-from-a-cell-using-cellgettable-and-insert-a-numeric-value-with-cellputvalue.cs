// Title: Aspose.Cells C# – Get Table from a Cell with Cell.GetTable and Overwrite Value using Cell.PutValue
// Description: Shows how to create a workbook, define a ListObject (table), retrieve the table that contains a specific cell via Cell.GetTable, replace the cell's numeric value with Cell.PutValue, and save the workbook.
// Keywords: Aspose.Cells | C# | Cell.GetTable | Cell.PutValue | ListObject | table cell update | retrieve table from cell | overwrite numeric value | Aspose.Cells example | GetTable C#
// Common Searches: Aspose.Cells GetTable example C# | Cell.PutValue replace value in table | How to retrieve ListObject from a cell Aspose.Cells | Update numeric cell in Aspose.Cells table | Aspose.Cells C# get table of a cell
// Developer Intent: Locate the ListObject that contains a given cell and modify that cell’s numeric content programmatically.
// Use Cases: Access table metadata (columns, rows) by obtaining the ListObject from a cell. | Replace a single numeric entry in a table without altering its structure. | Perform batch updates after acquiring the table reference. | Refresh specific cell values in generated reports.
// AI Prompts: Generate C# code using Aspose.Cells to locate the ListObject for cell C3 and set its value to 250. | Show how to loop through all rows of a table obtained via Cell.GetTable and update the 'Quantity' column with new numbers. | Explain the difference between Cell.PutValue and ListObject.UpdateCell in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableDemo
{
    // Shows how to create a workbook, define a ListObject (table), retrieve the table that contains a specific cell via Cell.GetTable, replace the cell's numeric value with Cell.PutValue, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data that will become a table
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(20);

            // Create a ListObject (table) covering the data range A1:B3
            int tableIndex = worksheet.ListObjects.Add("A1", "B3", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Choose a cell that belongs to the table (e.g., B2)
            Cell targetCell = cells["B2"];

            // Retrieve the table that contains this cell using GetTable()
            ListObject retrievedTable = targetCell.GetTable();

            // Verify that the table was retrieved
            if (retrievedTable != null)
            {
                // Insert a new numeric value into the same cell using PutValue
                // This will overwrite the existing value (10) with 99
                targetCell.PutValue(99);
            }

            // Save the workbook to a file
            workbook.Save("TableGetAndPutValueDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
