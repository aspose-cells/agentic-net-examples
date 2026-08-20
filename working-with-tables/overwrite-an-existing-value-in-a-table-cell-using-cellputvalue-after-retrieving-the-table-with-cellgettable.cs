// Title: C# – Overwrite a Table Cell with Cell.PutValue after retrieving the ListObject via Cell.GetTable (Aspose.Cells)
// Description: Shows how to create a workbook, add a ListObject, locate the table that contains a given cell using Cell.GetTable, compute the absolute cell address, replace its value with Cell.PutValue, and save the result.
// Keywords: Aspose.Cells | C# | Cell.GetTable | Cell.PutValue | ListObject | overwrite table cell | update Excel table programmatically | Aspose.Cells example | modify ListObject cell | Excel table edit .NET
// Common Searches: Aspose.Cells change value in ListObject cell | Cell.GetTable C# example | How to update Excel table cell using Aspose.Cells | Overwrite specific cell in Aspose.Cells table | Set value in table row column Aspose.Cells
// Developer Intent: Replace the value of a specific cell inside an Excel table by retrieving the table with Cell.GetTable and writing the new data with Cell.PutValue.
// Use Cases: Correct data‑entry mistakes in automatically generated reports. | Adjust price, quantity, or other numeric fields before exporting a financial workbook. | Synchronize primary‑key values after applying business‑logic transformations. | Recalculate a column value based on updated business rules within a table.
// AI Prompts: Write C# code using Aspose.Cells to locate the table containing cell C4 and set the value of the fourth data row, third column to 2500. | Explain step‑by‑step how Cell.GetTable and ListObject.StartRow/StartColumn are used to compute the absolute address of a cell that needs to be overwritten. | Provide robust error‑handling patterns for scenarios where a cell does not belong to any ListObject when using Cell.GetTable.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a ListObject, locate the table that contains a given cell using Cell.GetTable, compute the absolute cell address, replace its value with Cell.PutValue, and save the result.
    public class OverwriteTableCellDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data for the table
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Name");
                cells["A2"].PutValue(1);
                cells["B2"].PutValue("John");
                cells["A3"].PutValue(2);
                cells["B3"].PutValue("Mary");

                // Create a ListObject (table) that covers the data range A1:B3
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Choose a cell that belongs to the table (e.g., B2)
                Cell sampleCell = cells["B2"];

                // Retrieve the table that contains this cell using Cell.GetTable()
                ListObject retrievedTable = sampleCell.GetTable();

                // Verify that the table was retrieved
                if (retrievedTable == null)
                {
                    Console.WriteLine("The cell does not belong to any table.");
                    return;
                }

                // Define the row and column offsets within the table where we want to overwrite the value
                // For example, overwrite the value at row offset 1 (second data row) and column offset 0 (first column)
                int rowOffset = 1;    // corresponds to worksheet row 2 (zero‑based index)
                int columnOffset = 0; // corresponds to column A

                // Calculate the absolute cell coordinates using the table's start position
                int targetRow = retrievedTable.StartRow + rowOffset;
                int targetColumn = retrievedTable.StartColumn + columnOffset;

                // Get the target cell from the worksheet
                Cell targetCell = cells[targetRow, targetColumn];

                // Overwrite the existing value using Cell.PutValue
                targetCell.PutValue(999); // New value for the cell

                // Save the workbook (lifecycle: save)
                workbook.Save("OverwriteTableCellDemo.xlsx", SaveFormat.Xlsx);

                Console.WriteLine("Cell value overwritten and workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
