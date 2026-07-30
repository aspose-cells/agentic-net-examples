// Title: Overwrite a Table Cell with Cell.PutValue after Cell.GetTable in Aspose.Cells for .NET
// Description: Demonstrates how to create an Excel table (ListObject), retrieve the containing table of a specific cell using Cell.GetTable, and replace the cell's existing value with Cell.PutValue before saving the workbook.
// Keywords: Aspose.Cells | Cell.GetTable | Cell.PutValue | ListObject | C# Excel table example | overwrite table cell | retrieve table from cell | modify Excel table cell | Aspose.Cells .NET
// Common Searches: Aspose.Cells get table from cell | how to change a cell value in an Excel table using Aspose.Cells | Cell.PutValue example for ListObject | C# overwrite value in Excel table Aspose | retrieve ListObject with Cell.GetTable
// Developer Intent: Find the ListObject that contains a given cell and replace that cell’s value programmatically.
// Use Cases: Correct a data entry mistake in a specific row of an Excel table by locating the cell and applying PutValue. | Update calculated results in a numeric column of a ListObject after processing external data. | Synchronize values in an Excel table with a database by iterating rows and overwriting cells as needed.
// AI Prompts: Generate C# code that locates a cell inside a ListObject, uses Cell.GetTable to obtain its parent table, and overwrites the cell value with Cell.PutValue. | Show how to loop through all rows of a ListObject and set a new value for the 'Value' column using Cell.PutValue. | Provide an example that validates the retrieved ListObject matches the original table before modifying the cell.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to create an Excel table (ListObject), retrieve the containing table of a specific cell using Cell.GetTable, and replace the cell's existing value with Cell.PutValue before saving the workbook.
    public class OverwriteTableCellDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data that will become a table (list object)
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue(200);

                // Create a ListObject (Excel table) covering the data range A1:B3
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Retrieve a cell that belongs to the table
                Cell cellInTable = worksheet.Cells["B2"]; // This cell is inside the table

                // Get the table that contains this cell using Cell.GetTable()
                ListObject retrievedTable = cellInTable.GetTable();

                // Verify that the retrieved table is the same as the one we created
                if (retrievedTable != null && retrievedTable == table)
                {
                    // Overwrite the existing value in the cell using Cell.PutValue
                    cellInTable.PutValue(999); // New value replaces the original 100
                }

                // Save the workbook
                workbook.Save("OverwriteTableCell.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            OverwriteTableCellDemo.Run();
        }
    }
}
