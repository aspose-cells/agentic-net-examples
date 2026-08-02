// Title: Sum a numeric column in an Aspose.Cells ListObject using Cell.GetTable (C#)
// Description: Creates a workbook, adds a ListObject table, retrieves it with Cell.GetTable, iterates the DataRange rows, sums a numeric column, writes the total back, and saves the file.
// Keywords: Aspose.Cells | C# | ListObject | GetTable | sum column | DataRange iteration | numeric aggregation | Excel table total | Workbook automation | Aspose.Cells example
// Common Searches: Aspose.Cells sum column using GetTable | C# iterate ListObject rows to calculate total | How to retrieve a table from a cell in Aspose.Cells | Calculate column total in Aspose.Cells ListObject | Cell.GetTable example for numeric aggregation
// Developer Intent: Calculate the total of a numeric column in a ListObject that is obtained via Cell.GetTable.
// Use Cases: Generate a summary row with the total quantity of items in an inventory table. | Aggregate financial figures from a data table before exporting a report. | Validate and total numeric entries in a user‑filled spreadsheet to ensure data integrity.
// AI Prompts: Write C# code that uses Aspose.Cells Cell.GetTable to sum a numeric column and place the result in a worksheet cell. | Explain how to safely loop through a ListObject's DataRange, skip non‑numeric cells, and compute a column total. | Show error‑handling best practices when retrieving a ListObject from any cell and performing numeric aggregation.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, adds a ListObject table, retrieves it with Cell.GetTable, iterates the DataRange rows, sums a numeric column, writes the total back, and saves the file.
class TableSumExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with a header and numeric values
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("Cherry");
            cells["B4"].PutValue(30);

            // Create a table (ListObject) that includes the data range
            int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Choose a cell that belongs to the table and retrieve the table via GetTable()
            Cell anyCellInTable = cells["B2"]; // any cell inside the table
            ListObject retrievedTable = anyCellInTable.GetTable();

            // Compute the sum of the numeric column (Quantity column, index 1)
            double sum = 0;
            int numericColumnIndex = 1; // zero‑based index within the table (B column)

            // Get the data range of the table (excludes header/footer)
            AsposeRange dataRange = retrievedTable.DataRange;

            // Iterate through each row in the data range
            int startRow = dataRange.FirstRow;
            int startCol = dataRange.FirstColumn;
            for (int r = 0; r < dataRange.RowCount; r++)
            {
                Cell cell = sheet.Cells[startRow + r, startCol + numericColumnIndex];
                if (cell.Type == CellValueType.IsNumeric)
                {
                    sum += cell.DoubleValue;
                }
            }

            // Output the result to console
            Console.WriteLine($"Sum of Quantity column: {sum}");

            // Optionally write the sum back to the worksheet
            cells["B6"].PutValue("Total");
            cells["C6"].PutValue(sum);

            // Save the workbook
            workbook.Save("TableSumResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
