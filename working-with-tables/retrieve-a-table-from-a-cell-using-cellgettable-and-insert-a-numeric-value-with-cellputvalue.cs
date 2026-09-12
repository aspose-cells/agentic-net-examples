// Title: How to insert a numeric value into a cell and update the first data cell of its linked table using Aspose.Cells for .NET
// AI Prompts: Generate C# code that puts a number into cell A1, uses Cell.GetTable to locate any ListObject that begins at that cell, and then writes the same number into the first cell of the table's DataRange with Aspose.Cells. | Show a step‑by‑step example of inserting a numeric value into a worksheet cell, retrieving the associated table via Cell.GetTable, and updating the top‑left data cell of that table in C#.
// Common Searches: Aspose.Cells C# get ListObject from a specific cell and set value | How to use Cell.GetTable to find a table and write to its first data cell in .NET | C# example inserting numeric value into a cell and updating linked table with Aspose.Cells | Cell.PutValue followed by ListObject DataRange update Aspose.Cells
// Tags: Cell.GetTable retrieve ListObject | Cell.PutValue numeric insertion | Aspose.Cells ListObject DataRange update | C# Aspose.Cells table cell value modification | Aspose.Cells save workbook Result.xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// Demonstrates creating a workbook, inserting a numeric value into cell A1, retrieving any table that starts at that cell via Cell.GetTable, and writing the same value into the first cell of the table's DataRange before saving the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook(); // {CreateWorkbook}

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the cell at A1
            Cell cell = sheet.Cells["A1"];

            // Insert a numeric value into the cell
            cell.PutValue(12345); // {PutValue}

            // Retrieve the table (ListObject) that starts from this cell, if any
            ListObject table = cell.GetTable();

            // If a table exists, obtain its data range and write a value to its first cell
            if (table != null)
            {
                // DataRange returns an Aspose.Cells.Range object
                AsposeRange dataRange = table.DataRange;

                // Use FirstRow and FirstColumn to locate the top‑left cell of the range
                int startRow = dataRange.FirstRow;
                int startColumn = dataRange.FirstColumn;

                // Write a value to the first cell of the table's data range
                sheet.Cells[startRow, startColumn].PutValue(12345);
            }

            // Save the workbook
            workbook.Save("Result.xlsx"); // {SaveWorkbook}
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
