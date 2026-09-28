// Title: How to read an Excel table's display name from a specific cell using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook, uses Cell.GetTable to locate the ListObject containing cell A1, and prints the table's DisplayName while handling a null result. | Show a safe way to obtain the display name of a worksheet table by calling Cell.GetTable in Aspose.Cells and checking for null before accessing the DisplayName property.
// Common Searches: Aspose.Cells C# get table display name from cell reference | How to use Cell.GetTable to obtain ListObject name in .NET | Retrieve Excel table DisplayName with Aspose.Cells when cell is not part of a table | C# example reading Excel table name via cell A1 using Aspose.Cells | Aspose.Cells handling null result from Cell.GetTable
// Tags: cell.GetTable retrieve ListObject .NET | extract table DisplayName Aspose.Cells | load workbook access table properties C# | null handling for Cell.GetTable Aspose.Cells | Excel table name extraction using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads an .xlsx workbook, accesses the first worksheet, obtains the ListObject that contains cell A1 via Cell.GetTable, checks for a null result, reads the table's DisplayName property, and writes it to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify by name/index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get a cell that belongs to a table (adjust the address as needed)
            Cell cell = worksheet.Cells["A1"];

            // Retrieve the table (ListObject) that contains this cell
            ListObject table = cell.GetTable();

            // If the cell is not part of any table, handle gracefully
            if (table == null)
            {
                Console.WriteLine("The specified cell does not belong to any table.");
                return;
            }

            // Read the display name of the table
            string displayName = table.DisplayName;

            // Output the display name
            Console.WriteLine($"Table display name: {displayName}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
