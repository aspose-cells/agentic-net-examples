// Title: How to read an Excel table's display name from a cell using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, finds the table that includes cell A1, and prints the table's name. | Show a .NET example that calls Cell.GetTable() to obtain a ListObject and retrieves its DisplayName property.
// Common Searches: Aspose.Cells C# get table name from a specific cell | Retrieve ListObject name using Cell.GetTable in .NET | How to obtain the name of an Excel table that contains a given cell with Aspose.Cells | Read Excel table metadata (display name) via Aspose.Cells API
// Tags: Cell.GetTable display name Aspose.Cells | Aspose.Cells retrieve ListObject name | C# read Excel table metadata | Aspose.Cells load workbook access table properties | Excel table DisplayName extraction .NET

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample loads 'input.xlsx', accesses the first worksheet, selects cell A1, uses Cell.GetTable() to get the containing ListObject, and prints its DisplayName if the cell belongs to a table.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Choose a cell that belongs to the table (e.g., A1)
            Cell cell = worksheet.Cells["A1"];

            // Retrieve the table (ListObject) that the cell is part of
            ListObject table = cell.GetTable();

            if (table != null)
            {
                // Read the display name of the table
                string displayName = table.DisplayName;
                Console.WriteLine("Table Display Name: " + displayName);
            }
            else
            {
                Console.WriteLine("The specified cell is not part of any table.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
