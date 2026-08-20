// Title: Change Query Table SQL Command and Refresh Data with Aspose.Cells for .NET
// Description: Loads an existing workbook, accesses its first worksheet, retrieves the first query table and its external connection, replaces the command text with a new SQL SELECT statement, sets the command type to SqlStatement, refreshes all connections to re‑run the query, and saves the updated file.
// Keywords: Aspose.Cells query table | update SQL command Excel | refresh external connection .NET | OLEDBCommandType SqlStatement | C# Excel data refresh | modify query table command
// Common Searches: how to change sql command of a query table using Aspose.Cells | refresh query table after updating command text c# | aspnet update external connection command in excel workbook | set OLEDB command type for query table Aspose.Cells
// Developer Intent: Replace the SQL command of an existing query table and re‑execute it to refresh the worksheet data.
// Use Cases: Swap a static query in a template workbook with a dynamic filter before distribution. | Automate monthly reports by adjusting the WHERE clause of a query table, refreshing, and saving the result. | Programmatically change data sources for dashboard workbooks, refresh linked queries, and export the updated file.
// AI Prompts: Generate C# code with Aspose.Cells that updates a query table's command text to a parameterized SQL statement and refreshes the workbook. | Show how to set OLEDBCommandType to SqlStatement after modifying an external connection for a query table. | Explain how to iterate over multiple query tables in a worksheet, assign different SQL commands, and refresh each one.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Loads an existing workbook, accesses its first worksheet, retrieves the first query table and its external connection, replaces the command text with a new SQL SELECT statement, sets the command type to SqlStatement, refreshes all connections to re‑run the query, and saves the updated file.
    public class UpdateQueryTableCommandDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains a query table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one query table in the worksheet
            if (worksheet.QueryTables.Count == 0)
            {
                Console.WriteLine("No query tables found in the worksheet.");
                return;
            }

            // Get the first query table
            QueryTable queryTable = worksheet.QueryTables[0];

            // Retrieve the external connection associated with the query table
            ExternalConnection extConn = queryTable.ExternalConnection;

            if (extConn == null)
            {
                Console.WriteLine("The query table does not have an associated external connection.");
                return;
            }

            // Set a new SQL command text
            string newSql = "SELECT CustomerID, CompanyName FROM Customers WHERE Country = 'USA'";
            extConn.Command = newSql;
            Console.WriteLine("Updated command text to: " + extConn.Command);

            // Ensure the command type is set to SQL statement
            extConn.CommandType = OLEDBCommandType.SqlStatement;

            // Refresh all external connections (including the query table) to fetch data with the new command
            workbook.RefreshAll();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as {outputPath}");
        }
    }
}
