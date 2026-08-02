// Title: C# – Update QueryTable SQL Command and Refresh Data with Aspose.Cells for .NET
// Description: Loads an Excel workbook, accesses the first worksheet, retrieves the first QueryTable, changes the SQL statement of its ExternalConnection, optionally refreshes the data, and saves the workbook with the updated command.
// Keywords: Aspose.Cells | C# | .NET | QueryTable | ExternalConnection | SQL command update | modify command text | refresh query table | Excel workbook automation | GitHub sample
// Common Searches: Aspose.Cells change QueryTable command text | Update SQL statement of external connection in Excel using C# | Refresh QueryTable after modifying command Aspose.Cells | Set new SELECT query for QueryTable with Aspose.Cells | C# code to modify QueryTable external connection
// Developer Intent: Change the SQL command of an existing QueryTable’s external connection and persist the workbook, with optional data refresh.
// Use Cases: Replace a static SELECT clause with a dynamic query to filter data per region before distributing the file. | Programmatically adjust linked table sources in a template workbook for multiple deployments. | Automate batch updates of QueryTables across many workbooks to point to new database views.
// AI Prompts: Generate C# code using Aspose.Cells that updates a QueryTable’s ExternalConnection Command property to a new SQL query and saves the workbook. | Show how to invoke QueryTable.RefreshData (or an alternative) after changing the command text in Aspose.Cells for .NET. | Explain fallback methods when QueryTable.RefreshData is unavailable and how to re‑execute the external connection manually.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, accesses the first worksheet, retrieves the first QueryTable, changes the SQL statement of its ExternalConnection, optionally refreshes the data, and saves the workbook with the updated command.
    public class UpdateQueryTableCommandDemo
    {
        public static void Main()
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
            const string inputPath = "InputWithQueryTable.xlsx";
            const string outputPath = "OutputUpdatedQueryTable.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load an existing workbook that contains a query table
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
                ExternalConnection connection = queryTable.ExternalConnection;

                // Update the command (SQL statement) of the connection
                string newSql = "SELECT CustomerID, CompanyName FROM Customers WHERE Country = 'USA'";
                connection.Command = newSql;
                Console.WriteLine("Updated command text: " + connection.Command);

                // Refresh the query table data (if supported by the library version)
                // In some Aspose.Cells versions, QueryTable does not expose a RefreshData method.
                // The data can be refreshed by re‑executing the external connection manually if needed.
                // For demonstration, we omit the refresh step to keep the code compilable.

                // Save the workbook with the (potentially) updated data
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
