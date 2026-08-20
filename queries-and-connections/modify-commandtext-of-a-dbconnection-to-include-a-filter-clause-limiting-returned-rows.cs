// Title: Add a WHERE filter to a DBConnection command in Excel using Aspose.Cells for .NET
// Description: Demonstrates how to load an Excel workbook, locate the first DBConnection in its DataConnections collection, detect an existing WHERE clause, append or insert a filter condition (e.g., Country = 'USA'), update the Command property, and save the workbook with the modified connection.
// Keywords: Aspose.Cells DBConnection command | add WHERE clause Excel external connection | C# modify DBConnection Command | filter rows Aspose.Cells | update external data connection SQL | Aspose.Cells .NET example
// Common Searches: how to add a WHERE clause to a DBConnection in Aspose.Cells | C# update Excel DBConnection command text | append filter to external data connection Aspose.Cells | detect existing WHERE in DBConnection command | Aspose.Cells modify external connection SQL
// Developer Intent: Programmatically edit the Command string of a DBConnection to include a WHERE filter that restricts the rows returned by the query.
// Use Cases: Apply a country‑specific filter before refreshing data from a database connection. | Add extra conditions to an existing WHERE clause in an Excel DBConnection. | Automate SQL command adjustments for external data connections across multiple workbooks.
// AI Prompts: Generate C# code with Aspose.Cells that adds a WHERE clause to a DBConnection command, handling cases where a WHERE already exists. | Create a reusable method that accepts a Workbook and a filter expression, then updates the first DBConnection's Command accordingly. | Explain step‑by‑step how to inspect and modify the Command property of a DBConnection in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Demonstrates how to load an Excel workbook, locate the first DBConnection in its DataConnections collection, detect an existing WHERE clause, append or insert a filter condition (e.g., Country = 'USA'), update the Command property, and save the workbook with the modified connection.
class ModifyDbConnectionCommand
{
    static void Main()
    {
        // Load an existing workbook that contains a DBConnection
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of external data connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Find the first DBConnection in the collection
        DBConnection dbConn = null;
        foreach (ExternalConnection conn in connections)
        {
            if (conn is DBConnection dbConnection)
            {
                dbConn = dbConnection;
                break;
            }
        }

        if (dbConn == null)
        {
            Console.WriteLine("No DBConnection objects found in the workbook.");
        }
        else
        {
            // Display the original command
            Console.WriteLine("Original Command: " + dbConn.Command);

            // Define the filter clause to limit rows (example: only rows where Country = 'USA')
            string filterClause = "WHERE Country = 'USA'";

            // Determine if the existing command already contains a WHERE clause
            string originalCommand = dbConn.Command?.Trim() ?? string.Empty;
            string updatedCommand;

            if (originalCommand.IndexOf("WHERE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Append additional condition using AND
                updatedCommand = originalCommand + " AND Country = 'USA'";
            }
            else
            {
                // Append the filter clause
                // Ensure there is a space before WHERE if needed
                updatedCommand = originalCommand + (originalCommand.EndsWith(" ") ? "" : " ") + filterClause;
            }

            // Set the modified command back to the DBConnection
            dbConn.Command = updatedCommand;

            // Show the updated command
            Console.WriteLine("Updated Command: " + dbConn.Command);
        }

        // Save the workbook with the modified connection
        workbook.Save("output.xlsx");
        Console.WriteLine("Workbook saved as output.xlsx");
    }
}
