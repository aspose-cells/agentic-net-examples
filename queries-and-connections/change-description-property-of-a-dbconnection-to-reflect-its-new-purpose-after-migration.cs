// Title: C# – Update DBConnection ConnectionDescription after migration with Aspose.Cells for .NET
// Description: This example demonstrates how to load an Excel workbook, loop through its external DataConnections, detect DBConnection objects, assign a new ConnectionDescription that reflects the post‑migration purpose, output the change for verification, and save the file with the updated metadata using Aspose.Cells for .NET.
// Keywords: Aspose.Cells DBConnection description | C# update ConnectionDescription | Excel external data connection .NET | migrate DBConnection metadata | set workbook DataConnections description
// Common Searches: change DBConnection description Aspose.Cells C# | update ConnectionDescription after database migration | modify external connection metadata in Excel using .NET | Aspose.Cells set description for DataConnections | C# code to rename DBConnection purpose
// Developer Intent: Replace the existing ConnectionDescription of each DBConnection in a workbook with a string that describes its new role after a database migration.
// Use Cases: Standardize connection descriptions across corporate Excel reports after moving to a new reporting database. | Automate batch updates of workbooks to inform end‑users about revised data sources. | Validate the change by logging the new description before persisting the workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to change the ConnectionDescription of all DBConnection objects in an Excel file to a user‑provided value. | Show how to log both the original and updated ConnectionDescription when modifying a DBConnection. | Explain best practices for handling workbooks that contain no DBConnection objects to prevent runtime exceptions.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// This example demonstrates how to load an Excel workbook, loop through its external DataConnections, detect DBConnection objects, assign a new ConnectionDescription that reflects the post‑migration purpose, output the change for verification, and save the file with the updated metadata using Aspose.Cells for .NET.
class UpdateDbConnectionDescription
{
    static void Main()
    {
        // Load the workbook that already contains a DBConnection
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all external connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Check if the connection is a DBConnection
            if (connection is DBConnection dbConnection)
            {
                // Set a new description that reflects the connection's new purpose after migration
                dbConnection.ConnectionDescription = "Migrated connection for reporting purposes";

                // Optional: output the updated description to the console for verification
                Console.WriteLine($"Updated ConnectionDescription: {dbConnection.ConnectionDescription}");
            }
        }

        // Save the workbook with the modified connection description
        workbook.Save("output.xlsx");
    }
}
