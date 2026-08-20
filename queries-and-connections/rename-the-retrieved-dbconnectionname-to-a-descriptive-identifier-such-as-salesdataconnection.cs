// Title: Rename a DBConnection in an Excel workbook using Aspose.Cells for .NET (C#)
// Description: Load an existing workbook, iterate its DataConnections, locate DBConnection objects, assign a meaningful Name (e.g., SalesDataConnection), and save the updated file.
// Keywords: Aspose.Cells rename DBConnection | C# change Excel data connection name | set DBConnection.Name property | update external connection identifier .NET | programmatic Excel connection rename | Aspose.Cells ExternalConnection C# | Excel workbook data connections
// Common Searches: Aspose.Cells rename DBConnection C# | How to change Excel DB connection name programmatically | Set name of external connection in .xlsx using Aspose | Rename database connection in workbook with .NET
// Developer Intent: Modify the DBConnection.Name property to a descriptive identifier such as SalesDataConnection.
// Use Cases: Enforce a naming convention for data connections across a suite of reports before distribution. | Replace generic connection names with business‑specific identifiers after consolidating data sources. | Integrate connection‑renaming into CI/CD pipelines that generate Excel dashboards automatically.
// AI Prompts: Create C# code that prefixes each DBConnection name with its worksheet title using Aspose.Cells. | Write a method to verify the existence of a DBConnection with a specific name before renaming it. | Show how to log original and new DBConnection names to a text file while processing a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Load an existing workbook, iterate its DataConnections, locate DBConnection objects, assign a meaningful Name (e.g., SalesDataConnection), and save the updated file.
class RenameDbConnection
{
    static void Main()
    {
        // Load an existing workbook that contains a DBConnection
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all external connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Check if the connection is a DBConnection
            if (connection is DBConnection dbConnection)
            {
                // Rename the connection to a descriptive identifier
                dbConnection.Name = "SalesDataConnection";
                Console.WriteLine($"DBConnection renamed to: {dbConnection.Name}");
            }
        }

        // Save the workbook with the updated connection name
        workbook.Save("output.xlsx");
    }
}
