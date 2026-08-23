// Title: Append a WHERE filter to a DBConnection command in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that retrieves a DBConnection from a workbook, detects whether a WHERE clause already exists, and adds a filter condition to restrict rows to a specific country. | Write a method that updates the Command property of an Aspose.Cells DBConnection, inserting a new WHERE clause or appending an AND condition when a WHERE clause is present. | Create a script that loads an .xlsx file, modifies its external database connection to filter rows by Country = 'USA', and saves the updated workbook.
// Common Searches: how to modify the command text of a DBConnection in Aspose.Cells C# | add filter clause to Excel external DB connection using Aspose.Cells | append WHERE condition to existing DBConnection command in .NET | Aspose.Cells C# update data connection command to limit rows by country | handle existing WHERE clause when adding filter to DBConnection command
// Tags: Aspose.Cells DBConnection command modification | add filter to Excel external DB connection | filter rows in external DB connection C# | update DBConnection Command property Aspose | modify external connection query Excel .xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The program loads an existing Excel workbook, locates the first DBConnection, checks for an existing WHERE clause, and appends a filter (or adds an AND condition) to limit rows to Country = 'USA'. It then updates the connection's Command property and saves the workbook as a new file.
class Program
{
    static void Main()
    {
        // Load an existing workbook that contains a DBConnection
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of external connections
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
            Console.WriteLine("No DBConnection found in the workbook.");
            return;
        }

        // Display the original command text
        string originalCommand = dbConn.Command;
        Console.WriteLine("Original Command: " + originalCommand);

        // Define the filter clause to limit rows
        const string filterClause = " WHERE Country = 'USA'";

        // Append the filter clause (or add an AND if a WHERE already exists)
        string updatedCommand = originalCommand;
        if (!originalCommand.Contains("WHERE", StringComparison.OrdinalIgnoreCase))
        {
            // Remove any trailing semicolon before appending
            updatedCommand = originalCommand.TrimEnd().TrimEnd(';') + filterClause;
        }
        else
        {
            // Existing WHERE clause – add an additional condition
            updatedCommand = originalCommand.TrimEnd().TrimEnd(';') + " AND Country = 'USA'";
        }

        // Set the modified command back to the DBConnection
        dbConn.Command = updatedCommand;
        Console.WriteLine("Updated Command: " + dbConn.Command);

        // Save the workbook with the modified connection
        workbook.Save("output.xlsx");
        Console.WriteLine("Workbook saved as output.xlsx");
    }
}
