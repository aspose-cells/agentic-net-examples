// Title: Refresh a Query Table After Changing Its SQL Command with Aspose.Cells for .NET
// Description: Shows how to open an Excel workbook, find a query table, cast its ExternalConnection to a DBConnection, update the Command (or ServerCommand) SQL text, refresh the query via Workbook.RefreshAll, and save the refreshed workbook.
// Keywords: Aspose.Cells | C# | .NET | query table | refresh query table | modify SQL command | DBConnection | Workbook.RefreshAll | external connection | Excel automation
// Common Searches: Aspose.Cells change query table SQL and refresh | C# update DBConnection command text in Excel | Refresh external connections after SQL edit Aspose | How to modify and refresh a query table with Aspose.Cells | Workbook.RefreshAll after changing query SQL
// Developer Intent: Update the SQL statement of an existing query table and refresh it to reflect new data.
// Use Cases: Replace an old employee list with a filtered active‑employees query and regenerate the report. | Switch a financial dashboard to a different database view by altering the DBConnection command. | Adjust pagination or server‑side parameters in ServerCommand, refresh the table, and save the result. | Automate nightly data refreshes after dynamically building SQL based on user input.
// AI Prompts: Generate C# code that changes a query table's DBConnection Command and refreshes only that table using Aspose.Cells. | Explain best practices for handling exceptions when Workbook.RefreshAll fails after a SQL update. | Show how to verify the ExternalConnection type before casting to DBConnection and updating its Command property.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Shows how to open an Excel workbook, find a query table, cast its ExternalConnection to a DBConnection, update the Command (or ServerCommand) SQL text, refresh the query via Workbook.RefreshAll, and save the refreshed workbook.
class RefreshQueryTableDemo
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that already contains a query table.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed).
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure that a query table exists in the worksheet.
            if (worksheet.QueryTables.Count == 0)
            {
                Console.WriteLine("No query tables found in the worksheet.");
                return;
            }

            // Get the first query table.
            QueryTable queryTable = worksheet.QueryTables[0];

            // Obtain the external connection linked to the query table.
            ExternalConnection externalConn = queryTable.ExternalConnection;

            // Cast the connection to DBConnection to modify the SQL command.
            if (externalConn is DBConnection dbConn)
            {
                // Update the command text (SQL query) to reflect the new data source.
                dbConn.Command = "SELECT Id, Name FROM Employees WHERE IsActive = 1";

                // If a second command is required (e.g., for server‑based page fields), set it as well.
                // dbConn.ServerCommand = "SELECT Id, Name FROM Employees";
            }
            else
            {
                Console.WriteLine("The query table's connection is not a DBConnection.");
                return;
            }

            // Refresh the query table by refreshing all external connections in the workbook.
            try
            {
                workbook.RefreshAll();
            }
            catch (Exception refreshEx)
            {
                Console.WriteLine($"Failed to refresh query table: {refreshEx.Message}");
                return;
            }

            // Save the workbook with the refreshed data.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
