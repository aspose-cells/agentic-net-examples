// Title: Refresh a Specific External SQL Connection in an Excel Workbook with Aspose.Cells for .NET (C#)
// Description: Loads an existing Excel file, verifies external data connections, casts the target connection to a DBConnection, enables RefreshOnLoad, triggers a manual RefreshAll to pull the latest SQL data, and saves the refreshed workbook to a new file.
// Keywords: Aspose.Cells | C# | .NET | external SQL connection | DBConnection | RefreshOnLoad | RefreshAll | Excel data refresh | programmatic workbook update | Excel automation | data connections API
// Common Searches: Aspose.Cells refresh external SQL connection C# | how to programmatically update Excel data from SQL using Aspose | set RefreshOnLoad for DBConnection Aspose.Cells | refresh specific external connection before saving workbook | C# code to refresh all data connections in an Excel file
// Developer Intent: Programmatically refresh a workbook’s external SQL connection to retrieve the most recent data and save the updated file.
// Use Cases: Automate daily sales report refresh from a SQL Server database before distribution. | Include Excel data refresh in a nightly batch job for financial modeling. | Create a scheduled service that opens an Excel workbook, updates its SQL connection, and writes the refreshed version for downstream processing.
// AI Prompts: Generate C# code with Aspose.Cells that refreshes only the first DBConnection in a workbook and saves the result. | Show how to refresh a named external SQL connection while leaving other connections untouched using Aspose.Cells. | Explain how to enable RefreshOnLoad for a DBConnection and trigger a manual refresh in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Loads an existing Excel file, verifies external data connections, casts the target connection to a DBConnection, enables RefreshOnLoad, triggers a manual RefreshAll to pull the latest SQL data, and saves the refreshed workbook to a new file.
class RefreshExternalSqlConnection
{
    static void Main()
    {
        const string inputPath = "InputWithSqlConnection.xlsx";
        const string outputPath = "OutputRefreshed.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the external SQL connection.
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one external connection.
            if (workbook.DataConnections.Count > 0)
            {
                // Get the first external connection (assumed to be the target SQL connection).
                ExternalConnection externalConnection = workbook.DataConnections[0];

                // If the connection is a DBConnection, configure its properties.
                if (externalConnection is DBConnection dbConnection)
                {
                    // Set the connection to refresh when the workbook is opened.
                    dbConnection.RefreshOnLoad = true;
                }

                // Refresh all external connections (including the SQL connection).
                workbook.RefreshAll();
            }
            else
            {
                Console.WriteLine("Warning: No external connections found in the workbook.");
            }

            // Save the workbook with the refreshed data.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
