// Title: Retrieve the external ODBC connection string of a pivot table using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, locates the first pivot table, and prints its external ODBC connection string. | Show how to call PivotTable.GetSourceDataConnections in Aspose.Cells to extract the connection string of an ODBC data source. | Provide a .NET example that logs the ODBC connection string from a pivot table and saves the workbook without modifications.
// Common Searches: Aspose.Cells C# get ODBC connection string from Excel pivot table | How to read external data source of a pivot table using Aspose.Cells for .NET | C# example retrieving source data connections of a pivot table in a workbook | GetSourceDataConnections method Aspose.Cells pivot table external connection | Log external ODBC connection string of a pivot table with Aspose.Cells
// Tags: Aspose.Cells pivot GetSourceDataConnections | C# extract ODBC connection string from Excel pivot | Aspose.Cells external data connection retrieval | log pivot table source connection C# | retrieve ODBC connection string Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

// This C# sample loads an Excel workbook with Aspose.Cells, accesses the first worksheet's first pivot table, uses GetSourceDataConnections to obtain the external ODBC connection string, writes the string to the console, and saves the workbook.
class RetrievePivotODBCConnectionString
{
    static void Main()
    {
        try
        {
            const string inputFile = "PivotWithODBC.xlsx";
            const string outputFile = "PivotWithODBC_Output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputFile)}");
                return;
            }

            // Load the workbook that contains the pivot table with an external ODBC connection
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one pivot table
            if (sheet.PivotTables.Count > 0)
            {
                // Get the first pivot table
                PivotTable pivot = sheet.PivotTables[0];

                // Retrieve external data connections linked to the pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                if (connections.Length > 0)
                {
                    // Log the connection string of the first external connection (ODBC)
                    Console.WriteLine("External ODBC Connection String: " + connections[0].ConnectionString);
                }
                else
                {
                    Console.WriteLine("No external data connections found for the pivot table.");
                }
            }
            else
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
            }

            // Save the workbook (optional, as no modifications are made)
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
