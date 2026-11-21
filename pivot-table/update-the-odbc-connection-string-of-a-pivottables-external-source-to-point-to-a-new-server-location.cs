// Title: How to programmatically update the ODBC connection string of a PivotTable’s external data source using Aspose.Cells for .NET (C#)
// Common Searches: C# code to change the ODBC connection string of an Excel PivotTable with Aspose.Cells | Update external data source server name for PivotTables using Aspose.Cells .NET | Modify PivotTable ODBC connection string via reflection in Aspose.Cells | Change PivotTable external source connection string in a workbook programmatically | Aspose.Cells example for updating PivotTable ODBC connection to a new server
// Tags: Aspose.Cells pivot ODBC connection update | C# external data source modification | reflection to edit PivotTable connection string | Excel workbook pivot server rename | save workbook after connection change

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The sample loads an existing Excel workbook, iterates through all worksheets and their PivotTables, uses reflection to access each PivotTable's ExternalDataSource and its ConnectionString, replaces the old server name with a new one, and saves the workbook with the updated ODBC connection strings.
class UpdatePivotTableOdbcConnection
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the old and new server names (adjust as needed)
            string oldServer = "OldServerName";
            string newServer = "NewServerName";

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables on the current worksheet
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    try
                    {
                        // Use reflection to access the ExternalDataSource property (avoids compile‑time dependency)
                        var externalSourceProp = pivot.GetType().GetProperty("ExternalDataSource");
                        if (externalSourceProp == null) continue; // No external source available

                        var externalSource = externalSourceProp.GetValue(pivot);
                        if (externalSource == null) continue;

                        // Use reflection to get and set the ConnectionString property
                        var connProp = externalSource.GetType().GetProperty("ConnectionString");
                        if (connProp == null) continue;

                        string currentConnection = connProp.GetValue(externalSource) as string;
                        if (string.IsNullOrEmpty(currentConnection)) continue;

                        // Replace the old server name with the new one
                        string updatedConnection = currentConnection.Replace(oldServer, newServer, StringComparison.OrdinalIgnoreCase);

                        // Assign the updated connection string back
                        connProp.SetValue(externalSource, updatedConnection);
                    }
                    catch (Exception innerEx)
                    {
                        // Log pivot‑specific errors but continue processing other pivots
                        Console.WriteLine($"Warning: Could not update pivot table '{pivot.Name}': {innerEx.Message}");
                    }
                }
            }

            // Save the workbook with the updated connection strings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
