// Title: C# – Update Pivot Table External Connection String and Refresh Data with Aspose.Cells
// Description: Load a workbook, locate the first pivot table, modify its external data connection string to a new Excel file, refresh the pivot, and save the result using Aspose.Cells for .NET.
// Keywords: Aspose.Cells pivot external connection | C# change pivot data source | refresh pivot after connection update | modify pivot table source workbook | Aspose.Cells ExternalConnection example
// Common Searches: Aspose.Cells change pivot table connection string C# | Refresh pivot table after updating external data source .NET | How to edit PivotTable external connection with Aspose | Programmatically update pivot data source in Excel using C#
// Developer Intent: Change a pivot table's external connection string and refresh it to load data from a new source.
// Use Cases: Migrate existing pivot tables to a new data file without rebuilding them. | Automate bulk updates of pivot data sources after a database or file relocation. | Integrate pivot refresh into a nightly report generation pipeline.
// AI Prompts: Write C# code with Aspose.Cells that updates a pivot table's external connection string to a different Excel workbook and then refreshes the pivot. | Explain how to retrieve a PivotTable's ExternalConnection, set a new ConnectionString, and recalculate the pivot using Aspose.Cells. | Show error‑handling patterns for scenarios where a pivot table has no external connections before attempting to modify the connection string.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

// Load a workbook, locate the first pivot table, modify its external data connection string to a new Excel file, refresh the pivot, and save the result using Aspose.Cells for .NET.
class UpdatePivotExternalConnection
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook has at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: The workbook does not contain any worksheets.");
                return;
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found in the first worksheet.");
                return;
            }

            // Retrieve the first pivot table in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Obtain external data connections associated with the pivot table
            ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

            if (connections != null && connections.Length > 0)
            {
                // Use the first connection (typically there is only one)
                ExternalConnection connection = connections[0];

                // Update the connection string to point to the new data source
                connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NewData.xlsx;Extended Properties=\"Excel 12.0 Xml;HDR=YES\";";
            }
            else
            {
                Console.WriteLine("Warning: No external data connections found for the pivot table.");
            }

            // Refresh the pivot table to pull data from the updated connection
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
