// Title: Aspose.Cells for .NET – Update Pivot Table External Connection String and Refresh Data
// Description: Shows how to load an Excel workbook, locate the first pivot table, replace its external data connection string (e.g., to a new Access .accdb file), enable RefreshOnLoad, refresh and recalculate the pivot, and save the updated workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | pivot table | external connection | connection string | RefreshOnLoad | update data source | Excel automation | Access database
// Common Searches: Aspose.Cells change pivot table data source C# | update external connection string for Excel pivot table | refresh pivot table after modifying connection string .NET | set RefreshOnLoad for pivot table using Aspose.Cells | retrieve and edit pivot table source connections C#
// Developer Intent: The developer needs to programmatically change a pivot table’s external connection string and refresh the pivot so it pulls data from the new source.
// Use Cases: Switch a pivot table from one Access database to another without opening Excel manually. | Batch‑process multiple workbooks to point their pivots to a refreshed data warehouse. | Configure workbooks to automatically refresh pivots on open after the connection string is updated.
// AI Prompts: Write C# code that iterates through all pivot tables in a workbook, updates each external connection string to a specified Access file, and refreshes them with Aspose.Cells. | Explain how to add an external data connection to a pivot table when none exist, then refresh the pivot using Aspose.Cells for .NET. | Provide robust error handling for missing files, absent pivot tables, and missing external connections while updating a pivot table’s connection string.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Shows how to load an Excel workbook, locate the first pivot table, replace its external data connection string (e.g., to a new Access .accdb file), enable RefreshOnLoad, refresh and recalculate the pivot, and save the updated workbook using Aspose.Cells for .NET.
    public class UpdatePivotExternalConnection
    {
        public static void Run()
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook containing the pivot table with an external data connection
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the first pivot table in the worksheet
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                PivotTable pivotTable = worksheet.PivotTables[0];

                // Obtain external data connections used by the pivot table
                ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

                if (connections.Length > 0)
                {
                    // Update the connection string to point to the new data source
                    connections[0].ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NewDataSource.accdb;Persist Security Info=False;";

                    // Ensure the connection refreshes when the workbook is opened
                    connections[0].RefreshOnLoad = true;
                }
                else
                {
                    Console.WriteLine("No external connections found for the pivot table.");
                    return;
                }

                // Refresh the pivot table to pull data from the updated connection
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UpdatePivotExternalConnection.Run();
        }
    }
}
