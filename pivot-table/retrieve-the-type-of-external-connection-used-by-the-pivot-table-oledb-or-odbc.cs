// Title: Aspose.Cells for .NET – Retrieve PivotTable External Connection Type (ODBC or OLEDB) in C#
// Description: A concise C# example that loads an Excel workbook, accesses the first worksheet, extracts the first PivotTable, calls GetSourceDataConnections, and uses the ConnectionDataSourceType enum to identify whether the PivotTable’s external data source is ODBC‑based, OLEDB‑based, or another type. Includes robust file‑existence checks and error handling.
// Keywords: Aspose.Cells PivotTable external connection | C# GetSourceDataConnections | ConnectionDataSourceType ODBC | ConnectionDataSourceType OLEDB | read pivot data source type | Aspose.Cells .NET example | Excel workbook external link detection | GitHub Aspose.Cells pivot connection sample
// Common Searches: Aspose.Cells how to find ODBC or OLEDB source of a PivotTable | C# retrieve external connection type from PivotTable | GetSourceDataConnections Aspose.Cells example | determine pivot table data source type .NET | Aspose.Cells pivot external connection enum
// Developer Intent: Determine programmatically whether a PivotTable in an Excel file uses an ODBC‑based or OLEDB‑based external data source.
// Use Cases: Validate the connection type before refreshing a PivotTable to avoid runtime errors. | Log the external source (ODBC/OLEDB) for compliance and audit reporting. | Apply connection‑specific configuration (e.g., command timeout) based on the detected source type.
// AI Prompts: Write a C# function using Aspose.Cells that returns "ODBC", "OLEDB", or "Other" for any given PivotTable. | Generate error‑handling code for scenarios where a PivotTable has no external connections or an unsupported source type. | Create a PowerShell script that calls the above C# method and outputs the connection type for all PivotTables in a folder of workbooks.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // A concise C# example that loads an Excel workbook, accesses the first worksheet, extracts the first PivotTable, calls GetSourceDataConnections, and uses the ConnectionDataSourceType enum to identify whether the PivotTable’s external data source is ODBC‑based, OLEDB‑based, or another type. Includes robust file‑existence checks and error handling.
    public class RetrievePivotExternalConnectionType
    {
        public static void Run()
        {
            const string filePath = "PivotWithExternalConnection.xlsx";

            // Verify that the workbook file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook containing the pivot table with an external connection
                Workbook workbook = new Workbook(filePath);

                // Assume the pivot table is on the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one pivot table
                if (sheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivot = sheet.PivotTables[0];

                // Retrieve external connections used by the pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                if (connections.Length == 0)
                {
                    Console.WriteLine("The pivot table does not use any external connections.");
                    return;
                }

                // Examine the first connection
                ExternalConnection conn = connections[0];

                // Determine the source type (ODBC, OLEDB, etc.)
                ConnectionDataSourceType sourceType = conn.SourceType;

                string connectionKind = sourceType switch
                {
                    ConnectionDataSourceType.ODBCBasedSource => "ODBC",
                    ConnectionDataSourceType.OLEDBBasedSource => "OLEDB",
                    _ => $"Other ({sourceType})"
                };

                Console.WriteLine($"External connection type used by the pivot table: {connectionKind}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                RetrievePivotExternalConnectionType.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
