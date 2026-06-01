using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

class RetrievePivotExternalConnectionType
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Retrieve external connections used by the pivot table
            ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

            if (connections == null || connections.Length == 0)
            {
                Console.WriteLine("No external connections found for the pivot table.");
            }
            else
            {
                // Work with the first connection (most pivot tables have a single source connection)
                ExternalConnection conn = connections[0];

                // Display the generic class type (Database, WebQuery, etc.)
                Console.WriteLine($"Class Type: {conn.ClassType}");

                // Display the specific source type which indicates ODBC or OLE DB
                Console.WriteLine($"Source Type: {conn.SourceType}");

                // Determine whether the connection is ODBC or OLE DB
                string connectionKind = conn.SourceType switch
                {
                    ConnectionDataSourceType.ODBCBasedSource => "ODBC",
                    ConnectionDataSourceType.OLEDBBasedSource => "OLE DB",
                    _ => "Other"
                };

                Console.WriteLine($"Connection Kind: {connectionKind}");
            }

            // Save the workbook (optional, just to demonstrate lifecycle compliance)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}