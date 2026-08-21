// Title: Get PivotTable External Connection Type (OLE DB or ODBC) with Aspose.Cells for .NET
// Description: Loads a workbook, accesses the first worksheet and its first PivotTable, extracts external connections via GetSourceDataConnections, reads the ConnectionDataSourceType enum, and translates it to a readable OLE DB or ODBC label.
// Keywords: Aspose.Cells PivotTable external connection | ConnectionDataSourceType enum | OLE DB vs ODBC pivot source | C# get pivot table data source type | Aspose.Cells GetSourceDataConnections
// Common Searches: Aspose.Cells determine pivot table connection type | C# check if PivotTable uses OLE DB or ODBC | Get external connection source type from Excel pivot using Aspose | Read ConnectionDataSourceType of a PivotTable in .NET
// Developer Intent: Find out whether a PivotTable’s external data source is OLE DB or ODBC using Aspose.Cells.
// Use Cases: Validate that a PivotTable complies with required connection standards before processing. | Log the connection type for audit trails in automated reporting pipelines. | Branch refresh or caching logic based on OLE DB versus ODBC sources.
// AI Prompts: Generate C# code that enumerates all external connections of a PivotTable and prints their ConnectionDataSourceType using Aspose.Cells. | Create a method that receives a PivotTable object and returns "OLE DB", "ODBC", or "Other" based on its external connection type. | Add robust error handling for scenarios where a PivotTable has no external connections when retrieving the source type with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExternalConnectionDemo
{
    // Loads a workbook, accesses the first worksheet and its first PivotTable, extracts external connections via GetSourceDataConnections, reads the ConnectionDataSourceType enum, and translates it to a readable OLE DB or ODBC label.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a pivot table with an external connection.
            // Replace the path with the actual file location.
            Workbook workbook = new Workbook("input.xlsx");

            // Assume the first worksheet contains the pivot table.
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first pivot table on the worksheet.
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            PivotTable pivot = sheet.PivotTables[0];

            // Retrieve all external connections used by the pivot table.
            ExternalConnection[] connections = pivot.GetSourceDataConnections();

            if (connections.Length == 0)
            {
                Console.WriteLine("The pivot table does not use any external connections.");
                return;
            }

            // For demonstration, we handle the first connection.
            ExternalConnection conn = connections[0];

            // The SourceType property indicates whether the connection is ODBC or OLE DB.
            // Values are defined in the ConnectionDataSourceType enumeration.
            ConnectionDataSourceType sourceType = conn.SourceType;

            Console.WriteLine($"Connection Name : {conn.Name}");
            Console.WriteLine($"Class Type      : {conn.ClassType}");
            Console.WriteLine($"Source Type     : {sourceType}");

            // Determine and display a friendly description.
            string friendlyType = sourceType switch
            {
                ConnectionDataSourceType.ODBCBasedSource => "ODBC",
                ConnectionDataSourceType.OLEDBBasedSource => "OLE DB",
                _ => "Other"
            };

            Console.WriteLine($"External connection is of type: {friendlyType}");

            // Save the workbook if any modifications were made (optional).
            // workbook.Save("output.xlsx");
        }
    }
}
