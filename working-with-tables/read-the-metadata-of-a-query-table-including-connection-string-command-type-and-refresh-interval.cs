using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class ReadQueryTableMetadata
{
    static void Main()
    {
        // Load an existing workbook that contains query tables
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Skip worksheets without query tables
            if (sheet.QueryTables.Count == 0)
                continue;

            Console.WriteLine($"Worksheet: {sheet.Name}");

            // Process each query table in the worksheet
            for (int i = 0; i < sheet.QueryTables.Count; i++)
            {
                QueryTable queryTable = sheet.QueryTables[i];
                Console.WriteLine($"  QueryTable {i + 1} Name: {queryTable.Name}");

                // Retrieve the external connection associated with the query table
                ExternalConnection externalConnection = queryTable.ExternalConnection;

                if (externalConnection != null)
                {
                    // Connection string used to connect to the external data source
                    Console.WriteLine($"    Connection String: {externalConnection.ConnectionString}");

                    // Command type (e.g., Text, StoredProcedure, Table, etc.)
                    Console.WriteLine($"    Command Type: {externalConnection.CommandType}");

                    // Refresh interval in minutes (how often the data is refreshed automatically)
                    Console.WriteLine($"    Refresh Interval (minutes): {externalConnection.RefreshInternal}");
                }
                else
                {
                    Console.WriteLine("    No external connection associated with this query table.");
                }
            }
        }

        // Optionally save the workbook after reading (no modifications made here)
        workbook.Save("output.xlsx");
    }
}