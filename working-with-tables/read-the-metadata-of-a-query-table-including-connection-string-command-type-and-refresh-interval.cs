using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsQueryTableMetadata
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains query tables
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Check if the worksheet has any query tables
                if (sheet.QueryTables.Count == 0)
                    continue;

                Console.WriteLine($"Worksheet: {sheet.Name}");

                // Iterate through each query table in the worksheet
                for (int i = 0; i < sheet.QueryTables.Count; i++)
                {
                    QueryTable qt = sheet.QueryTables[i];
                    Console.WriteLine($"  QueryTable Name: {qt.Name}");

                    // Access the external connection associated with the query table
                    ExternalConnection extConn = qt.ExternalConnection;
                    if (extConn == null)
                    {
                        Console.WriteLine("    No external connection associated with this query table.");
                        continue;
                    }

                    // Connection string
                    Console.WriteLine($"    Connection String: {extConn.ConnectionString}");

                    // Command type (enum value)
                    Console.WriteLine($"    Command Type: {extConn.CommandType}");

                    // Refresh interval (in minutes)
                    Console.WriteLine($"    Refresh Interval (minutes): {extConn.RefreshInternal}");
                }
            }

            // Save the workbook (optional, as we only read metadata)
            workbook.Save("output.xlsx");
        }
    }
}