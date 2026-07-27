using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class EnumerateExternalQueryTables
    {
        public static void Main()
        {
            // Load the existing workbook that contains query tables
            // Replace "input.xlsx" with the actual path to your workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of query tables in the current worksheet
                QueryTableCollection queryTables = sheet.QueryTables;

                // If there are no query tables, continue to the next worksheet
                if (queryTables.Count == 0)
                    continue;

                Console.WriteLine($"Worksheet: {sheet.Name}");

                // Enumerate each query table
                for (int i = 0; i < queryTables.Count; i++)
                {
                    QueryTable queryTable = queryTables[i];

                    // Retrieve the external connection associated with the query table
                    ExternalConnection externalConnection = queryTable.ExternalConnection;

                    // If the query table has an external connection, display its details
                    if (externalConnection != null)
                    {
                        Console.WriteLine($"  Query Table Name: {queryTable.Name}");
                        Console.WriteLine($"    Connection ID   : {queryTable.ConnectionId}");
                        Console.WriteLine($"    External Conn ID: {externalConnection.Id}");
                        Console.WriteLine($"    External Conn Name: {externalConnection.Name}");
                        Console.WriteLine($"    Connection String : {externalConnection.ConnectionString}");
                        Console.WriteLine($"    Refresh On Load   : {externalConnection.RefreshOnLoad}");
                    }
                }
            }

            // Optionally save the workbook (unchanged) to a new file
            workbook.Save("output.xlsx");
        }
    }
}