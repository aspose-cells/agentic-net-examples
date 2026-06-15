using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsQueryTableRefreshInterval
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data (simulating a source for a query table)
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("John");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Mary");

            // ------------------------------------------------------------
            // NOTE: In a real scenario you would create an ExternalConnection
            // (e.g., WebQueryConnection) and then add a QueryTable that uses it.
            // For demonstration purposes we assume a QueryTable already exists.
            // ------------------------------------------------------------

            if (worksheet.QueryTables.Count > 0)
            {
                // Get the first query table
                QueryTable queryTable = worksheet.QueryTables[0];

                // Access the associated external connection
                ExternalConnection connection = queryTable.ExternalConnection;

                // Set the automatic refresh interval to 30 minutes
                connection.RefreshInternal = 30;

                Console.WriteLine("Refresh interval set to " + connection.RefreshInternal + " minutes.");
            }
            else
            {
                Console.WriteLine("No query tables found in the worksheet. Refresh interval not set.");
            }

            // Save the workbook with the updated settings
            workbook.Save("QueryTableRefreshIntervalDemo.xlsx");
        }
    }
}