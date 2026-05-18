using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    class DisableBackgroundRefreshDemo
    {
        static void Main()
        {
            // Load an existing workbook that contains a DBConnection.
            // Replace "input.xlsx" with the path to your workbook.
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all external connections in the workbook.
            foreach (ExternalConnection conn in workbook.DataConnections)
            {
                // Check if the connection is a DBConnection (SQL/ODBC/OLE DB).
                if (conn is DBConnection dbConn)
                {
                    // Disable background refresh to force synchronous (sequential) execution.
                    dbConn.BackgroundRefresh = false;

                    // Optional: output the new setting for verification.
                    Console.WriteLine($"Connection \"{dbConn.Name}\" BackgroundRefresh set to {dbConn.BackgroundRefresh}");
                }
            }

            // Save the workbook with the modified connection settings.
            // Replace "output.xlsx" with the desired output file path.
            workbook.Save("output.xlsx");
        }
    }
}