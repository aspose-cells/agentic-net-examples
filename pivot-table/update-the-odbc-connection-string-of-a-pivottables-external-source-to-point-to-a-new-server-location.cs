using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class UpdatePivotTableOdbcConnection
    {
        public static void Main()
        {
            // Load the workbook that contains the pivot table with an ODBC connection
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables on the current worksheet
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Get all external data connections used by the pivot table
                    ExternalConnection[] connections = pivot.GetSourceDataConnections();

                    // Process each connection
                    foreach (ExternalConnection conn in connections)
                    {
                        // Only handle DBConnection (ODBC/OLE DB) types
                        if (conn is DBConnection dbConn)
                        {
                            // Example: replace old server name with new server name in the connection string
                            // Adjust "OldServerName" and "NewServerName" to match your scenario
                            string oldServer = "OldServerName";
                            string newServer = "NewServerName";

                            if (!string.IsNullOrEmpty(dbConn.ConnectionString) && dbConn.ConnectionString.Contains(oldServer))
                            {
                                dbConn.ConnectionString = dbConn.ConnectionString.Replace(oldServer, newServer);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}