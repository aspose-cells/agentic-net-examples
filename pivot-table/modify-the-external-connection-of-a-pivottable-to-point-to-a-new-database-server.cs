using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

class ModifyPivotTableConnection
{
    static void Main()
    {
        // Load the workbook that contains the PivotTable with an external DB connection
        Workbook workbook = new Workbook("input.xlsx");

        // Define the new connection string that points to the new database server
        // Example for OLE DB: Provider=SQLOLEDB;Data Source=NewServer;Initial Catalog=NewDatabase;Integrated Security=SSPI;
        string newConnectionString = "Provider=SQLOLEDB;Data Source=NewServer;Initial Catalog=NewDatabase;Integrated Security=SSPI;";

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all pivot tables in the worksheet
            foreach (PivotTable pivot in sheet.PivotTables)
            {
                // Get all external data connections used by the pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                foreach (ExternalConnection conn in connections)
                {
                    // We are interested only in database connections (DBConnection)
                    if (conn is DBConnection dbConn)
                    {
                        // Update the connection string to point to the new server
                        dbConn.ConnectionString = newConnectionString;

                        // Optionally, update the command if it contains server‑specific references
                        // dbConn.Command = dbConn.Command.Replace("OldServer", "NewServer");

                        // If the obsolete SeverCommand is used, update it as well
                        if (!string.IsNullOrEmpty(dbConn.SeverCommand))
                        {
                            dbConn.SeverCommand = dbConn.SeverCommand.Replace("OldServer", "NewServer");
                        }
                    }
                }

                // Refresh the pivot table to apply the new connection
                pivot.RefreshData();
                pivot.CalculateData();
            }
        }

        // Refresh all pivot tables in the workbook (optional but ensures consistency)
        workbook.Worksheets.RefreshPivotTables();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}