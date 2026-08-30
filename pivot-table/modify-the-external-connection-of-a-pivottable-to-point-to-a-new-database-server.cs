// Title: Change the external database server of an Excel PivotTable with Aspose.Cells for .NET (C#)
// AI Prompts: Replace the old server name with a new one in every DBConnection used by PivotTables in an Excel workbook using Aspose.Cells C#. | Loop through all worksheets, locate each PivotTable, update ODBC/OLEDB connection properties (ConnectionString, SourceFile, Command, SecondCommand) and then refresh the PivotTable. | Save the modified workbook after adjusting external connections and recalculate PivotTable data programmatically in a .NET application.
// Common Searches: asp.net change pivot table external connection server name programmatically | c# aspose.cells update odbc connection string for pivot tables | how to refresh pivot tables after modifying dbconnection in excel using aspose | replace old sql server with new one in excel pivot table source using aspose.cells | iterate worksheets to modify external connections of pivot tables in c#
// Tags: Aspose.Cells modify PivotTable external DB connection | C# update Excel PivotTable ODBC connection string | refresh PivotTable after connection change Aspose | iterate worksheets PivotTables Aspose.Cells | replace server name in DBConnection properties C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

// The example loads an Excel workbook, iterates through each worksheet and its PivotTables, retrieves any DBConnection objects, replaces occurrences of a specified old server name with a new one in the connection string, source file, command, and second command, refreshes and recalculates each PivotTable, and finally saves the updated workbook.
class ModifyPivotTableConnection
{
    static void Main()
    {
        // Paths to the source and destination workbooks
        string inputFile = "input.xlsx";
        string outputFile = "output.xlsx";

        // Server names to replace in the connection string
        string oldServer = "OldServerName";
        string newServer = "NewServerName";

        // Load the workbook (create rule)
        Workbook workbook = new Workbook(inputFile);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all pivot tables in the worksheet
            foreach (PivotTable pivot in sheet.PivotTables)
            {
                // Get external connections used by the pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                foreach (ExternalConnection conn in connections)
                {
                    // Only process database connections (ODBC/OLEDB)
                    if (conn is DBConnection dbConn)
                    {
                        // Update the connection string to point to the new server
                        if (!string.IsNullOrEmpty(dbConn.ConnectionString))
                        {
                            dbConn.ConnectionString = dbConn.ConnectionString.Replace(oldServer, newServer);
                        }

                        // Optionally update other properties that may contain the server path
                        if (!string.IsNullOrEmpty(dbConn.SourceFile))
                        {
                            dbConn.SourceFile = dbConn.SourceFile.Replace(oldServer, newServer);
                        }

                        if (!string.IsNullOrEmpty(dbConn.Command))
                        {
                            dbConn.Command = dbConn.Command.Replace(oldServer, newServer);
                        }

                        if (!string.IsNullOrEmpty(dbConn.SecondCommand))
                        {
                            dbConn.SecondCommand = dbConn.SecondCommand.Replace(oldServer, newServer);
                        }
                    }
                }

                // Refresh the pivot table to apply the new connection
                pivot.RefreshData();
                pivot.CalculateData();
            }
        }

        // Save the modified workbook (save rule)
        workbook.Save(outputFile);
    }
}
