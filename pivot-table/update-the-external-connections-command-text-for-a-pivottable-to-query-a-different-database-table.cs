// Title: Update a PivotTable external DB connection command to query a different table with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that changes the Command and SecondCommand of a DBConnection used by a PivotTable via Aspose.Cells. | Show how to replace the source SELECT statement of a PivotTable’s external connection in a .NET workbook using Aspose.Cells. | Write a script that iterates over a PivotTable’s external connections and sets a new query string for each DBConnection.
// Common Searches: aspnet change pivot table external connection query Aspose.Cells | C# update DBConnection command text for Excel pivot table using Aspose.Cells | how to set new SELECT statement for PivotTable source data in Aspose.Cells .NET | modify pivot table external data source command property in C# workbook
// Tags: pivot table external db connection command Aspose.Cells | update pivot table source query C# | modify DBConnection Command property .NET | Aspose.Cells change external connection query | set new SELECT for Excel pivot via Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.Pivot;

// The example loads a workbook, accesses the first PivotTable, iterates its external connections, and updates the Command and SecondCommand of any DBConnection (or the generic Command of other connections) to "SELECT * FROM NewTable", then saves the workbook as output.xlsx.
class UpdatePivotConnection
{
    static void Main()
    {
        // Load the workbook that contains the PivotTable
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one PivotTable
        if (worksheet.PivotTables.Count > 0)
        {
            // Get the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Retrieve the external connections used by the PivotTable
            ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

            foreach (ExternalConnection connection in connections)
            {
                // If the connection is a DBConnection, update its Command property
                if (connection is DBConnection dbConn)
                {
                    // Set the new command text to query a different database table
                    dbConn.Command = "SELECT * FROM NewTable";

                    // Also update the SecondCommand (used for server‑based page fields)
                    dbConn.SecondCommand = "SELECT * FROM NewTable";
                }
                else
                {
                    // For other types of external connections, update the generic Command property
                    connection.Command = "SELECT * FROM NewTable";
                }
            }
        }

        // Save the workbook with the modified connection
        workbook.Save("output.xlsx");
    }
}
