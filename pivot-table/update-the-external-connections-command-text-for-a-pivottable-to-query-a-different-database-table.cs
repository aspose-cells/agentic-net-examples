using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.Pivot;

class UpdatePivotTableConnection
{
    static void Main()
    {
        // Load the workbook that contains the PivotTable with an external DB connection
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets (or target a specific one)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all PivotTables on the worksheet
            foreach (PivotTable pivot in sheet.PivotTables)
            {
                // Retrieve the external connections used by the PivotTable
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                // Update the command text of any DBConnection found
                foreach (ExternalConnection conn in connections)
                {
                    if (conn is DBConnection dbConn)
                    {
                        // Set a new SQL command that points to a different table
                        dbConn.Command = "SELECT * FROM NewTable";

                        // If the PivotTable uses server‑based page fields, also update SecondCommand
                        if (!string.IsNullOrEmpty(dbConn.SecondCommand))
                        {
                            dbConn.SecondCommand = "SELECT * FROM NewTable";
                        }
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}