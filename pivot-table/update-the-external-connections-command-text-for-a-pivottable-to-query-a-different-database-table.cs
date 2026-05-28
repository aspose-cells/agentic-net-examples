using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

class UpdatePivotTableConnection
{
    static void Main()
    {
        // Load the workbook that contains the pivot table with an external DB connection
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the pivot table is on the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Iterate through all pivot tables on the worksheet
        foreach (PivotTable pivotTable in worksheet.PivotTables)
        {
            // Retrieve the external connections used by this pivot table
            ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

            foreach (ExternalConnection connection in connections)
            {
                // Check if the connection is a DBConnection (ODBC/OLE DB)
                if (connection is DBConnection dbConnection)
                {
                    // Update the command text to query a different database table
                    dbConnection.Command = "SELECT * FROM NewTable";

                    // If a second command is required, it can be set similarly:
                    // dbConnection.SecondCommand = "SELECT * FROM NewTable";
                }
            }
        }

        // Save the workbook with the updated connection
        workbook.Save("output.xlsx");
    }
}