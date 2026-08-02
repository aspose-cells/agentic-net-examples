using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load the existing workbook that contains the PivotTable
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the PivotTable is on the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one PivotTable
        if (worksheet.PivotTables.Count > 0)
        {
            // Get the first PivotTable (adjust index as needed)
            PivotTable pivot = worksheet.PivotTables[0];

            // Retrieve all external connections used by this PivotTable
            ExternalConnection[] connections = pivot.GetSourceDataConnections();

            foreach (ExternalConnection connection in connections)
            {
                // We're interested only in database connections (ODBC/OLEDB)
                if (connection is DBConnection dbConnection)
                {
                    // Example: replace the server name in the connection string
                    // Adjust the pattern according to your actual connection string format
                    string oldConnectionString = dbConnection.ConnectionString;
                    string newConnectionString = oldConnectionString.Replace("Server=OldServer;", "Server=NewServer;");

                    // Apply the updated connection string
                    dbConnection.ConnectionString = newConnectionString;

                    // Optionally, update other properties such as Command if needed
                    // dbConnection.Command = "SELECT * FROM NewTable";
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}