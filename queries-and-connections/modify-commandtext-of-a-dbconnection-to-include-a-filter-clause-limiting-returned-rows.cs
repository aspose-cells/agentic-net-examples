using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Load an existing workbook that contains a DBConnection
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Iterate through the connections to find the first DBConnection
        foreach (ExternalConnection connection in connections)
        {
            if (connection is DBConnection dbConn)
            {
                // Read the current command text
                string originalCommand = dbConn.Command;

                // Define the filter clause you want to apply
                string filterClause = " WHERE Id > 100";

                // If the original command already contains a WHERE clause, append using AND
                // Otherwise, simply add the new WHERE clause
                if (!originalCommand.Contains("WHERE", StringComparison.OrdinalIgnoreCase))
                {
                    dbConn.Command = originalCommand + filterClause;
                }
                else
                {
                    dbConn.Command = originalCommand + " AND Id > 100";
                }

                // Output the updated command for verification
                Console.WriteLine("Updated Command: " + dbConn.Command);
            }
        }

        // Save the workbook with the modified DBConnection command
        workbook.Save("output.xlsx");
    }
}