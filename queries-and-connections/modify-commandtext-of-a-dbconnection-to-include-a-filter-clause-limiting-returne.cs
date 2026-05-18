using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a DBConnection
            Workbook workbook = new Workbook("input.xlsx");

            // Access the collection of external connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Find the first DBConnection in the collection
            DBConnection dbConn = null;
            foreach (ExternalConnection conn in connections)
            {
                if (conn is DBConnection dbConnection)
                {
                    dbConn = dbConnection;
                    break;
                }
            }

            if (dbConn == null)
            {
                Console.WriteLine("No DBConnection objects found in the workbook.");
                return;
            }

            // Display the original command
            Console.WriteLine("Original Command: " + dbConn.Command);

            // Define the filter clause to limit rows (example: only rows where Country = 'USA')
            string filterClause = " WHERE Country = 'USA'";

            // Append the filter clause if the command does not already contain a WHERE clause
            if (!dbConn.Command.Contains("WHERE", StringComparison.OrdinalIgnoreCase))
            {
                dbConn.Command += filterClause;
            }
            else
            {
                // If a WHERE already exists, you could modify it as needed.
                // For simplicity, we replace the existing WHERE clause with the new filter.
                int whereIndex = dbConn.Command.IndexOf("WHERE", StringComparison.OrdinalIgnoreCase);
                dbConn.Command = dbConn.Command.Substring(0, whereIndex) + filterClause;
            }

            // Display the updated command
            Console.WriteLine("Updated Command: " + dbConn.Command);

            // Save the workbook with the modified connection
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as output.xlsx");
        }
    }
}