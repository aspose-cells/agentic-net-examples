using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class ModifyDbConnectionCommand
    {
        public static void Main()
        {
            // Load an existing workbook that already contains a DBConnection.
            // Replace "input.xlsx" with the path to your source workbook.
            Workbook workbook = new Workbook("input.xlsx");

            // Access the collection of external connections in the workbook.
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Locate the first DBConnection object (if any).
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
            }
            else
            {
                // Show the original command for reference.
                Console.WriteLine("Original Command: " + dbConn.Command);

                // Define the filter clause you want to apply.
                // Example: limit rows to those where Country = 'USA'.
                string filterClause = " WHERE Country = 'USA'";

                // If the command already contains a WHERE clause, append using AND.
                // Simple check for the word "WHERE" (case‑insensitive).
                if (dbConn.Command.IndexOf("WHERE", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Replace the first occurrence of "WHERE" with "WHERE ... AND"
                    // This is a naive approach suitable for demonstration purposes.
                    int wherePos = dbConn.Command.IndexOf("WHERE", StringComparison.OrdinalIgnoreCase);
                    string beforeWhere = dbConn.Command.Substring(0, wherePos + 5); // include "WHERE"
                    string afterWhere = dbConn.Command.Substring(wherePos + 5);
                    dbConn.Command = beforeWhere + " " + afterWhere.Trim() + " AND Country = 'USA'";
                }
                else
                {
                    // Append the filter clause directly.
                    dbConn.Command = dbConn.Command.TrimEnd() + filterClause;
                }

                // Display the updated command.
                Console.WriteLine("Updated Command: " + dbConn.Command);
            }

            // Save the modified workbook.
            // Replace "output.xlsx" with the desired output path.
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as output.xlsx");
        }
    }
}