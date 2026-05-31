using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Load an existing workbook that may contain DB connections.
        // Replace "input.xlsx" with the path to your workbook.
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all external data connections in the workbook.
        foreach (ExternalConnection conn in workbook.DataConnections)
        {
            // Process only DBConnection instances.
            if (conn is DBConnection dbConn)
            {
                Console.WriteLine("DBConnection found:");
                
                // Retrieve the command text (SQL, table name, etc.).
                Console.WriteLine($"Command (CommandText): {dbConn.Command}");
                
                // Retrieve the command type (e.g., SqlStatement, TableName).
                Console.WriteLine($"CommandType: {dbConn.CommandType}");
                
                // Retrieve the obsolete ConnectionInfo property for inspection.
                Console.WriteLine($"ConnectionInfo (obsolete): {dbConn.ConnectionInfo}");
                
                // Preferred property: the full connection string.
                Console.WriteLine($"ConnectionString: {dbConn.ConnectionString}");
                
                Console.WriteLine();
            }
        }

        // Save the workbook if any modifications were made.
        // Replace "output.xlsx" with the desired output path.
        workbook.Save("output.xlsx");
    }
}