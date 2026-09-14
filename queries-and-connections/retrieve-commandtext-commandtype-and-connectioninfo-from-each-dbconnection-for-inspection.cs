// Title: Inspect CommandText, CommandType, and ConnectionInfo of DBConnection objects in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file using Aspose.Cells, iterates over workbook.DataConnections, and prints the Command, CommandType, and ConnectionInfo of each DBConnection. | Create a reusable method that returns a collection of objects containing command text, command type, and connection info for every DBConnection found in a given Aspose.Cells Workbook.
// Common Searches: how to get the OLEDB command type from a DBConnection in an Excel file using Aspose.Cells | retrieve command text of external database connections in a .xlsx with Aspose.Cells .NET | list DBConnection properties such as Command, CommandType, ConnectionInfo in Aspose.Cells workbook | Aspose.Cells enumerate external DB connections and display their connection strings | C# example to inspect DataConnections DBConnection details in an Excel workbook
// Tags: read DBConnection command text Aspose.Cells | enumerate external DB connections in Excel C# | retrieve OLEDB command type Aspose.Cells | extract connection info from DBConnection Aspose.Cells | list DataConnections properties Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The program loads an Excel workbook with Aspose.Cells, iterates through its DataConnections, identifies DBConnection objects, and outputs each connection's Command (CommandText), CommandType, and ConnectionInfo (plus the modern ConnectionString) to the console before saving the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all external connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Process only DBConnection objects
            if (connection is DBConnection dbConnection)
            {
                Console.WriteLine("DBConnection found:");
                
                // CommandText (the command string)
                Console.WriteLine($"Command (CommandText): {dbConnection.Command}");
                
                // CommandType (OLEDB command type)
                Console.WriteLine($"CommandType: {dbConnection.CommandType}");
                
                // ConnectionInfo (obsolete property, still accessible)
                Console.WriteLine($"ConnectionInfo (obsolete): {dbConnection.ConnectionInfo}");
                
                // For completeness, also display the modern ConnectionString property
                Console.WriteLine($"ConnectionString: {dbConnection.ConnectionString}");
                
                Console.WriteLine();
            }
        }

        // Save the workbook (even if unchanged) to demonstrate a complete flow
        workbook.Save("output.xlsx");
    }
}
