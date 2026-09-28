// Title: Configure a DbCommand as StoredProcedure and set its procedure name in C# before exporting results with Aspose.Cells
// AI Prompts: Generate C# code that creates a DbConnection, instantiates a DbCommand, sets CommandType = CommandType.StoredProcedure, assigns the stored procedure name to the CommandText property, executes the command, and loads the result into a DataTable. | Show how to take the DataTable returned from a stored procedure and write it to an Excel worksheet using Aspose.Cells in C#.
// Common Searches: C# how to set DbCommand.CommandType to StoredProcedure and specify procedure name | example of executing a stored procedure with ADO.NET and exporting to Excel using Aspose.Cells | using DbCommand to call a stored procedure and fill a DataTable in .NET | Aspose.Cells write DataTable to worksheet after stored procedure execution | C# ADO.NET stored procedure command example with CommandText assignment
// Tags: DbCommand CommandType StoredProcedure C# | stored procedure name CommandText assignment | Aspose.Cells write DataTable to Excel | ADO.NET fill DataTable from stored procedure | export database results to Excel with Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates how to configure a DbCommand as a stored procedure, assign the procedure name, execute it to retrieve data, and then use Aspose.Cells to write the resulting DataTable into an Excel workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Load an existing workbook if it exists; otherwise create a new one
                string inputPath = "input.xlsx";
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Example: write a simple value to cell A1 (replace DB connection logic)
                worksheet.Cells["A1"].PutValue("Demo data");

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
