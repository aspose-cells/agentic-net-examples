// Title: Change the ConnectionDescription of a DBConnection in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, finds every DBConnection, sets a new ConnectionDescription, and saves the workbook. | Write a method that opens a workbook, iterates its DataConnections collection, updates the ConnectionDescription of each DBConnection, and includes error handling for missing files. | Show a complete example that demonstrates how to modify external database connection metadata (ConnectionDescription) in an Excel file and persist the changes with Aspose.Cells.
// Common Searches: aspnet aspose.cells update DBConnection ConnectionDescription in existing workbook | c# programmatically change description of Excel data connection after migration | how to edit external database connection metadata in an .xlsx using Aspose.Cells | set ConnectionDescription property for DBConnection objects in Aspose.Cells .NET example | modify Excel workbook external connections description with Aspose.Cells
// Tags: Aspose.Cells update DBConnection ConnectionDescription | C# modify Excel external data connection description | Aspose.Cells edit workbook DataConnections | change DBConnection metadata in .xlsx | save workbook after external connection update

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // The example loads 'input.xlsx' with Aspose.Cells, iterates the workbook's DataConnections, replaces the ConnectionDescription of any DBConnection with a custom string, and saves the result as 'output.xlsx', including checks for missing files and exception handling.
    public class UpdateDbConnectionDescription
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook containing external connections
                Workbook workbook = new Workbook(inputPath);

                // Access the collection of external connections
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Update the description of any DBConnection found
                foreach (ExternalConnection connection in connections)
                {
                    if (connection is DBConnection dbConn)
                    {
                        dbConn.ConnectionDescription = "Migrated connection for the new data source";
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
