// Title: How to audit external data connections of an Excel pivot table using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, finds the first pivot table, and prints each external connection's name, source type, command, and connection string. | Show how to call PivotTable.GetSourceDataConnections() to enumerate connection details and log the RefreshOnLoad flag, then save the workbook. | Create a script that checks whether a pivot table has external data connections and outputs all connection properties using Aspose.Cells.
// Common Searches: Aspose.Cells C# retrieve external connections from a pivot table | list source data connections of Excel pivot table using Aspose.Cells .NET | how to get connection string of a pivot table's external data source in C# | audit pivot table data source details with Aspose.Cells library | C# example for enumerating pivot table external connections in an Excel workbook
// Tags: Aspose.Cells pivot table source connections | C# enumerate external connections Excel | retrieve pivot table connection details .NET | audit Excel pivot data source Aspose | GetSourceDataConnections API usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // The example loads an Excel file, accesses the first worksheet, obtains the first pivot table, and uses GetSourceDataConnections() to fetch any external data connections. It then iterates through each ExternalConnection object, printing properties such as Name, ClassType, SourceType, Command, ConnectionString, and RefreshOnLoad, and finally saves the workbook.
    public class PivotTableSourceConnectionAudit
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Access first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure at least one pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Retrieve external data connections
            ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

            // Output connection details
            if (connections.Length == 0)
            {
                Console.WriteLine("The pivot table does not have any external data connections.");
            }
            else
            {
                Console.WriteLine($"Pivot Table \"{pivotTable.Name}\" has {connections.Length} external connection(s):");
                for (int i = 0; i < connections.Length; i++)
                {
                    ExternalConnection conn = connections[i];
                    Console.WriteLine($"--- Connection {i + 1} ---");
                    Console.WriteLine($"Name               : {conn.Name}");
                    Console.WriteLine($"Class Type         : {conn.ClassType}");
                    Console.WriteLine($"Source Type        : {conn.SourceType}");
                    Console.WriteLine($"Command            : {conn.Command}");
                    Console.WriteLine($"Connection String  : {conn.ConnectionString}");
                    Console.WriteLine($"Refresh On Load    : {conn.RefreshOnLoad}");
                }
            }

            // Save workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
    }
}
