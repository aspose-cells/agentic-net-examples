using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
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
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "AuditedWorkbook.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook containing the pivot table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet has at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Retrieve external data connections used by the pivot table
                ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

                // Report if no external connections are present
                if (connections == null || connections.Length == 0)
                {
                    Console.WriteLine("The pivot table does not use any external data connections.");
                    return;
                }

                // Display audit information for each connection
                foreach (ExternalConnection conn in connections)
                {
                    Console.WriteLine("=== External Connection Details ===");
                    Console.WriteLine($"Name                : {conn.Name}");
                    Console.WriteLine($"Class Type          : {conn.ClassType}");
                    Console.WriteLine($"Source Type         : {conn.SourceType}");
                    Console.WriteLine($"Command             : {conn.Command}");
                    Console.WriteLine($"Connection String   : {conn.ConnectionString}");
                    Console.WriteLine($"Description         : {conn.ConnectionDescription}");
                    Console.WriteLine($"Refresh On Load     : {conn.RefreshOnLoad}");
                    Console.WriteLine($"Background Refresh  : {conn.BackgroundRefresh}");
                    Console.WriteLine();
                }

                // Save the workbook (if any modifications were made)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File error: {fnfEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}