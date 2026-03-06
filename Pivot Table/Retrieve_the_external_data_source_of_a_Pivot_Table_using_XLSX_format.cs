using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

class RetrievePivotExternalDataSource
{
    static void Main()
    {
        // Determine the path of the workbook relative to the executable
        string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        string workbookPath = Path.Combine(exeDir, "PivotWithExternalSource.xlsx");

        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Workbook not found at path: {workbookPath}");
            return;
        }

        // Load the workbook that contains the pivot table with an external data source
        Workbook workbook = new Workbook(workbookPath);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the worksheet contains at least one pivot table
        if (worksheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Get the first pivot table
        PivotTable pivotTable = worksheet.PivotTables[0];

        // Retrieve external data connections associated with the pivot table
        ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

        // Check if any external connections exist
        if (connections.Length == 0)
        {
            Console.WriteLine("The pivot table does not have any external data connections.");
        }
        else
        {
            // Display information about each external connection
            foreach (ExternalConnection conn in connections)
            {
                Console.WriteLine($"Connection Name : {conn.Name}");
                Console.WriteLine($"Class Type      : {conn.ClassType}");
                Console.WriteLine($"Source Type     : {conn.SourceType}");
                Console.WriteLine($"Command         : {conn.Command}");
                Console.WriteLine();
            }
        }

        // Save the workbook (optional, no modifications made)
        string outputPath = Path.Combine(exeDir, "PivotExternalDataSourceInfo.xlsx");
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to: {outputPath}");
    }
}