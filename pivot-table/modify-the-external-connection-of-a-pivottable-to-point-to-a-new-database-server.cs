// Title: Change the external SQL Server connection string of an existing PivotTable using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, finds the first PivotTable, updates its PivotCacheDefinition.ConnectionString to a new SQL Server, and saves the workbook. | Show how to use reflection in C# to set the read‑only ConnectionString property of a PivotTable's cache definition when working with Aspose.Cells. | Create a reusable C# method that takes an input file, a new connection string, and an output file, then updates all PivotTables' external data sources to the specified SQL Server using Aspose.Cells.
// Common Searches: aspnet change pivot table data source connection string programmatically | Aspose.Cells update pivot cache connection to new database server | C# reflectively set PivotTable external connection string in Excel file | how to modify pivot table SQL Server source using Aspose.Cells .NET | update multiple pivot tables connection string in one workbook Aspose.Cells
// Tags: Aspose.Cells update pivot cache connection string | C# modify pivot table external data source | reflection access PivotCacheDefinition Aspose.Cells | change Excel pivot table SQL Server source | programmatic pivot table data source update .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel workbook, locates the first PivotTable, uses reflection to access its PivotCacheDefinition, sets a new SQL Server connection string, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the existing workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet where the PivotTable resides (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the first worksheet.");
                return;
            }

            // Retrieve the first PivotTable in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Attempt to update the connection string of the PivotCacheDefinition via reflection
            // (used to avoid compile‑time dependency on PivotCacheDefinition which may be absent in some versions)
            try
            {
                var cacheDefProp = pivotTable.GetType().GetProperty("PivotCacheDefinition");
                if (cacheDefProp != null)
                {
                    var cacheDef = cacheDefProp.GetValue(pivotTable);
                    var connProp = cacheDef?.GetType().GetProperty("ConnectionString");
                    if (connProp != null && connProp.CanWrite)
                    {
                        connProp.SetValue(cacheDef,
                            "Provider=SQLOLEDB;Data Source=NewServerName;Initial Catalog=NewDatabaseName;Integrated Security=SSPI;");
                        Console.WriteLine("Pivot cache connection string updated.");
                    }
                    else
                    {
                        Console.WriteLine("ConnectionString property not found or read‑only.");
                    }
                }
                else
                {
                    Console.WriteLine("PivotCacheDefinition property not available on this PivotTable.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to update Pivot cache connection: {ex.Message}");
            }

            // Save the workbook with the (potentially) modified PivotTable connection
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
