// Title: How to change a pivot table’s external connection string and refresh its data with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that updates the ConnectionString of a pivot table’s external data source using Aspose.Cells and then refreshes the pivot cache. | Demonstrate using reflection in C# to modify PivotCacheDefinition.ConnectionString when the property is not directly exposed, followed by workbook recalculation.
// Common Searches: Aspose.Cells C# update pivot table external data source connection string | Refresh pivot table after changing connection string with Aspose.Cells | Set PivotCacheDefinition.ConnectionString via reflection Aspose.Cells .NET | Recalculate workbook formulas after pivot refresh using Aspose.Cells | Handle missing PivotCacheDefinition property in Aspose.Cells pivot tables
// Tags: Aspose.Cells pivot table external connection string | Aspose.Cells refresh pivot cache .NET | C# reflection modify PivotCacheDefinition | Aspose.Cells workbook recalculate after pivot refresh | Aspose.Cells handle missing PivotCacheDefinition property

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing workbook, verifies a pivot table exists, uses reflection to set a new external connection string on the pivot cache when the PivotCacheDefinition property is not directly accessible, refreshes and calculates the pivot data, recalculates workbook formulas, ensures the output directory exists, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
                throw new InvalidOperationException("No pivot tables were found on the first worksheet.");

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // If the pivot table uses an external data source, update its connection string.
            // Note: In some Aspose.Cells versions the PivotCacheDefinition property may not be exposed.
            // The following block safely attempts to update the connection when available.
            try
            {
                // Attempt to retrieve the cache definition via reflection if the property is not directly accessible.
                var cacheProp = typeof(PivotTable).GetProperty("PivotCacheDefinition");
                if (cacheProp != null)
                {
                    var cacheDef = cacheProp.GetValue(pivotTable);
                    if (cacheDef != null)
                    {
                        var connProp = cacheDef.GetType().GetProperty("ConnectionString");
                        if (connProp != null && connProp.CanWrite)
                        {
                            connProp.SetValue(cacheDef,
                                "Provider=SQLOLEDB;Data Source=MyServer;Initial Catalog=MyDatabase;Integrated Security=SSPI;");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log but do not stop execution if connection update fails.
                Console.WriteLine($"Warning: Unable to update connection string. {ex.Message}");
            }

            // Refresh the pivot table to load data from the (potentially) new connection
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Recalculate formulas in the workbook (if any depend on pivot data)
            workbook.CalculateFormula();

            // Ensure the output directory exists (handle case when outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (string.IsNullOrEmpty(outputDir))
                outputDir = Directory.GetCurrentDirectory();

            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the updated workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
