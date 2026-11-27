// Title: Change a PivotTable’s external data source connection string to a new SQL Server with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells update pivot table external data source connection string | how to point Excel pivot table to a different SQL Server using Aspose.Cells | programmatically set connection string for PivotTable ExternalDataSource in .NET | use reflection to modify ExternalDataSource property of a PivotTable with Aspose.Cells | change Excel pivot table data source without opening Excel via Aspose.Cells API
// Tags: Aspose.Cells pivot table external connection | update pivot table data source .NET | set pivot table connection string programmatically | reflection modify ExternalDataSource Aspose.Cells | SQL Server connection string for Excel pivot

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, checks for a PivotTable, and uses reflection to access the internal ExternalDataSource object. It then updates the ConnectionString and Command properties to target a new SQL Server database before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure a pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Attempt to update external data source connection string via reflection
            try
            {
                PropertyInfo extProp = pivotTable.GetType().GetProperty("ExternalDataSource");
                if (extProp != null)
                {
                    object externalDataSource = extProp.GetValue(pivotTable);
                    if (externalDataSource != null)
                    {
                        // Set ConnectionString
                        PropertyInfo connProp = externalDataSource.GetType().GetProperty("ConnectionString");
                        if (connProp != null && connProp.CanWrite)
                        {
                            connProp.SetValue(externalDataSource,
                                "Provider=SQLOLEDB;Data Source=NewServer;Initial Catalog=NewDatabase;Integrated Security=SSPI;");
                        }

                        // Optional: set Command text
                        PropertyInfo cmdProp = externalDataSource.GetType().GetProperty("Command");
                        if (cmdProp != null && cmdProp.CanWrite)
                        {
                            cmdProp.SetValue(externalDataSource, "SELECT * FROM dbo.YourTable");
                        }

                        Console.WriteLine("External data source connection updated.");
                    }
                    else
                    {
                        Console.WriteLine("Pivot table does not have an external data source.");
                    }
                }
                else
                {
                    Console.WriteLine("ExternalDataSource property not available in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to modify external data source: {ex.Message}");
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
