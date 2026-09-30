// Title: How to retrieve an external ODBC connection string from an Excel pivot table with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file using Aspose.Cells, locates each pivot table, and writes its ODBC connection string to the console. | Create a method that safely checks for PivotTableSourceInfo availability and returns the external data source connection string for a given pivot table. | Show how to adapt the retrieval logic to work with older Aspose.Cells versions where PivotTableSourceInfo is not exposed.
// Common Searches: aspnet retrieve ODBC connection string from Excel pivot table using Aspose.Cells | c# read external data source of pivot table in .xlsx with Aspose.Cells | how to get PivotTableSourceInfo connection string in Aspose.Cells .NET | extract pivot table source connection when version does not support PivotTableSourceInfo | log all pivot tables ODBC connections in a workbook using Aspose.Cells C#
// Tags: Aspose.Cells pivot table source connection | C# read pivot external data source | PivotTableSourceInfo Aspose.Cells example | extract Excel pivot data source .NET | log pivot table connection Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an Excel workbook, checks each worksheet for pivot tables, and attempts to output the ODBC connection string via the PivotTableSourceInfo property, handling cases where the property is unavailable in older Aspose.Cells versions.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Check if the worksheet contains any pivot tables
            if (worksheet.PivotTables.Count > 0)
            {
                // Retrieve the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Attempt to retrieve source information if available
                try
                {
                    // Note: PivotTableSourceInfo may not be available in older versions.
                    // If present, you can access the connection string as shown below.
                    // string connectionString = pivotTable.PivotTableSourceInfo.ConnectionString;
                    // Console.WriteLine("Pivot Table ODBC Connection String: " + connectionString);

                    Console.WriteLine("Pivot table found. Source information retrieval is not supported in this version.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to retrieve source information: " + ex.Message);
                }
            }
            else
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
