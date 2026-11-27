// Title: How to determine whether a PivotTable uses OLEDB or ODBC connection with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get pivot table external connection type OLEDB ODBC | Read ExternalConnectionInfo of a PivotTable using Aspose.Cells .NET | Determine external data source type of Excel pivot table with Aspose.Cells | Using dynamic binding to access ExternalConnectionInfo property in Aspose.Cells | Check if pivot table connection is OLEDB or ODBC in C# Aspose.Cells
// Tags: Aspose.Cells pivot external connection info | C# retrieve pivot table connection type | ExternalConnectionInfo property Aspose.Cells | OLEDB vs ODBC pivot data source Aspose.Cells | dynamic binding Aspose.Cells pivot table

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, verifies a pivot table exists, and uses dynamic binding to read the PivotTable's ExternalConnectionInfo.Type property, outputting whether the data source is OLEDB or ODBC while handling versions of Aspose.Cells that may lack the property.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Attempt to retrieve external connection information using dynamic binding.
            // This avoids compile‑time dependencies on types that may not exist in older Aspose.Cells versions.
            try
            {
                dynamic dynPivot = pivotTable;
                var externalConnection = dynPivot.ExternalConnectionInfo;
                var connectionType = externalConnection.Type;
                Console.WriteLine("External connection type: " + connectionType);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                Console.WriteLine("External connection information is not available in this Aspose.Cells version.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
