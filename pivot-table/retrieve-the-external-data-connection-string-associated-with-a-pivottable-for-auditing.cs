// Title: How to retrieve the external data connection name of a PivotTable in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an .xlsx file and prints the external data connection name of each PivotTable on the first worksheet. | Show how to iterate through all worksheets and all PivotTables in a workbook to list their linked external data source names using Aspose.Cells. | Provide a C# example that checks for missing PivotTables and gracefully handles PivotTables without an external data connection, using Aspose.Cells.
// Common Searches: Aspose.Cells C# get external connection name of pivot table | Read pivot table data source string from Excel file using Aspose.Cells .NET | Audit external data connections of pivot tables in a workbook with Aspose.Cells | C# code to list pivot table linked data sources in an .xlsx using Aspose.Cells | How to check if a pivot table uses an external data source with Aspose.Cells
// Tags: aspose.cells c# retrieve pivot table data source | excel pivot table external connection audit aspnet | enumerate pivot tables aspnet aspose.cells | pivot table data source extraction .xlsx c# | handle missing pivot tables aspose.cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The example loads an .xlsx workbook with Aspose.Cells, verifies that worksheets contain PivotTables, reads the first entry of each PivotTable's DataSource array (the external connection name), and writes the name to the console while handling missing files, absent PivotTables, or PivotTables without external connections.
class PivotTableConnectionAudit
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is on the first worksheet; adjust as needed
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the worksheet.");
                return;
            }

            // Get the first pivot table (or iterate as required)
            PivotTable pivot = sheet.PivotTables[0];

            // DataSource may be a string array (multiple sources) – take the first if present
            string[] dataSourceArray = pivot.DataSource;
            string dataSourceName = (dataSourceArray != null && dataSourceArray.Length > 0) ? dataSourceArray[0] : null;

            // If the pivot table is not based on an external connection, DataSource will be null or empty
            if (string.IsNullOrEmpty(dataSourceName))
            {
                Console.WriteLine("Pivot table is not linked to an external data connection.");
                return;
            }

            // Output the external data source name for auditing purposes
            Console.WriteLine($"Pivot Table External Data Source Name: {dataSourceName}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
