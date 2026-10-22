// Title: How to list all query tables and their external data sources in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that iterates through every worksheet and prints each query table's name. | Extend the enumeration to also output the connection string or data source details for each query table. | Add logic to gracefully skip worksheets that contain no query tables while continuing the overall scan.
// Common Searches: Aspose.Cells C# enumerate query tables in each worksheet | retrieve external data connections from Excel workbook using Aspose.Cells | list query table names and connection strings with Aspose.Cells .NET | how to loop through query tables in an Excel file using Aspose.Cells
// Tags: Aspose.Cells query table enumeration | C# iterate worksheet query tables | retrieve external data source from Excel query table | Aspose.Cells list query tables per sheet | handle missing query tables Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, checks that the file exists, then walks through each worksheet, accesses its QueryTables collection, and writes the worksheet and query table names to the console. It includes per‑worksheet error handling to ensure that worksheets without query tables do not interrupt the process.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook that contains query tables
            Workbook workbook = new Workbook(inputPath);

            // Loop through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Access the collection of query tables on the current worksheet
                    QueryTableCollection queryTables = sheet.QueryTables;

                    // Enumerate each query table
                    foreach (QueryTable qt in queryTables)
                    {
                        // Output worksheet name and query table name
                        Console.WriteLine($"Worksheet: {sheet.Name}, QueryTable: {qt.Name}");
                    }
                }
                catch (Exception exSheet)
                {
                    // Handle errors related to a specific worksheet
                    Console.WriteLine($"Error processing worksheet '{sheet.Name}': {exSheet.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
