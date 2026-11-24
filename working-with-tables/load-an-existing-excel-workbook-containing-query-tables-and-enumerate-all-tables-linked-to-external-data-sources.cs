// Title: Enumerate query tables that have external data connections in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: aspocells c# list query tables with external data connections | how to get connection string of a query table in an existing Excel file using Aspose.Cells | iterate through worksheets and query tables to find linked data sources in .NET | retrieve external connection details from Excel query tables with Aspose.Cells API | C# code to identify query tables linked to databases in an Excel workbook
// Tags: enumerate query tables Aspose.Cells | query table external connection extraction | retrieve query table connection string .NET | worksheet query table iteration C# | reflection based access to query table properties | list linked query tables in Excel workbook

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, iterates over each worksheet, accesses the QueryTables collection, uses reflection to read each table's Connection property, and prints the worksheet name, query table name, and connection string for tables that are linked to external data sources.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook that contains query tables
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of query tables on the current worksheet
                QueryTableCollection queryTables = sheet.QueryTables;

                // Enumerate all query tables
                foreach (QueryTable qt in queryTables)
                {
                    // Retrieve the connection string via reflection (property may differ across versions)
                    string connection = qt.GetType()
                                          .GetProperty("Connection")?
                                          .GetValue(qt) as string;

                    // If a connection string exists, the table is linked to an external data source
                    if (!string.IsNullOrEmpty(connection))
                    {
                        Console.WriteLine($"Worksheet: {sheet.Name}, QueryTable: {qt.Name}, Connection: {connection}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
