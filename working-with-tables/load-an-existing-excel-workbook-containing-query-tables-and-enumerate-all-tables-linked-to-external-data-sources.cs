// Title: List all query tables and their external data sources in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates through every worksheet and prints each query table's name together with its SourceData value. | Extend the example to also output the connection string (or ODBC connection) for each query table, handling cases where the property is missing. | Create a reusable method that returns a collection of objects containing worksheet name, query table name, and source data for all query tables in a given workbook.
// Common Searches: aspocells c# enumerate query tables in a workbook | how to get source data of Excel query tables using Aspose.Cells .NET | retrieve external data connections from query tables with Aspose.Cells | list query tables and their connection strings in an Excel file using Aspose.Cells | aspocells reflection access QueryTable SourceData property
// Tags: enumerate query tables Aspose.Cells | retrieve query table source data .NET | list external data connections Excel Aspose.Cells | reflection access QueryTable properties C# | process workbook query tables Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The sample loads an existing Excel workbook, walks through each worksheet, accesses its QueryTableCollection, and uses reflection to read the SourceData property of every query table. It prints the worksheet name, query table name, and source data (or a placeholder when unavailable), with robust error handling for missing files or properties.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook containing query tables
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                QueryTableCollection queryTables = sheet.QueryTables;

                // Enumerate each query table on the worksheet
                foreach (QueryTable qt in queryTables)
                {
                    try
                    {
                        // Attempt to retrieve the SourceData property via reflection
                        string sourceData = string.Empty;
                        PropertyInfo propInfo = qt.GetType().GetProperty("SourceData", BindingFlags.Public | BindingFlags.Instance);
                        if (propInfo != null)
                        {
                            object value = propInfo.GetValue(qt);
                            sourceData = value as string ?? string.Empty;
                        }

                        if (!string.IsNullOrEmpty(sourceData))
                        {
                            Console.WriteLine(
                                $"Worksheet: {sheet.Name}, QueryTable: {qt.Name}, SourceData: {sourceData}");
                        }
                        else
                        {
                            // Fallback output when SourceData is unavailable
                            Console.WriteLine(
                                $"Worksheet: {sheet.Name}, QueryTable: {qt.Name}, SourceData: <not available>");
                        }
                    }
                    catch (Exception innerEx)
                    {
                        // Handle unexpected errors for a specific query table
                        Console.WriteLine($"Error processing query table \"{qt.Name}\": {innerEx.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Handle errors related to workbook loading or overall processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
