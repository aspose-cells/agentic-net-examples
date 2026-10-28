// Title: Refresh all Excel query tables after changing external connection strings and capture any refresh errors with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells refresh all query tables after updating connection string | how to capture refresh errors from Excel query tables using Aspose.Cells .NET | programmatically change external connection string of query tables in Aspose.Cells | using reflection to invoke query table Refresh method in Aspose.Cells C# | log query table refresh failures when processing a workbook with Aspose.Cells
// Tags: refresh query tables Aspose.Cells C# | set external connection string query table .NET | query table refresh error handling Aspose | iterate worksheets query tables Aspose.Cells | reflection based query table manipulation .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through each worksheet's QueryTable collection, uses reflection to update the external connection string, invokes the Refresh method, checks for any RefreshError values, aggregates errors, saves the workbook, and reports success or the collected error details.
class RefreshQueryTables
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // New connection string to be applied to all query tables (if supported)
        string newConnectionString = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Data\\NewDataSource.accdb;Persist Security Info=False;";

        // List to collect any refresh errors
        List<string> refreshErrors = new List<string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            QueryTableCollection queryTables = sheet.QueryTables;

            // Iterate through each query table
            foreach (QueryTable qt in queryTables)
            {
                try
                {
                    // Attempt to set the external connection string via reflection (if the property exists)
                    if (qt.ExternalConnection != null)
                    {
                        PropertyInfo connStrProp = qt.ExternalConnection.GetType().GetProperty("ConnectionString", BindingFlags.Public | BindingFlags.Instance);
                        if (connStrProp != null && connStrProp.CanWrite)
                        {
                            connStrProp.SetValue(qt.ExternalConnection, newConnectionString);
                        }
                    }

                    // Attempt to refresh the query table via reflection (if the method exists)
                    MethodInfo refreshMethod = qt.GetType().GetMethod("Refresh", BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
                    if (refreshMethod != null)
                    {
                        refreshMethod.Invoke(qt, null);
                    }

                    // Check for refresh errors via reflection (if the property exists)
                    PropertyInfo errorProp = qt.GetType().GetProperty("RefreshError", BindingFlags.Public | BindingFlags.Instance);
                    if (errorProp != null)
                    {
                        string error = errorProp.GetValue(qt) as string;
                        if (!string.IsNullOrEmpty(error))
                        {
                            refreshErrors.Add($"Worksheet: {sheet.Name}, QueryTable: {qt.Name}, Error: {error}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Capture unexpected exceptions during refresh
                    refreshErrors.Add($"Worksheet: {sheet.Name}, QueryTable: {qt.Name}, Exception: {ex.Message}");
                }
            }
        }

        // Output any collected errors
        if (refreshErrors.Count > 0)
        {
            Console.WriteLine("Refresh completed with errors:");
            foreach (string err in refreshErrors)
            {
                Console.WriteLine(err);
            }
        }
        else
        {
            Console.WriteLine("All query tables refreshed successfully (or no refresh functionality available).");
        }

        // Save the workbook after processing
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
