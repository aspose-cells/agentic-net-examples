// Title: Refresh a named ODBC query table in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells refresh query table by name from ODBC source | programmatically update a specific query table in an Excel file using Aspose.Cells | invoke RefreshData on Aspose.Cells QueryTable with formatting preservation | refresh ODBC linked query table in .xlsx using Aspose.Cells .NET API | load workbook, locate query table, refresh data and save with Aspose.Cells
// Tags: refresh query table Aspose.Cells .NET | ODBC linked query table update Excel | query table name lookup Aspose.Cells | invoke RefreshData method via reflection | save workbook after query refresh Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace RefreshQueryTableExample
{
    // The example loads an existing Excel workbook, finds a query table named "MyQueryTable" on the first worksheet, uses reflection to call the appropriate RefreshData overload (preserving formatting when available) to pull the latest data from its ODBC connection, and then saves the refreshed workbook to a new file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains the query table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust the index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the collection of query tables on the worksheet
                QueryTableCollection queryTables = worksheet.QueryTables;

                // Name of the query table to refresh
                string queryTableName = "MyQueryTable";

                // Locate the query table by name (manual search for compatibility)
                QueryTable queryTable = null;
                foreach (QueryTable qt in queryTables)
                {
                    if (qt.Name.Equals(queryTableName, StringComparison.OrdinalIgnoreCase))
                    {
                        queryTable = qt;
                        break;
                    }
                }

                if (queryTable == null)
                {
                    Console.WriteLine($"Query table '{queryTableName}' not found.");
                    return;
                }

                // Attempt to refresh the query table using reflection (covers different API versions)
                try
                {
                    MethodInfo refreshMethod = queryTable.GetType().GetMethod("RefreshData", new[] { typeof(bool) });
                    if (refreshMethod != null)
                    {
                        // Prefer overload with preserveFormatting argument
                        refreshMethod.Invoke(queryTable, new object[] { true });
                    }
                    else
                    {
                        // Fallback to parameterless overload if it exists
                        refreshMethod = queryTable.GetType().GetMethod("RefreshData", Type.EmptyTypes);
                        refreshMethod?.Invoke(queryTable, null);
                    }

                    Console.WriteLine($"Query table '{queryTableName}' refreshed successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to refresh query table: {ex.Message}");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with refreshed data
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
