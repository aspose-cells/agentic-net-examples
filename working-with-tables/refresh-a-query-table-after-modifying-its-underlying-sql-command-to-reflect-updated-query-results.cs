// Title: How to refresh an Excel QueryTable after changing its SQL command using Aspose.Cells for .NET (C#)
// AI Prompts: Load a .xlsx workbook, assign a new SqlCommand to the first QueryTable, invoke Refresh, and save the file with Aspose.Cells in C#. | Write C# code that verifies the workbook contains QueryTables, updates the SELECT statement, uses dynamic for version‑agnostic access, and refreshes the query data. | Create a robust Aspose.Cells routine that handles missing files, modifies the query, refreshes the table, and logs success or error messages.
// Common Searches: Aspose.Cells C# change SQL statement of a QueryTable and refresh data | programmatically refresh Excel query table after updating its SELECT query using Aspose.Cells | C# example to modify and refresh a QueryTable in an existing workbook with Aspose.Cells | how to use dynamic to set SqlCommand on QueryTable in older Aspose.Cells versions
// Tags: querytable sqlcommand update Aspose.Cells | refresh querytable programmatically .xlsx | dynamic invocation Aspose.Cells querytable | handle missing querytables Aspose.Cells | save workbook after query refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsQueryTableExample
{
    // // Loads an existing workbook, replaces the SqlCommand of the first QueryTable, refreshes the table to pull updated data, and saves the workbook while handling missing files and potential version differences.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "InputWithQueryTable.xlsx";
                const string outputPath = "OutputAfterRefresh.xlsx";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the workbook that contains a query table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Get the collection of query tables on the worksheet
                QueryTableCollection queryTables = sheet.QueryTables;

                // Ensure there is at least one query table
                if (queryTables.Count > 0)
                {
                    // Retrieve the first query table
                    QueryTable queryTable = queryTables[0];

                    // New SQL command to be applied
                    string newSql = "SELECT CustomerID, ContactName FROM Customers WHERE Country = 'USA'";

                    // Use dynamic to invoke members that may not exist in older library versions
                    try
                    {
                        dynamic dynQueryTable = queryTable;
                        dynQueryTable.SqlCommand = newSql;   // Set the new SQL command
                        dynQueryTable.Refresh();            // Refresh the query table
                        Console.WriteLine("Query table refreshed successfully.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Query table operation not supported or failed: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("No query tables were found in the worksheet.");
                }

                // Save the workbook with the (potentially) refreshed query results
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
