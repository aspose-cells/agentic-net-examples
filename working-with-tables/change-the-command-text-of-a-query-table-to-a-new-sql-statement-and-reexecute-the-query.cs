using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class UpdateQueryTableCommandDemo
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains a query table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one query table
                if (worksheet.QueryTables.Count == 0)
                {
                    Console.WriteLine("No query tables found in the worksheet.");
                    return;
                }

                // Get the first query table
                QueryTable queryTable = worksheet.QueryTables[0];

                // Obtain the external connection associated with the query table
                ExternalConnection extConn = queryTable.ExternalConnection;

                // Set a new SQL command text
                string newSql = "SELECT CustomerID, CompanyName FROM Customers WHERE Country = 'USA'";
                extConn.Command = newSql;

                // Optionally, set the command type to indicate a SQL statement
                extConn.CommandType = OLEDBCommandType.SqlStatement;

                // NOTE: In some Aspose.Cells versions QueryTable.Refresh() is not available.
                // If needed, you can re‑execute the query by re‑creating the query table or
                // by using other APIs. Here we simply proceed to save the workbook.

                // Save the workbook with the updated query table
                workbook.Save(outputPath);
                Console.WriteLine($"Query table command updated and workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            UpdateQueryTableCommandDemo.Run();
        }
    }
}