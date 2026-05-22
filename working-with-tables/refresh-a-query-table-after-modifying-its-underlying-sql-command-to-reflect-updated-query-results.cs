using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsQueryTableRefreshDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook that contains a query table
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one query table
                if (worksheet.QueryTables.Count == 0)
                {
                    Console.WriteLine("No query tables found in the worksheet.");
                    return;
                }

                // Get the first query table
                QueryTable queryTable = worksheet.QueryTables[0];

                // Access the external connection linked to the query table
                ExternalConnection extConn = queryTable.ExternalConnection;
                if (extConn is DBConnection dbConn)
                {
                    // Modify the underlying SQL command
                    dbConn.Command = "SELECT Id, Name, Price FROM Products WHERE Price > 100";

                    // Refresh the query table – Aspose.Cells versions prior to 23.x do not expose a Refresh method.
                    // If the Refresh method is available in your version, uncomment the line below:
                    // queryTable.Refresh();

                    Console.WriteLine("Query table command updated.");
                }
                else
                {
                    Console.WriteLine("The query table does not use a DBConnection that can be modified.");
                }

                // Save the workbook with the (potentially) updated data
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}