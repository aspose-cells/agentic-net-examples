using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class SetQueryTableBackgroundRefresh
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one query table in the worksheet
                if (worksheet.QueryTables.Count > 0)
                {
                    // Get the first query table
                    QueryTable queryTable = worksheet.QueryTables[0];

                    // Retrieve the associated external connection (read‑only property)
                    ExternalConnection externalConnection = queryTable.ExternalConnection;

                    if (externalConnection != null)
                    {
                        // Set BackgroundRefresh to false for synchronous data retrieval
                        externalConnection.BackgroundRefresh = false;
                        Console.WriteLine("BackgroundRefresh set to: " + externalConnection.BackgroundRefresh);
                    }
                    else
                    {
                        Console.WriteLine("The query table does not have an associated external connection.");
                    }
                }
                else
                {
                    Console.WriteLine("No query tables found in the worksheet.");
                }

                // Save the workbook with the modified settings
                string outputPath = "QueryTableBackgroundRefreshDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to " + Path.GetFullPath(outputPath));
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetQueryTableBackgroundRefresh.Run();
        }
    }
}