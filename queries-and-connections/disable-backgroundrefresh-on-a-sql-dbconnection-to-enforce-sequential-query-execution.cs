using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class DisableBackgroundRefreshDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWithDbConnection.xlsx";
            const string outputPath = "Output_NoBackgroundRefresh.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook containing a DB connection
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all external connections
                foreach (ExternalConnection conn in workbook.DataConnections)
                {
                    // Process only DB connections (SQL/ODBC/OLE DB)
                    if (conn is DBConnection dbConn)
                    {
                        // Disable background refresh for synchronous execution
                        dbConn.BackgroundRefresh = false;

                        // Display the updated setting for verification
                        Console.WriteLine($"Connection \"{dbConn.Name}\" BackgroundRefresh set to {dbConn.BackgroundRefresh}");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}