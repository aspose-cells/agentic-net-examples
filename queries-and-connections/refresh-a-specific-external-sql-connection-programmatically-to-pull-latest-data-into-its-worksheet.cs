using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class RefreshSpecificSqlConnection
    {
        public static void Main(string[] args)
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
            const string inputPath = "InputWithSqlConnection.xlsx";
            const string outputPath = "OutputAfterRefresh.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that contains the external SQL connection
                Workbook workbook = new Workbook(inputPath);

                // Identify the external connection to refresh (first one in the collection)
                if (workbook.DataConnections.Count == 0)
                {
                    Console.WriteLine("No external data connections found in the workbook.");
                    return;
                }

                ExternalConnection connection = workbook.DataConnections[0];

                // Ensure the connection is a DBConnection (SQL based)
                if (connection is DBConnection dbConn)
                {
                    // Force the connection to refresh when the workbook is opened
                    dbConn.RefreshOnLoad = true;
                    dbConn.RefreshInternal = 0; // immediate refresh

                    // Save the workbook with updated connection settings
                    workbook.Save(outputPath);
                    Console.WriteLine($"External SQL connection refreshed and workbook saved to '{outputPath}'.");
                }
                else
                {
                    Console.WriteLine("The specified connection is not a DBConnection (SQL).");
                }
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}