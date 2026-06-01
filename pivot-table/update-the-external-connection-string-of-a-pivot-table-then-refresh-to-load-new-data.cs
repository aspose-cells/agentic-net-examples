using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class UpdatePivotExternalConnection
    {
        // Entry point required for console application
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
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook that contains the pivot table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (assumed to hold the pivot table)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the first pivot table in the worksheet
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                PivotTable pivotTable = worksheet.PivotTables[0];

                // Obtain external data connections associated with the pivot table
                ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

                if (connections.Length > 0)
                {
                    // Update the connection string to point to the new data source
                    connections[0].ConnectionString =
                        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NewDataSource.accdb;Persist Security Info=False;";
                }
                else
                {
                    Console.WriteLine("No external connections found for the pivot table.");
                }

                // Refresh the pivot table to load data from the updated connection
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}