using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class ConfigurePivotTableDefaultConnection
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(850);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(430);

                // Add a pivot table based on the sample data
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "E1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Retrieve external data connections associated with the pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                // If a connection exists, set its connection string to the default (empty) value
                if (connections.Length > 0)
                {
                    // Setting an empty string effectively uses the default connection settings
                    connections[0].ConnectionString = string.Empty;
                }

                // Save the workbook with the configured pivot table
                workbook.Save("PivotTableWithDefaultConnection.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConfigurePivotTableDefaultConnection.Run();
        }
    }
}