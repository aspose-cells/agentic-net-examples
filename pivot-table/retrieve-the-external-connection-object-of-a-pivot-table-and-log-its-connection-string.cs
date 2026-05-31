using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class RetrievePivotExternalConnection
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data that will be used as the pivot table source
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1000);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(1500);

                // Add a pivot table based on the sample data
                int pivotIndex = sheet.PivotTables.Add("A1:B3", "E3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Retrieve external data connections associated with the pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                // Log the connection string if a connection exists
                if (connections.Length > 0 && connections[0] != null)
                {
                    Console.WriteLine("Pivot Table External Connection String: " + connections[0].ConnectionString);
                }
                else
                {
                    Console.WriteLine("No external connection found for the pivot table.");
                }

                // Save the workbook
                workbook.Save("PivotExternalConnectionDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RetrievePivotExternalConnection.Run();
        }
    }
}