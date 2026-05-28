using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class RetrievePivotTableExternalConnection
    {
        public static void Main()
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
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(850);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(430);

                // Add a pivot table based on the sample data
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Retrieve external data connections associated with the pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                // Display connection details if any exist
                if (connections != null && connections.Length > 0)
                {
                    ExternalConnection conn = connections[0];
                    Console.WriteLine("External Connection Details:");
                    Console.WriteLine($"Name          : {conn.Name}");
                    Console.WriteLine($"Class Type    : {conn.ClassType}");
                    Console.WriteLine($"Source Type   : {conn.SourceType}");
                    Console.WriteLine($"Command       : {conn.Command}");
                    Console.WriteLine($"ConnectionStr : {conn.ConnectionString}");
                }
                else
                {
                    Console.WriteLine("No external connections found for the pivot table.");
                }

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                workbook.Save("PivotTableExternalConnectionDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}