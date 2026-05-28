using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class PivotTableConnectionStringAudit
    {
        // Entry point required for console application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will serve as the pivot source
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(1500);

            // Add a pivot table based on the sample data range
            int pivotIndex = sheet.PivotTables.Add("A1:B3", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Retrieve external data connections linked to the pivot table
            ExternalConnection[] connections = pivot.GetSourceDataConnections();

            // Audit: display the connection string if a connection exists
            if (connections != null && connections.Length > 0)
            {
                // Typically a pivot table has a single source connection
                ExternalConnection conn = connections[0];
                Console.WriteLine("Pivot Table Connection String: " + conn.ConnectionString);
            }
            else
            {
                Console.WriteLine("No external data connection associated with the pivot table.");
            }

            // Save the workbook (optional, for verification)
            string outputPath = "PivotTableAudit.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }
        }
    }
}