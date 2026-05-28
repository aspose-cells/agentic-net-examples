using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class RetrievePivotTableOdbcConnectionString
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (this data will be the source for the pivot table)
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1200);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(850);

            // Add a pivot table based on the sample data
            int pivotIndex = worksheet.PivotTables.Add("A1:B3", "D1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Retrieve external data connections associated with the pivot table
            ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

            // Log the ODBC/OLE DB connection string if a connection exists
            if (connections.Length > 0)
            {
                // The ConnectionString property holds the ODBC/OLE DB connection information
                Console.WriteLine("External Connection String: " + connections[0].ConnectionString);
            }
            else
            {
                Console.WriteLine("No external data connection found for the pivot table.");
            }

            // Save the workbook (optional, demonstrates lifecycle compliance)
            string outputPath = "PivotTableWithOdbcConnection.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }
        }
    }
}