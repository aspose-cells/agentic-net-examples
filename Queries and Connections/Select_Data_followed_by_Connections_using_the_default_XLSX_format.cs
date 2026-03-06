using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDataAndConnectionsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default format is XLSX)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (A1:C4)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(0.3);

            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["C4"].PutValue(0.8);

            // Select the data range A1:C4
            AsposeRange dataRange = sheet.Cells.CreateRange("A1:C4");

            // (Optional) Demonstrate exporting the selected range to a DataTable
            // var dt = dataRange.ExportDataTable(0, true);

            // Access the workbook's external data connections collection
            var connections = workbook.DataConnections;

            // Display the current number of connections (expected to be 0 at this point)
            Console.WriteLine("Initial DataConnections count: " + connections.Count);

            // Example: Add a simple WebQueryConnection (if needed)
            /*
            WebQueryConnection webConn = new WebQueryConnection(workbook);
            webConn.Name = "SampleWebQuery";
            webConn.ConnectionString = "http://example.com/data.html";
            connections.Add(webConn);
            Console.WriteLine("DataConnections count after adding: " + connections.Count);
            */

            // Save the workbook in the default XLSX format
            workbook.Save("DataAndConnectionsDemo.xlsx");
        }
    }
}