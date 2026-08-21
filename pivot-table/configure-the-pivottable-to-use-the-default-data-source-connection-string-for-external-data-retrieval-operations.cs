// Title: Set a default OLE DB connection string for a PivotTable with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add sample data, insert a PivotTable, retrieve its external data connections, assign a default OLE DB connection string, set the source type to OLEDBBasedSource, and save the file using Aspose.Cells in C#.
// Keywords: Aspose.Cells PivotTable connection string | C# set OLE DB source for PivotTable | external data connection Aspose.Cells | Configure PivotTable data source .NET | ConnectionDataSourceType OLEDBBasedSource | Aspose.Cells ExternalConnection example
// Common Searches: how to change pivot table connection string Aspose.Cells | set default OLE DB source for PivotTable C# | retrieve and modify external connections of a PivotTable Aspose | Aspose.Cells OLEDBBasedSource usage | programmatically update PivotTable data source .NET
// Developer Intent: Programmatically assign or update the default OLE DB connection string of a PivotTable’s external data source and ensure the source type is correctly set.
// Use Cases: Create a new workbook with a PivotTable that reads data from an Access database via a predefined OLE DB connection. | Update an existing PivotTable to point to a different database without rebuilding the table. | Validate and correct the source type after changing the connection string to avoid runtime errors.
// AI Prompts: Generate C# code using Aspose.Cells to set a custom OLE DB connection string for a PivotTable and change its SourceType to OLEDBBasedSource. | Show how to list all external connections of a PivotTable, modify the first connection’s ConnectionString, and save the workbook. | Explain how to add a new OLE DB external connection to a PivotTable when none exist, using Aspose.Cells in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

// Demonstrates how to create a workbook, add sample data, insert a PivotTable, retrieve its external data connections, assign a default OLE DB connection string, set the source type to OLEDBBasedSource, and save the file using Aspose.Cells in C#.
class ConfigurePivotTableDefaultConnection
{
    static void Main()
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
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Retrieve external data connections associated with the pivot table
            ExternalConnection[] connections = pivot.GetSourceDataConnections();

            // If a connection exists, configure its default connection string
            if (connections != null && connections.Length > 0)
            {
                // Example default connection string (adjust as needed for your environment)
                string defaultConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Data\\SalesData.accdb;Persist Security Info=False;";

                // Set the connection string on the first (and typically only) connection
                connections[0].ConnectionString = defaultConnectionString;

                // Ensure the source type matches the connection string provider
                connections[0].SourceType = ConnectionDataSourceType.OLEDBBasedSource;
            }

            // Save the workbook with the configured pivot table
            workbook.Save("ConfiguredPivotTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
