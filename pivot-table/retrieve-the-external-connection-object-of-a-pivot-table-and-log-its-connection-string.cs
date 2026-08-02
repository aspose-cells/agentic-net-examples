// Title: Aspose.Cells for .NET – Get and Log a Pivot Table’s External Connection String (C#)
// Description: Learn how to create a workbook, add a pivot table, retrieve its external data connections with PivotTable.GetSourceDataConnections(), and write the connection name and connection string to the console using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | pivot table external connection | GetSourceDataConnections | connection string | Excel automation | .NET | retrieve pivot source data | log external connection
// Common Searches: Aspose.Cells get pivot table external connection C# | How to read connection string from a pivot table using Aspose.Cells | Retrieve source data connections for a pivot table .NET | Log external connection details of a pivot table Aspose.Cells | C# example GetSourceDataConnections Aspose.Cells
// Developer Intent: Obtain the external connection object of a pivot table and output its connection string.
// Use Cases: Audit data sources by recording the connection string of each pivot table. | Debug missing or broken external connections during workbook generation. | Validate or modify connection strings before refreshing pivot data programmatically.
// AI Prompts: Generate C# code that updates the ConnectionString of a pivot table’s external connection with Aspose.Cells. | Show an example that iterates over all external connections of a pivot table and prints every property. | Explain how to add a new external data connection to a pivot table and then retrieve it using GetSourceDataConnections.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsPivotExternalConnectionDemo
{
    // Learn how to create a workbook, add a pivot table, retrieve its external data connections with PivotTable.GetSourceDataConnections(), and write the connection name and connection string to the console using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will be used as the pivot table source
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(850);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(430);

            // Add a pivot table based on the sample data
            // The source range is A1:B4, the destination cell is D1, and the pivot table name is "PivotTable1"
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // OPTIONAL: configure the pivot table (add fields) – not required for connection retrieval
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product column as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column as data field

            // Retrieve external data connections associated with the pivot table
            ExternalConnection[] connections = pivot.GetSourceDataConnections();

            // Log connection information
            if (connections != null && connections.Length > 0)
            {
                // Assuming the first connection is the relevant one
                ExternalConnection conn = connections[0];
                Console.WriteLine("External Connection Name: " + conn.Name);
                Console.WriteLine("Connection String: " + conn.ConnectionString);
            }
            else
            {
                Console.WriteLine("No external data connections found for the pivot table.");
            }

            // Save the workbook (required by lifecycle rule)
            workbook.Save("PivotTableWithExternalConnection.xlsx");
        }
    }
}
