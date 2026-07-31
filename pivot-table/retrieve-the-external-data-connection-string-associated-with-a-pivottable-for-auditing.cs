// Title: Get External Data Connection String of a PivotTable with Aspose.Cells for .NET
// Description: Creates a workbook, adds a PivotTable, uses PivotTable.GetSourceDataConnections() to fetch any linked external connections, and outputs the first connection's ConnectionString for auditing before saving the file.
// Keywords: Aspose.Cells PivotTable external connection | GetSourceDataConnections C# | retrieve pivot table connection string | audit external data source Aspose | Aspose.Cells .NET data connection string
// Common Searches: Aspose.Cells read pivot table external connection string | PivotTable.GetSourceDataConnections example C# | how to audit external data connections in Aspose.Cells | list external connections of a PivotTable using Aspose
// Developer Intent: Extract the connection string of an external data source attached to a PivotTable for verification or compliance checks.
// Use Cases: Confirm that a generated PivotTable references the correct database before distribution. | Log all external connection strings from workbooks to satisfy audit requirements. | Detect missing or incorrect external connections during automated report creation.
// AI Prompts: Generate C# code with Aspose.Cells that enumerates all external data connections of a PivotTable and prints each connection string. | Show how to safely handle a PivotTable with no external connections and log an appropriate warning. | Create a reusable method that returns the first external connection string of a given PivotTable or null if none exist.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsPivotConnectionAudit
{
    // Creates a workbook, adds a PivotTable, uses PivotTable.GetSourceDataConnections() to fetch any linked external connections, and outputs the first connection's ConnectionString for auditing before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
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

                // Add a pivot table that uses the sample data as its source
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Retrieve external data connections associated with the pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                // Audit: display the connection string of the first connection (if any)
                if (connections != null && connections.Length > 0 && connections[0] != null)
                {
                    Console.WriteLine("External Connection String: " + connections[0].ConnectionString);
                }
                else
                {
                    Console.WriteLine("No external data connections found for the pivot table.");
                }

                // Save the workbook (required by lifecycle rule)
                workbook.Save("PivotTableWithAuditedConnection.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
