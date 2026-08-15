// Title: Aspose.Cells C# – Get and Log Pivot Table External Connection String
// Description: Shows how to create a workbook, add a pivot table, call PivotTable.GetSourceDataConnections() to obtain ExternalConnection objects, and write the first non‑empty connection string to the console (or report none). The workbook is then saved.
// Keywords: Aspose.Cells C# | PivotTable ExternalConnection | GetSourceDataConnections | connection string logging | Excel pivot external data source | Aspose.Cells .NET example | debug pivot connection | retrieve pivot source data | Aspose.Cells external connection array | log connection string
// Common Searches: Aspose.Cells get pivot table external connection string | PivotTable.GetSourceDataConnections C# example | How to read external connection of an Excel pivot table using Aspose | Log external data source of a pivot table in .NET | Aspose.Cells retrieve external connection object
// Developer Intent: Retrieve a pivot table’s external connection and display its connection string.
// Use Cases: Confirm that a pivot table is linked to the correct external data source before publishing. | Debug missing or incorrect connection strings by printing them during development. | Audit all external connections of a pivot table for compliance or migration purposes.
// AI Prompts: Write C# code with Aspose.Cells that iterates over all ExternalConnection objects from GetSourceDataConnections() and writes each connection string to a log file. | Generate an Aspose.Cells example that throws a custom exception when a pivot table’s external connection string is empty or null. | Create a script that extracts and returns a list of connection strings from every pivot table in a workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsPivotExternalConnectionDemo
{
    // Shows how to create a workbook, add a pivot table, call PivotTable.GetSourceDataConnections() to obtain ExternalConnection objects, and write the first non‑empty connection string to the console (or report none). The workbook is then saved.
    class Program
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
                sheet.Cells["B2"].PutValue(1000);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(1500);

                // Add a pivot table based on the sample data
                int pivotIndex = sheet.PivotTables.Add("A1:B3", "E3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Retrieve external connection(s) associated with the pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                // Log the connection string if a connection exists
                if (connections != null && connections.Length > 0 && !string.IsNullOrEmpty(connections[0].ConnectionString))
                {
                    Console.WriteLine("Pivot Table Connection String: " + connections[0].ConnectionString);
                }
                else
                {
                    Console.WriteLine("No external connection with a connection string found for the pivot table.");
                }

                // Save the workbook (optional, demonstrates usage of save rule)
                workbook.Save("PivotTableWithExternalConnection.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
