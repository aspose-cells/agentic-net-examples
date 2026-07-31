// Title: How to Retrieve External Connection Details from a PivotTable using Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds sample data, builds a PivotTable, and uses PivotTable.GetSourceDataConnections() to fetch any linked ExternalConnection objects. It displays the first connection's name, class type, source type, command, connection string, and refresh‑on‑load flag, then saves the workbook.
// Keywords: Aspose.Cells PivotTable external connection | C# GetSourceDataConnections | Aspose.Cells ExternalConnection example | read pivot table data source | retrieve connection string Aspose.Cells | pivot table source data connections .NET | Aspose.Cells API external connections
// Common Searches: Aspose.Cells get external connection from PivotTable | C# retrieve PivotTable source data connections | How to read connection string of a PivotTable in Aspose.Cells | Get external connection details for a PivotTable using .NET | Aspose.Cells PivotTable GetSourceDataConnections usage
// Developer Intent: Extract the external data connection information associated with a PivotTable in an Aspose.Cells workbook.
// Use Cases: Display or log the name, type, command, and connection string of a PivotTable's external data source. | Validate whether a PivotTable relies on an external connection before publishing a report. | Programmatically modify or refresh external connections of a PivotTable for dynamic data updates.
// AI Prompts: Generate C# code that enumerates all ExternalConnection objects of a PivotTable and prints each property. | Write a method that returns true if a given PivotTable has any external connections, otherwise false. | Show how to change the ConnectionString of a PivotTable's external connection and trigger a data refresh using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds sample data, builds a PivotTable, and uses PivotTable.GetSourceDataConnections() to fetch any linked ExternalConnection objects. It displays the first connection's name, class type, source type, command, connection string, and refresh‑on‑load flag, then saves the workbook.
    public class RetrievePivotTableExternalConnection
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will be used as the pivot source
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

            // (Optional) Configure the pivot fields
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1); // Sales as data field

            // Retrieve external data connections associated with the pivot table
            ExternalConnection[] connections = pivot.GetSourceDataConnections();

            // Display connection information if any exist
            if (connections.Length > 0)
            {
                ExternalConnection conn = connections[0];
                Console.WriteLine("Connection Name: " + conn.Name);
                Console.WriteLine("Class Type: " + conn.ClassType);
                Console.WriteLine("Source Type: " + conn.SourceType);
                Console.WriteLine("Command: " + conn.Command);
                Console.WriteLine("Connection String: " + conn.ConnectionString);
                Console.WriteLine("Refresh On Load: " + conn.RefreshOnLoad);
            }
            else
            {
                Console.WriteLine("No external connections associated with the pivot table.");
            }

            // Save the workbook (demonstrates the required lifecycle rule)
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PivotTableExternalConnectionDemo.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}
