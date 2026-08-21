// Title: C# – Retrieve PivotTable ExternalConnection Details with Aspose.Cells for .NET
// Description: Creates a workbook, adds a PivotTable, calls GetSourceDataConnections to fetch ExternalConnection objects, and prints properties such as Name, ClassType, SourceType, Command, and ConnectionString before saving the file.
// Keywords: Aspose.Cells PivotTable external connection | C# GetSourceDataConnections | Aspose.Cells ExternalConnection example | retrieve pivot table data source details | .NET workbook pivot connection string | list pivot table external connections
// Common Searches: Aspose.Cells get external connections of a PivotTable | C# retrieve PivotTable connection string Aspose | How to list PivotTable data sources with Aspose.Cells | GetSourceDataConnections example C# | Aspose.Cells external connection properties
// Developer Intent: Extract metadata of the external data source linked to a PivotTable using Aspose.Cells.
// Use Cases: Display connection name, class type, source type, command, and connection string for auditing. | Validate that a PivotTable references the correct database before report generation. | Log external connection details when automating Excel workbook creation.
// AI Prompts: Generate C# code that iterates over all ExternalConnection objects of a PivotTable and prints each property using Aspose.Cells. | Write a method that accepts a PivotTable and returns a collection of its ExternalConnection metadata. | Show how to add an ODBC external connection to a PivotTable and then retrieve its details with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a PivotTable, calls GetSourceDataConnections to fetch ExternalConnection objects, and prints properties such as Name, ClassType, SourceType, Command, and ConnectionString before saving the file.
    public class RetrievePivotTableExternalConnection
    {
        public static void Run()
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
                if (connections.Length > 0)
                {
                    ExternalConnection conn = connections[0];
                    Console.WriteLine("Connection Name: " + conn.Name);
                    Console.WriteLine("Class Type: " + conn.ClassType);
                    Console.WriteLine("Source Type: " + conn.SourceType);
                    Console.WriteLine("Command: " + conn.Command);
                    Console.WriteLine("Connection String: " + conn.ConnectionString);
                }
                else
                {
                    Console.WriteLine("No external data connections found for the pivot table.");
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save("PivotTableWithExternalConnection.xlsx");
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
            RetrievePivotTableExternalConnection.Run();
        }
    }
}
