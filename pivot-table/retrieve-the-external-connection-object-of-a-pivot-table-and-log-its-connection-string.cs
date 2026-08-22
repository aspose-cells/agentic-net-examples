// Title: Retrieve and log the external connection string of a PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a PivotTable, calls GetSourceDataConnections on the PivotTable, and prints the first connection's ConnectionString to the console. | Show how to detect and display an external data connection of an Excel PivotTable with Aspose.Cells in a .NET console application.
// Common Searches: Aspose.Cells C# get external connection string from a PivotTable | How to read source data connections of an Excel PivotTable using Aspose.Cells .NET | Example of logging PivotTable external connection in a console app with Aspose.Cells
// Tags: Aspose.Cells GetSourceDataConnections pivot table | log external connection string C# | retrieve pivot table external data connection .NET | Aspose.Cells external connection Excel workbook | C# console output PivotTable connection

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // The example creates a workbook, populates sample data, adds a PivotTable, uses GetSourceDataConnections to obtain any external connections, and writes the first connection string to the console before saving the file.
    public class PivotTableExternalConnectionDemo
    {
        // Entry point required for console application
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
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Retrieve external data connections associated with the pivot table
            ExternalConnection[] connections = pivot.GetSourceDataConnections();

            // Log the connection string of the first connection (if any)
            if (connections.Length > 0 && !string.IsNullOrEmpty(connections[0].ConnectionString))
            {
                Console.WriteLine("Pivot Table Connection String: " + connections[0].ConnectionString);
            }
            else
            {
                Console.WriteLine("No external connection string found for the pivot table.");
            }

            // Save the workbook
            string outputPath = "PivotTableExternalConnectionDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}
