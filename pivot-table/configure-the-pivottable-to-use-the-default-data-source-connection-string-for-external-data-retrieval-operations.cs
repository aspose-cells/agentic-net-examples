// Title: How to configure a PivotTable to use the default external data connection string with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that retrieves a PivotTable's external connections and sets its ConnectionString to an empty value using Aspose.Cells. | Show how to change a PivotTable's source data connection type to Unknown and clear the connection string in a .NET workbook with Aspose.Cells. | Provide a complete example that creates a workbook, adds a PivotTable, modifies its external connection settings, and saves the file using Aspose.Cells.
// Common Searches: Aspose.Cells C# set PivotTable external connection string to empty | How to clear default data source for a PivotTable using Aspose.Cells .NET | Change PivotTable source data connection type to Unknown in Aspose.Cells | Example of configuring PivotTable external connections in a .NET workbook
// Tags: pivot table external connection Aspose.Cells | clear external connection string Aspose.Cells C# | set pivot source connection type Unknown Aspose.Cells | save workbook after pivot table changes Aspose.Cells | retrieve pivot table source connections Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, inserts a PivotTable, retrieves its external data connections, clears the first connection's ConnectionString (making it default), optionally sets its SourceType to Unknown, and saves the workbook as ConfiguredPivotTable.xlsx.
    public class ConfigurePivotTableDefaultConnection
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
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

            // If a connection exists, configure it to use the default (empty) connection string
            if (connections != null && connections.Length > 0)
            {
                ExternalConnection conn = connections[0];
                // Setting an empty string effectively uses the default connection settings
                conn.ConnectionString = string.Empty;
                // Optionally, you can also set the source type to Unknown if needed
                conn.SourceType = ConnectionDataSourceType.Unknown;
            }

            // Save the workbook with the configured pivot table
            workbook.Save("ConfiguredPivotTable.xlsx");
            Console.WriteLine("Workbook saved as ConfiguredPivotTable.xlsx");
        }
    }
}
