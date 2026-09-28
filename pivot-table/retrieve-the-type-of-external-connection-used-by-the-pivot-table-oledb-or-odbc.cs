// Title: How to identify whether a pivot table’s external connection is OLEDB or ODBC using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that retrieves the ExternalConnection objects of a given PivotTable and returns the connection name together with its source type (OLEDB, ODBC, or other). | Create a method that accepts a PivotTable instance, calls GetSourceDataConnections, examines the ConnectionDataSourceType enumeration, and prints the detected connection kind.
// Common Searches: aspnet how to check pivot table external connection type with Aspose.Cells | c# Aspose.Cells GetSourceDataConnections OLEDB vs ODBC | determine source data connection of a pivot table in Aspose.Cells .NET | retrieve external connection name and type from Aspose.Cells pivot table
// Tags: Aspose.Cells GetSourceDataConnections usage | pivot table external connection type detection | ConnectionDataSourceType OLEDB ODBC enumeration | C# Aspose.Cells pivot table source data | retrieve external connection details Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // The example builds a workbook, adds sample data, creates a pivot table, and then uses the PivotTable.GetSourceDataConnections method to obtain any associated ExternalConnection objects. It checks each connection's SourceType to determine if it is OLEDB, ODBC, or another type, outputs the connection name and detected kind, and finally saves the workbook.
    public class RetrievePivotExternalConnectionType
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i - 1);
                sheet.Cells[$"B{i}"].PutValue($"Item_{i - 1}");
            }

            // NOTE: Adding a DBConnection programmatically is not supported in the
            // current Aspose.Cells version used for this example, so we proceed
            // without explicitly creating one. The pivot table will still be created.

            // Create a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B6", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);
            pivot.AddFieldToArea(PivotFieldType.Data, 1);

            // Retrieve external connections associated with the pivot table
            ExternalConnection[] connections = pivot.GetSourceDataConnections();

            if (connections.Length > 0)
            {
                ExternalConnection conn = connections[0];

                // Determine connection type based on SourceType
                string connectionKind = conn.SourceType switch
                {
                    ConnectionDataSourceType.OLEDBBasedSource => "OLEDB",
                    ConnectionDataSourceType.ODBCBasedSource => "ODBC",
                    _ => "Other"
                };

                Console.WriteLine($"Pivot Table External Connection Name: {conn.Name}");
                Console.WriteLine($"Connection Class Type: {conn.ClassType}");
                Console.WriteLine($"Detected Connection Type: {connectionKind}");
            }
            else
            {
                Console.WriteLine("No external connections are associated with the pivot table.");
            }

            // Save the workbook (optional)
            string outputPath = "PivotExternalConnectionTypeDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
