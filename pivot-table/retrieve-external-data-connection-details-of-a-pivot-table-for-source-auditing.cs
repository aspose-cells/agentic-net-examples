using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace PivotTableConnectionAudit
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook if you have one, otherwise create a new one
            Workbook workbook;
            string inputPath = "SourceWorkbook.xlsx";

            if (System.IO.File.Exists(inputPath))
            {
                // Load existing workbook
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with sample data and a pivot table
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(850);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(430);

                // Add a pivot table (this will create an internal data connection)
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);
                pivot.AddFieldToArea(PivotFieldType.Data, 1);
            }

            // Iterate through all worksheets and their pivot tables
            foreach (Worksheet ws in workbook.Worksheets)
            {
                for (int i = 0; i < ws.PivotTables.Count; i++)
                {
                    PivotTable pt = ws.PivotTables[i];
                    Console.WriteLine($"Worksheet: {ws.Name}, PivotTable: {pt.Name}");

                    // Retrieve external connection data sources
                    ExternalConnection[] connections = pt.GetSourceDataConnections();

                    if (connections.Length == 0)
                    {
                        Console.WriteLine("  No external data connections associated with this pivot table.");
                        continue;
                    }

                    // Display details for each connection
                    for (int c = 0; c < connections.Length; c++)
                    {
                        ExternalConnection conn = connections[c];
                        Console.WriteLine($"  Connection #{c + 1}:");
                        Console.WriteLine($"    Name               : {conn.Name}");
                        Console.WriteLine($"    Class Type         : {conn.ClassType}");
                        Console.WriteLine($"    Source Type        : {conn.SourceType}");
                        Console.WriteLine($"    Command            : {conn.Command}");
                        Console.WriteLine($"    Connection String  : {conn.ConnectionString}");
                        Console.WriteLine($"    Description        : {conn.ConnectionDescription}");
                        Console.WriteLine($"    Refresh On Load    : {conn.RefreshOnLoad}");
                    }
                }
            }

            // Save the workbook (optional, to persist any changes)
            string outputPath = "AuditedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}