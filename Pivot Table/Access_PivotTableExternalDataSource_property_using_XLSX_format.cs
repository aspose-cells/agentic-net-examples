using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class PivotTableExternalDataSourceDemo
    {
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
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(850);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(430);

            // Add a pivot table based on the sample data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve external connection data sources associated with the pivot table
            ExternalConnection[] connections = pivot.GetSourceDataConnections() ?? Array.Empty<ExternalConnection>();

            // Display information about each external connection (if any)
            if (connections.Length > 0)
            {
                foreach (ExternalConnection conn in connections)
                {
                    Console.WriteLine("Connection Name: " + conn.Name);
                    Console.WriteLine("Class Type: " + conn.ClassType);
                    Console.WriteLine("Source Type: " + conn.SourceType);
                    Console.WriteLine("Command: " + conn.Command);
                    Console.WriteLine();
                }
            }
            else
            {
                Console.WriteLine("No external connections are associated with this pivot table.");
            }

            // Save the workbook in XLSX format
            workbook.Save("PivotTableExternalDataSourceDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableExternalDataSourceDemo.Run();
        }
    }
}