using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
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
                sheet.Cells["B2"].PutValue(1000);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(1500);

                // Add a pivot table based on the sample data
                int pivotIndex = sheet.PivotTables.Add("A1:B3", "E3", "PivotTable1");
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
                    Console.WriteLine("No external connections found for the pivot table.");
                }

                // Save the workbook (optional, demonstrates lifecycle compliance)
                string outputPath = "PivotTableExternalConnectionDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
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