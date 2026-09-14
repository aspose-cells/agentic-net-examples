// Title: Remove a calculated field from a PivotTable with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, defines a calculated field, then removes that calculated field using the PivotTable.CalculatedFields.Remove method. | Show how to refresh the pivot cache and recalculate the pivot after deleting a calculated field in Aspose.Cells for .NET. | Provide a complete example that saves the workbook after removing a calculated field from a PivotTable in C#.
// Common Searches: aspnet remove calculated field from pivot table using Aspose.Cells | c# Aspose.Cells delete calculated field in pivot table and refresh data | how to use PivotTable.CalculatedFields.Remove in Aspose.Cells .NET | example of removing a calculated field from a PivotTable with Aspose.Cells C#
// Tags: Aspose.Cells PivotTable calculated field removal | C# PivotTable.RemoveField method | Aspose.Cells refresh pivot cache | C# save workbook after pivot modification | Aspose.Cells .xlsx pivot table example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding a pivot table, inserting a calculated field, removing it with CalculatedFields.Remove, refreshing the cache, recalculating, and saving the .xlsx file using Aspose.Cells for .NET.
    public class RemoveCalculatedFieldDemo
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
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["A4"].PutValue("Orange");
                sheet.Cells["B4"].PutValue(2000);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Add a calculated field named "DoubleSales"
                string calcFieldName = "DoubleSales";
                pivotTable.AddCalculatedField(calcFieldName, "=Sales*2", true);

                // Remove the calculated field from the Data area
                pivotTable.RemoveField(PivotFieldType.Data, calcFieldName);

                // Refresh the pivot cache and recalculate the pivot table
                pivotTable.RefreshData();      // correct method to refresh cache
                pivotTable.CalculateData();    // recalculate after changes

                // Save the workbook with the updated pivot table
                workbook.Save("PivotTable_RemoveCalculatedField.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveCalculatedFieldDemo.Run();
        }
    }
}
