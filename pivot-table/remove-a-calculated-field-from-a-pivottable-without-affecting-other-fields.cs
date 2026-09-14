// Title: How to delete a single calculated field from an Aspose.Cells PivotTable in C# without altering other data fields
// AI Prompts: Write C# code that uses Aspose.Cells to locate a calculated field named "Profit" in a PivotTable and delete it with RemoveField while preserving other data fields. | Demonstrate iterating through PivotTable.DataFields to identify the calculated field, delete it, then call RefreshData and CalculateData to update the pivot. | Create a full example that saves the workbook after the calculated field is deleted and the pivot view is refreshed.
// Common Searches: Aspose.Cells C# remove calculated field from pivot table without affecting other fields | delete specific pivot table calculated field using Aspose.Cells API | how to refresh a pivot table after removing a calculated field in C# | C# code example for removing a calculated field named Profit from Aspose.Cells pivot | Aspose.Cells remove calculated field from Data area programmatically
// Tags: delete calculated field Aspose.Cells pivot | pivot table data field deletion C# | refresh pivot after field removal Aspose.Cells | find calculated field in DataFields Aspose.Cells | Aspose.Cells calculated field removal example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, adds a pivot table with Product, Sales, and Cost fields, defines a calculated field "Profit", locates that calculated field in the DataFields collection, deletes it using RemoveField, refreshes and recalculates the pivot, and saves the result as RemovedCalculatedField.xlsx.
    class RemoveCalculatedFieldDemo
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
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Orange");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(1000);
                sheet.Cells["B3"].PutValue(2000);
                sheet.Cells["B4"].PutValue(1500);
                sheet.Cells["C1"].PutValue("Cost");
                sheet.Cells["C2"].PutValue(400);
                sheet.Cells["C3"].PutValue(800);
                sheet.Cells["C4"].PutValue(600);

                // Add a pivot table covering the data range
                int ptIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[ptIndex];

                // Add row and data fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Cost");

                // Add a calculated field named "Profit"
                pivotTable.AddCalculatedField("Profit", "=Sales-Cost", true);

                // Refresh and calculate the pivot table to populate data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Locate the calculated field in the DataFields collection
                PivotField calculatedField = null;
                foreach (PivotField df in pivotTable.DataFields)
                {
                    if (df.IsCalculatedField && df.Name == "Profit")
                    {
                        calculatedField = df;
                        break;
                    }
                }

                // If the calculated field is found, remove it from the Data area
                if (calculatedField != null)
                {
                    // Use RemoveField with the field name to delete only this calculated field
                    pivotTable.RemoveField(PivotFieldType.Data, calculatedField.Name);
                }

                // Recalculate after removal to update the pivot view
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook with the updated pivot table
                workbook.Save("RemovedCalculatedField.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            RemoveCalculatedFieldDemo.Run();
        }
    }
}
