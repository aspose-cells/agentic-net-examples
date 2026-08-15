// Title: C# – Remove a Calculated Field from an Aspose.Cells PivotTable without Affecting Other Fields
// Description: Demonstrates how to create a workbook, add sample data, build a PivotTable, insert a calculated field (e.g., DoubleSales), locate that field in the DataFields collection, remove it with PivotTable.RemoveField, and refresh the PivotTable so remaining fields stay unchanged. The workbook is then saved with the updated layout.
// Keywords: Aspose.Cells C# remove calculated field | PivotTable Delete Calculated Field .NET | Aspose.Cells RemoveField example | C# PivotTable calculated field removal | Aspose.Cells PivotTable refresh after delete
// Common Searches: how to delete a calculated field from a pivot table using Aspose.Cells for .NET | remove specific calculated data field without changing other fields Aspose.Cells | Aspose.Cells PivotTable RemoveField method usage | C# code to drop a calculated field from a PivotTable | Aspose.Cells example removing DoubleSales field
// Developer Intent: Delete a calculated field from an Aspose.Cells PivotTable while preserving all other row, column, and data fields.
// Use Cases: Temporarily add a calculated metric for analysis and then remove it before publishing the workbook. | Provide UI controls that let users toggle calculated fields on a PivotTable in real time. | Clean up intermediate calculated fields in automated report generation pipelines.
// AI Prompts: Generate C# code with Aspose.Cells to remove a calculated field named 'ProfitMargin' from a PivotTable without altering other fields. | Explain step‑by‑step how to find a calculated field in a PivotTable's DataFields collection and delete it using Aspose.Cells. | Show how to verify that a calculated field has been removed, then refresh and recalculate the PivotTable in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, build a PivotTable, insert a calculated field (e.g., DoubleSales), locate that field in the DataFields collection, remove it with PivotTable.RemoveField, and refresh the PivotTable so remaining fields stay unchanged. The workbook is then saved with the updated layout.
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
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Orange");
                sheet.Cells["B4"].PutValue(200);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Add a calculated field (e.g., double the sales)
                pivotTable.AddCalculatedField("DoubleSales", "=Sales*2", true);

                // Refresh and calculate to populate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Locate the calculated field in the DataFields collection
                PivotField calculatedField = null;
                foreach (PivotField field in pivotTable.DataFields)
                {
                    if (field.IsCalculatedField && field.Name == "DoubleSales")
                    {
                        calculatedField = field;
                        break;
                    }
                }

                // If the calculated field is found, remove it from the Data area
                if (calculatedField != null)
                {
                    // Remove by field name (alternatively, use the field index)
                    pivotTable.RemoveField(PivotFieldType.Data, calculatedField.Name);
                    // Recalculate after removal
                    pivotTable.RefreshData();
                    pivotTable.CalculateData();
                }

                // Save the workbook with the updated pivot table
                workbook.Save("PivotTable_RemoveCalculatedField.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveCalculatedFieldDemo.Run();
        }
    }
}
