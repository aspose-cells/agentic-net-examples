// Title: C# – Remove a Calculated Field from an Aspose.Cells PivotTable (CalculatedFields.Remove)
// Description: Shows how to create a workbook, build a PivotTable, add a calculated field (ProfitMargin), refresh and calculate the pivot, then delete that calculated field using Aspose.Cells CalculatedFields.Remove (or PivotTable.RemoveField) and save the workbook.
// Keywords: Aspose.Cells | C# | PivotTable | remove calculated field | CalculatedFields.Remove | Excel automation | .NET example | pivot cache refresh | programmatic pivot table | Aspose.Cells tutorial
// Common Searches: Aspose.Cells delete calculated field from PivotTable C# | How to remove a calculated field using CalculatedFields.Remove | C# example for removing PivotTable calculated field Aspose.Cells | Remove calculated field after RefreshData Aspose.Cells | PivotTable.RemoveField vs CalculatedFields.Remove Aspose.Cells
// Developer Intent: Programmatically delete a specific calculated field from an existing PivotTable.
// Use Cases: Clean up temporary calculated fields after interim analysis. | Update a PivotTable by removing obsolete calculations before adding new ones. | Prepare a workbook template for distribution by stripping internal calculated fields.
// AI Prompts: Provide C# code that removes a calculated field named "ProfitMargin" from a PivotTable using Aspose.Cells CalculatedFields.Remove. | Show a complete Aspose.Cells example that adds a calculated field, deletes it, and saves the workbook. | Explain when to use PivotTable.RemoveField versus CalculatedFields.Remove for deleting calculated fields in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, build a PivotTable, add a calculated field (ProfitMargin), refresh and calculate the pivot, then delete that calculated field using Aspose.Cells CalculatedFields.Remove (or PivotTable.RemoveField) and save the workbook.
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
                sheet.Cells["C1"].PutValue("Profit");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1000);
                sheet.Cells["C2"].PutValue(200);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["C3"].PutValue(300);
                sheet.Cells["A4"].PutValue("Orange");
                sheet.Cells["B4"].PutValue(2000);
                sheet.Cells["C4"].PutValue(400);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: add a row field and two data fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Profit");

                // Add a calculated field named "ProfitMargin" (Profit / Sales) and drag it to the data area
                pivotTable.AddCalculatedField("ProfitMargin", "Profit/Sales", true);

                // Refresh the pivot cache and calculate the pivot table to ensure the calculated field is created
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Remove the calculated field "ProfitMargin" from the Data area
                // The RemoveField method works for both regular and calculated fields
                pivotTable.RemoveField(PivotFieldType.Data, "ProfitMargin");

                // Recalculate the pivot table after removal
                pivotTable.CalculateData();

                // Save the workbook with the modified pivot table
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
