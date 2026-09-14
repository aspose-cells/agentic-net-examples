// Title: Remove all data fields from an Aspose.Cells PivotTable in C# to produce a row‑column summary
// AI Prompts: Write C# code using Aspose.Cells that iterates through a PivotTable's DataFields collection, removes each field, and then recalculates the pivot. | Show how to clear the data area of a PivotTable while keeping row and column fields intact with Aspose.Cells for .NET. | Provide a complete example that refreshes the pivot cache, deletes every data field, and saves the workbook as a summary‑only file.
// Common Searches: C# Aspose.Cells how to clear data fields from a pivot table | remove data area from pivot table using Aspose.Cells .NET | Aspose.Cells pivot table without values only rows and columns | example code to delete all data fields in a PivotTable with Aspose.Cells | convert Aspose.Cells pivot table to plain row column summary
// Tags: Aspose.Cells pivot table purge data fields | C# clear pivot data area Aspose.Cells | Aspose.Cells refresh pivot cache after field removal | row column summary pivot Aspose.Cells | iterate PivotTable DataFields C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a PivotTable with a row field, then loops through and removes all data fields, recalculates the pivot, and saves the result as a row‑column summary workbook.
    public class RemoveAllDataFieldsFromPivotTable
    {
        public static void Main()
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
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Amount");
                worksheet.Cells["A2"].PutValue("Fruit");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("Fruit");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("Vegetable");
                worksheet.Cells["B4"].PutValue(15);

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add fields: Category as Row, Amount as Data
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh the pivot cache and calculate data
                pivotTable.RefreshData();          // Correct API to refresh the cache
                pivotTable.CalculateData();

                // ---- Remove all data fields ----
                // Loop until no data fields remain
                while (pivotTable.DataFields.Count > 0)
                {
                    // Remove the first data field in the collection by its name
                    string fieldName = pivotTable.DataFields[0].Name;
                    pivotTable.RemoveField(PivotFieldType.Data, fieldName);
                }

                // Recalculate after removal to reflect the changes
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTable_NoDataFields.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
            }
        }
    }
}
