// Title: C# – Remove All Data Fields from an Aspose.Cells PivotTable for a Row‑Only Summary
// Description: Creates a workbook, adds sample data, builds a PivotTable with a row field (Category) and a data field (Amount), then removes every data field using PivotTable.RemoveField, refreshes the cache, recalculates, and saves the result as PivotWithoutDataFields.xlsx.
// Keywords: Aspose.Cells C# PivotTable | remove data fields PivotTable .NET | clear PivotTable values | row‑only pivot Aspose.Cells | PivotTable.RemoveField example | RefreshData CalculateData Aspose | Aspose.Cells API delete data fields | programmatic pivot layout C#
// Common Searches: Aspose.Cells remove data fields from pivot table | C# code to delete all data fields in a PivotTable | how to create a row‑only pivot with Aspose.Cells | refresh pivot after removing data fields .NET | Aspose.Cells PivotTable.RemoveField usage
// Developer Intent: Delete every data field from a PivotTable so that only row (and optional column) headings remain.
// Use Cases: Produce a category list without aggregated numbers for documentation. | Provide a pivot template that users can populate with their own data fields later. | Export the structural layout of a pivot table while omitting the calculated values.
// AI Prompts: Generate C# code using Aspose.Cells to remove all data fields from an existing PivotTable and refresh it. | Show an alternative to the while‑loop for clearing data fields in a PivotTable with Aspose.Cells. | Explain how to keep row and column fields intact while deleting data fields in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, builds a PivotTable with a row field (Category) and a data field (Amount), then removes every data field using PivotTable.RemoveField, refreshes the cache, recalculates, and saves the result as PivotWithoutDataFields.xlsx.
    class RemoveDataFieldsFromPivot
    {
        static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
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
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Fruit");
            worksheet.Cells["A3"].PutValue("Vegetable");
            worksheet.Cells["A4"].PutValue("Fruit");
            worksheet.Cells["B1"].PutValue("Amount");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(15);
            worksheet.Cells["B4"].PutValue(20);

            // Add a pivot table based on the sample data
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Add a row field and a data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh pivot cache and calculate data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Remove all data fields from the pivot table
            while (pivotTable.DataFields.Count > 0)
            {
                string dataFieldName = pivotTable.DataFields[0].Name;
                pivotTable.RemoveField(PivotFieldType.Data, dataFieldName);
            }

            // Refresh and recalculate after removal
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the modified pivot table
            string outputPath = "PivotWithoutDataFields.xlsx";
            workbook.Save(outputPath);
        }
    }
}
