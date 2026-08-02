// Title: C# – Limit PivotField Items in Aspose.Cells with ShowItemsCount / AutoShow
// Description: Demonstrates how to create a workbook, add a pivot table, and restrict the number of row items displayed by configuring the PivotField's AutoShow properties (IsAutoShow, IsAscendShow, AutoShowCount, AutoShowField) or ShowItemsCount. The example shows the top three categories before saving the file.
// Keywords: Aspose.Cells pivot limit items | PivotField ShowItemsCount C# | AutoShow Aspose.Cells | top N items pivot row field | Aspose.Cells .NET PivotTable example | limit pivot row items
// Common Searches: Aspose.Cells limit pivot field items C# | How to show only top 3 rows in a pivot table using Aspose.Cells | Set ShowItemsCount for PivotField Aspose.Cells | Enable AutoShow for pivot rows Aspose.Cells .NET | C# code to hide low‑volume categories in a pivot
// Developer Intent: Restrict the number of items displayed in a PivotField to a specific count.
// Use Cases: Display only the top‑selling product categories in a sales dashboard. | Hide rarely used items in a financial summary to keep the report concise. | Improve readability of large pivot tables by limiting visible rows.
// AI Prompts: Generate C# code that limits a PivotField to the top 5 items using ShowItemsCount in Aspose.Cells. | Explain the effect of IsAutoShow, IsAscendShow, AutoShowCount, and AutoShowField on pivot item visibility. | Create an example that toggles between showing the top N and bottom N items in a pivot field.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a pivot table, and restrict the number of row items displayed by configuring the PivotField's AutoShow properties (IsAutoShow, IsAscendShow, AutoShowCount, AutoShowField) or ShowItemsCount. The example shows the top three categories before saving the file.
    public class LimitPivotFieldItemsDemo
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
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(100);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(200);
            cells["A4"].PutValue("C");
            cells["B4"].PutValue(150);
            cells["A5"].PutValue("D");
            cells["B5"].PutValue(120);
            cells["A6"].PutValue("E");
            cells["B6"].PutValue(80);
            cells["A7"].PutValue("F");
            cells["B7"].PutValue(60);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B7", "D3", "DemoPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the "Amount" field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Limit the number of items displayed in the row field.
            // Use AutoShow to display only the top N items.
            PivotField rowField = pivotTable.RowFields[0];
            rowField.IsAutoShow = true;          // Enable auto‑show feature
            rowField.IsAscendShow = true;        // true = top items, false = bottom items
            rowField.AutoShowCount = 3;          // Number of items to display
            rowField.AutoShowField = -1;         // -1 means the field itself is used for ranking

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("LimitPivotFieldItemsDemo.xlsx");
        }
    }
}
