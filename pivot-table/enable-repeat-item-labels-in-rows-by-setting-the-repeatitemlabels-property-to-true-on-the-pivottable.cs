// Title: How to Enable Repeating Row Item Labels in an Aspose.Cells PivotTable (C#)
// Description: Creates a workbook with sample sales data, adds a pivot table on A1:C5, assigns "Category" and "Product" as row fields and "Sales" as a data field, then sets IsRepeatItemLabels = true for every row field, refreshes the pivot, and saves the result as PivotTable_RepeatItemLabels.xlsx.
// Keywords: Aspose.Cells | PivotTable | IsRepeatItemLabels | repeat row labels | C# | Excel automation | row field repeat labels | Aspose.Cells API | pivot table formatting | GitHub Aspose.Cells examples
// Common Searches: Aspose.Cells repeat row labels C# | Set IsRepeatItemLabels in PivotTable Aspose | Enable repeating item labels for all rows Aspose.Cells | C# code to repeat pivot table row items | Aspose.Cells pivot table label repetition
// Developer Intent: Programmatically turn on the IsRepeatItemLabels flag for each row field so that row item labels are displayed on every row of the pivot table.
// Use Cases: Print‑ready sales reports where each product line shows its category on every row. | Hierarchical data exports that keep context visible after converting to PDF or image. | Automated Excel generation for dashboards that require non‑collapsed row labels.
// AI Prompts: Show C# code that enables repeating row item labels for all fields in an Aspose.Cells pivot table. | How do I set IsRepeatItemLabels only for the "Category" field in a pivot table using Aspose.Cells? | Explain the visual impact of repeating row labels in an Excel pivot and how to toggle it with the Aspose.Cells API.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRepeatItemLabels
{
    // Creates a workbook with sample sales data, adds a pivot table on A1:C5, assigns "Category" and "Product" as row fields and "Sales" as a data field, then sets IsRepeatItemLabels = true for every row field, refreshes the pivot, and saves the result as PivotTable_RepeatItemLabels.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue("Banana");
            sheet.Cells["C3"].PutValue(800);

            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue("Carrot");
            sheet.Cells["C4"].PutValue(600);

            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue("Broccoli");
            sheet.Cells["C5"].PutValue(700);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");    // Additional row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");     // Data field

            // Enable repeating item labels for all row fields
            foreach (PivotField rowField in pivotTable.RowFields)
            {
                rowField.IsRepeatItemLabels = true;
            }

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the modified pivot table
            workbook.Save("PivotTable_RepeatItemLabels.xlsx");
        }
    }
}
