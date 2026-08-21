// Title: Hide a PivotField item in Aspose.Cells for .NET using PivotField.HideItem
// Description: Creates a workbook, adds sample data, builds a pivot table, and hides the row‑field item "Apple" by calling PivotField.HideItem("Apple", true) before refreshing and saving the file.
// Keywords: Aspose.Cells | PivotTable | HideItem | C# | .NET | hide pivot item | row field | Excel export | programmatic pivot manipulation
// Common Searches: Aspose.Cells hide pivot item C# | How to hide a row field value in a pivot table using Aspose.Cells | PivotField.HideItem example .NET | Hide specific category in Aspose.Cells pivot table | Aspose.Cells hide low volume items programmatically
// Developer Intent: Programmatically hide a specific item (e.g., "Apple") in a pivot table row field with Aspose.Cells for .NET.
// Use Cases: Exclude a particular category from a sales summary generated as an Excel pivot report. | Suppress low‑volume or confidential items before exporting a pivot table to end users. | Apply business rules to dynamically hide or show row labels in automated reporting pipelines.
// AI Prompts: Show how to hide multiple pivot items in a loop with Aspose.Cells for .NET. | Provide code to toggle the hidden state of a pivot field item based on a boolean flag. | Explain the steps to unhide a previously hidden pivot item using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHideFieldDemo
{
    // Creates a workbook, adds sample data, builds a pivot table, and hides the row‑field item "Apple" by calling PivotField.HideItem("Apple", true) before refreshing and saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["A5"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["B4"].PutValue(20);
            sheet.Cells["B5"].PutValue(5);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the "Quantity" field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Get the row field (the one we want to hide a specific item from)
            PivotField rowField = pivotTable.RowFields[0];

            // Hide the pivot item "Apple" within this field
            // Using the HideItem(string itemValue, bool isHidden) method
            rowField.HideItem("Apple", true);

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotFieldItemHiddenDemo.xlsx");
        }
    }
}
