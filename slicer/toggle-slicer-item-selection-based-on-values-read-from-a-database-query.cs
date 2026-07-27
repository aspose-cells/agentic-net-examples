using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerToggle
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data for the pivot table
            cells["A1"].PutValue("Category");
            cells["A2"].PutValue("A");
            cells["A3"].PutValue("B");
            cells["A4"].PutValue("C");
            cells["A5"].PutValue("A");
            cells["B1"].PutValue("Amount");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);
            cells["B5"].PutValue(40);

            // Create a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Amount
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer for the "Category" field
            int slicerIdx = sheet.Slicers.Add(pivot, "F3", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // ------------------------------------------------------------
            // Simulate reading values from a database query.
            // In a real scenario, replace this with actual DB access code.
            // ------------------------------------------------------------
            List<string> valuesToSelect = GetValuesFromDatabase();

            // Toggle slicer items based on the values retrieved from the database
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                // Select the item if its value exists in the database result set
                item.Selected = valuesToSelect.Contains(item.Value);
            }

            // Refresh the slicer to apply the selection changes
            slicer.Refresh();

            // Save the workbook
            workbook.Save("SlicerToggleDemo.xlsx");
        }

        // Placeholder method representing a database query.
        // Replace with actual data access logic as needed.
        private static List<string> GetValuesFromDatabase()
        {
            // Example: suppose the query returns categories "A" and "C"
            return new List<string> { "A", "C" };
        }
    }
}