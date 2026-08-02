using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsSmartMarkersMergedDemo
{
    // Sample data class
    public class Item
    {
        public string Name { get; set; }
        public double Value { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (template)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Create a merged header (A1:C1)
            cells.Merge(0, 0, 1, 3); // Row 0, Column 0, 1 row, 3 columns
            cells[0, 0].PutValue("Report Title");

            // 3. Place smart markers for data rows starting at row 2 (A2 and B2)
            cells["A2"].PutValue("&=$Items.Name");
            cells["B2"].PutValue("&=$Items.Value");

            // 4. Create a placeholder merged cell that should expand with data rows (column D)
            // Initially merge only the first data row cell (D2)
            cells.Merge(1, 3, 1, 1); // Row 1 (A2), Column 3 (D), 1 row, 1 column
            cells[1, 3].PutValue("Static Info");

            // 5. Prepare sample data source
            List<Item> items = new List<Item>
            {
                new Item { Name = "Item A", Value = 123.45 },
                new Item { Name = "Item B", Value = 678.90 },
                new Item { Name = "Item C", Value = 234.56 },
                new Item { Name = "Item D", Value = 789.01 }
            };

            // 6. Process smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Items", items);
            designer.Process(); // Populate the smart markers

            // 7. After processing, determine the last row that contains data
            int lastDataRow = sheet.Cells.MaxDataRow; // zero‑based index

            // 8. Expand the previously merged range (column D) to cover all data rows
            // First unmerge the original single‑cell range
            cells.UnMerge(1, 3, 1, 1);
            // Then merge from the first data row (row 1) down to the last data row
            int rowsToMerge = lastDataRow - 1 + 1; // include both start and end rows
            cells.Merge(1, 3, rowsToMerge, 1);

            // 9. Auto‑fit rows, ensuring merged cells are considered
            AutoFitterOptions fitOptions = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine
            };
            sheet.AutoFitRows(fitOptions);

            // 10. Save the resulting workbook
            workbook.Save("SmartMarkersMergedOutput.xlsx");
        }
    }
}