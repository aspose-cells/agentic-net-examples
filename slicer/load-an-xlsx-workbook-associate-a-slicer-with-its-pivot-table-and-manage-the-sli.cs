using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (assumed to contain the pivot table)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table; otherwise create a simple one for demo
        if (worksheet.PivotTables.Count == 0)
        {
            // Sample data for a pivot table
            worksheet.Cells["A1"].PutValue("Fruit");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(15);
            worksheet.Cells["A4"].PutValue("Banana");
            worksheet.Cells["B4"].PutValue(20);

            // Add a pivot table covering the data range
            int pivotIdx = worksheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivot = worksheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
        }

        // Retrieve the first pivot table
        PivotTable pivotTable = worksheet.PivotTables[0];

        // Add a slicer linked to the pivot table.
        // The slicer will be placed with its upper‑left corner at cell D1
        // and will filter the "Fruit" field.
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "D1", "Fruit");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Set some slicer properties
        slicer.Caption = "Fruit Filter";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer.NumberOfColumns = 2;
        slicer.WidthPixel = 200;
        slicer.HeightPixel = 150;

        // Modify underlying data to demonstrate Refresh()
        worksheet.Cells["A2"].PutValue("Mango"); // add a new fruit
        slicer.Refresh(); // refresh slicer and its pivot table

        // Optional: remove the slicer (uncomment to test)
        // worksheet.Slicers.RemoveAt(slicerIndex);

        // Save the workbook with the slicer changes
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}