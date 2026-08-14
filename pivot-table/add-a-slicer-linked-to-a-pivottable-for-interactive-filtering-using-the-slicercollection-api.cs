// Title: C# – Add a Slicer to a PivotTable using Aspose.Cells SlicerCollection API
// Description: Learn how to create a workbook, build a PivotTable, and attach a slicer that filters the "Fruit" field. The example shows positioning the slicer at cell E2, customizing its caption, style, column count, and size, then saving the file as PivotTableWithSlicer.xlsx.
// Keywords: Aspose.Cells slicer example | C# add slicer to pivot table | Aspose.Cells SlicerCollection API | PivotTable slicer C# | customize slicer appearance Aspose | save workbook with slicer | .NET Excel slicer tutorial
// Common Searches: how to add a slicer to a pivot table with Aspose.Cells | Aspose.Cells C# slicer linked to pivot table | set slicer position and style Aspose.Cells | Aspose.Cells example for slicer collection | C# code to create pivot table and slicer
// Developer Intent: Generate a slicer linked to a PivotTable for interactive Excel filtering via Aspose.Cells in C#.
// Use Cases: Enable end‑users to filter a PivotTable by fruit type with a clickable slicer. | Match corporate report design by customizing slicer caption, style, column layout, and pixel dimensions. | Automate workbook generation that includes pre‑configured slicers for downstream analysis.
// AI Prompts: Show how to add multiple slicers for different PivotTable fields using Aspose.Cells in C#. | Provide code to programmatically select or deselect slicer items at runtime with Aspose.Cells. | Explain how to apply custom styles and arrange slicers on a worksheet using the SlicerCollection API.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // Learn how to create a workbook, build a PivotTable, and attach a slicer that filters the "Fruit" field. The example shows positioning the slicer at cell E2, customizing its caption, style, column count, and size, then saving the file as PivotTableWithSlicer.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Quantity";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 10;
            cells["A3"].Value = "Orange";
            cells["B3"].Value = 5;
            cells["A4"].Value = "Banana";
            cells["B4"].Value = 8;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the "Fruit" field to the row area and "Quantity" to the data area
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Refresh and calculate the pivot table to ensure it has data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table.
            // The slicer will be placed with its upper‑left corner at cell E2
            // and will filter based on the "Fruit" field.
            int slicerIndex = sheet.Slicers.Add(pivot, "E2", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Optional: customize slicer appearance
            slicer.Caption = "Fruit Filter";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.NumberOfColumns = 1;
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 120;

            // Save the workbook to a file
            workbook.Save("PivotTableWithSlicer.xlsx");
        }
    }
}
