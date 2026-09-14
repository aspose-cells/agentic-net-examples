// Title: Add a slicer linked to a PivotTable in C# with Aspose.Cells for interactive Excel filtering
// AI Prompts: Generate C# code that creates a workbook, builds a PivotTable from a data range, and inserts a slicer bound to a chosen field using the Aspose.Cells SlicerCollection API. | Demonstrate how to set the slicer's caption, style, column count, width, and height after it is linked to a PivotTable. | Explain how to add additional slicers for other PivotTable fields and position each slicer on different worksheets programmatically.
// Common Searches: aspnet add slicer to pivot table programmatically using Aspose.Cells | c# example of linking Excel slicer to a pivot table with Aspose.Cells | how to customize slicer style and size in Aspose.Cells .NET | multiple slicers for different fields in a pivot table using Aspose.Cells C# | Aspose.Cells SlicerCollection API usage for interactive Excel reports
// Tags: Aspose.Cells add slicer to pivot table | C# slicer customization Aspose.Cells | interactive Excel filtering with slicer | SlicerCollection API Aspose.Cells | pivot table slicer placement C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // Shows how to create a workbook, populate sample data, add a PivotTable, insert a slicer linked to the "Fruit" field, customize its caption, style, column layout, width, and height, and save the file as PivotTableWithSlicer.xlsx using Aspose.Cells for .NET.
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
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            // Place the "Fruit" field in the row area and "Quantity" in the data area
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Add a slicer linked to the pivot table.
            // The slicer will be placed with its upper‑left corner at cell E2
            // and will filter by the "Fruit" field.
            int slicerIdx = sheet.Slicers.Add(pivot, "E2", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Optional: customize slicer appearance
            slicer.Caption = "Fruit Filter";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.NumberOfColumns = 1;
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 120;

            // Save the workbook
            workbook.Save("PivotTableWithSlicer.xlsx");
        }
    }
}
