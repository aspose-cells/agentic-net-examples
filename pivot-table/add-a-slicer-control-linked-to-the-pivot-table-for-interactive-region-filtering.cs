// Title: Add a Region Slicer to a Pivot Table with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts Region and Sales data, builds a pivot table, adds a slicer linked to the Region field, applies a visual style and caption, and saves the file as PivotWithRegionSlicer.xlsx. The slicer provides an interactive filter for the pivot.
// Keywords: aspose.cells slicer c# | pivot table slicer asp.net | region slicer aspose.cells | interactive pivot filter c# | aspose.cells pivot example | c# add slicer to pivot | aspose.cells workbook slicer | pivot table filter control | aspose.cells style slicer | c# excel slicer example
// Common Searches: how to add a slicer to a pivot table using Aspose.Cells | Aspose.Cells C# region slicer example | set slicer style and caption in Aspose.Cells | interactive pivot table filter with slicer C# | Aspose.Cells add slicer to worksheet
// Developer Intent: Generate a slicer control linked to a pivot table that lets users filter the Region field interactively.
// Use Cases: Enable end‑users to click a Region filter directly on the worksheet. | Refresh pivot data automatically when slicer selections change. | Customize the slicer's appearance with predefined styles and captions.
// AI Prompts: Write C# code to add a slicer for the "Category" field to an existing Aspose.Cells pivot table. | Show how to read the selected items from an Aspose.Cells slicer and programmatically refresh the pivot table. | Explain positioning multiple slicers on a worksheet and applying different visual styles using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // Creates a workbook, inserts Region and Sales data, builds a pivot table, adds a slicer linked to the Region field, applies a visual style and caption, and saves the file as PivotWithRegionSlicer.xlsx. The slicer provides an interactive filter for the pivot.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a "Region" field and a "Sales" field
            sheet.Cells["A1"].Value = "Region";
            sheet.Cells["B1"].Value = "Sales";

            sheet.Cells["A2"].Value = "North";
            sheet.Cells["B2"].Value = 1200;

            sheet.Cells["A3"].Value = "South";
            sheet.Cells["B3"].Value = 950;

            sheet.Cells["A4"].Value = "East";
            sheet.Cells["B4"].Value = 780;

            sheet.Cells["A5"].Value = "West";
            sheet.Cells["B5"].Value = 1100;

            sheet.Cells["A6"].Value = "North";
            sheet.Cells["B6"].Value = 1300;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot: Region as row field, Sales as data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Region");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a slicer linked to the pivot table for the "Region" field
            // The slicer will be placed with its upper‑left corner at cell G3
            int slicerIndex = sheet.Slicers.Add(pivot, "G3", "Region");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Optional: set a visual style for the slicer
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.Caption = "Region Filter";

            // Save the workbook to a file
            workbook.Save("PivotWithRegionSlicer.xlsx");
        }
    }
}
