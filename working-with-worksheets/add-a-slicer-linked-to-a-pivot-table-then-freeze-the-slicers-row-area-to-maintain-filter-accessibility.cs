// Title: Add a Pivot Table Slicer and Lock Its Position with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate sample data, build a pivot table, add a slicer linked to the "Category" field, customize its caption, size and layout, and then freeze the slicer’s position using the LockedPosition property so it remains fixed on the worksheet.
// Keywords: Aspose.Cells slicer | C# pivot table slicer | lock slicer position | freeze slicer Aspose.Cells | LockedPosition property | pivot table dashboard .NET | Aspose.Cells add slicer example | Excel slicer programmatically | Aspose.Cells worksheet UI
// Common Searches: Aspose.Cells add slicer to pivot table C# | How to lock slicer position in Aspose.Cells | Set slicer LockedPosition property .NET | Create pivot table with slicer using Aspose.Cells | Freeze slicer on worksheet Aspose
// Developer Intent: Create a slicer linked to a pivot table and lock its placement to keep the filter always accessible.
// Use Cases: Design a read‑only dashboard where the category filter stays visible. | Generate financial or sales reports with a fixed slicer layout for consistent sharing. | Build interactive worksheets that require slicers to remain stationary during scrolling or printing.
// AI Prompts: Generate C# code with Aspose.Cells that adds a slicer to a pivot table and locks its position. | Show how to configure slicer properties such as Caption, WidthPixel, HeightPixel, NumberOfColumns, and LockedPosition in Aspose.Cells. | Explain step‑by‑step how to create a pivot table, attach a slicer, and freeze the slicer for a dashboard worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // Demonstrates how to create a workbook, populate sample data, build a pivot table, add a slicer linked to the "Category" field, customize its caption, size and layout, and then freeze the slicer’s position using the LockedPosition property so it remains fixed on the worksheet.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Fruit";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Vegetable";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Fruit";
            cells["B4"].Value = 150;
            cells["A5"].Value = "Vegetable";
            cells["B5"].Value = 70;

            // Add a pivot table based on the data range A1:B5, place it at C3
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "C3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row field, Amount as data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table for the "Category" field.
            // The slicer will be placed with its upper‑left corner at cell E3.
            int slicerIndex = sheet.Slicers.Add(pivot, "E3", "Category");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Set slicer properties
            slicer.Caption = "Category Filter";
            slicer.NumberOfColumns = 1;
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 100;

            // Freeze the slicer’s position so it cannot be moved or resized by the user.
            slicer.LockedPosition = true;

            // Save the workbook
            workbook.Save("SlicerWithLockedPosition.xlsx");
        }
    }
}
