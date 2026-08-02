// Title: Aspose.Cells for .NET (C#): Add a Pivot Table Slicer and Lock Its Position
// Description: C# example that creates a workbook, fills sample sales data, builds a pivot table, inserts a slicer linked to the "Category" field at cell G3, customizes its caption, size and column layout, then locks the slicer so it cannot be moved or resized before saving the file.
// Keywords: Aspose.Cells slicer pivot table C# | lock slicer position Aspose.Cells | freeze slicer Aspose.Cells .NET | add slicer to pivot table Aspose | C# workbook slicer example
// Common Searches: Aspose.Cells add slicer to pivot table C# | How to lock slicer position in Aspose.Cells | Prevent slicer movement in Excel using Aspose.Cells | C# code for pivot table slicer with fixed location | Aspose.Cells freeze slicer row area
// Developer Intent: Insert a slicer that is linked to a pivot table and make its position immutable in a .NET workbook.
// Use Cases: Design a sales dashboard where the category filter remains visible and fixed for all users. | Automate report generation that includes a non‑movable slicer to preserve layout consistency across devices. | Create shared workbooks with locked slicers to prevent accidental repositioning during collaborative editing.
// AI Prompts: Generate C# code with Aspose.Cells to add a slicer for the "Region" field of an existing pivot table and lock its position. | Show how to lock multiple slicers on a worksheet using Aspose.Cells for .NET. | Explain how to set slicer size, column count, and lock its position while aligning it with a pivot table in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // C# example that creates a workbook, fills sample sales data, builds a pivot table, inserts a slicer linked to the "Category" field at cell G3, customizes its caption, size and column layout, then locks the slicer so it cannot be moved or resized before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Product";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Fruits";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 120;

            cells["A3"].Value = "Fruits";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 80;

            cells["A4"].Value = "Vegetables";
            cells["B4"].Value = "Carrot";
            cells["C4"].Value = 150;

            cells["A5"].Value = "Vegetables";
            cells["B5"].Value = "Tomato";
            cells["C5"].Value = 90;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Category, data = Sales
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table for the "Category" field
            // The slicer will be placed with its top‑left corner at cell G3
            int slicerIndex = sheet.Slicers.Add(pivot, "G3", "Category");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Set slicer properties
            slicer.Caption = "Category Filter";
            slicer.NumberOfColumns = 1;
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 100;

            // Freeze the slicer’s position so it cannot be moved or resized by the user
            slicer.LockedPosition = true;

            // Save the workbook
            workbook.Save("SlicerWithLockedPosition.xlsx");
        }
    }
}
