// Title: Clone an existing slicer and add the copy to a different worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new slicer on a target worksheet, links it to the same pivot table, and copies all visual settings from an existing slicer using Aspose.Cells. | Show how to replicate a slicer's caption, style, size, and behavior on another sheet by cloning its properties with Aspose.Cells in a .NET project.
// Common Searches: aspnet cells duplicate slicer to another worksheet c# example | copy slicer properties from one sheet to another using Aspose.Cells | how to clone a slicer linked to a pivot table in C# Aspose.Cells | Aspose.Cells create slicer on target sheet with same settings as source slicer
// Tags: slicer duplication Aspose.Cells C# | slicer property copy Aspose.Cells | add slicer to another worksheet Aspose.Cells | pivot table slicer cloning .NET | slicer visual settings transfer Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerCloneDemo
{
    // The example creates a workbook with sample data, adds a pivot table, inserts a slicer on a source sheet, then adds a second worksheet, creates a new slicer linked to the same pivot table, copies visual properties (caption, style, size, etc.) from the original slicer, and saves the file as SlicerCloneDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare source worksheet with data, pivot table and slicer
            // -------------------------------------------------
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Populate sample data
            sourceSheet.Cells["A1"].PutValue("Fruit");
            sourceSheet.Cells["B1"].PutValue("Sales");
            sourceSheet.Cells["A2"].PutValue("Apple");
            sourceSheet.Cells["B2"].PutValue(120);
            sourceSheet.Cells["A3"].PutValue("Orange");
            sourceSheet.Cells["B3"].PutValue(150);
            sourceSheet.Cells["A4"].PutValue("Banana");
            sourceSheet.Cells["B4"].PutValue(90);

            // Add a pivot table based on the data
            int pivotIdx = sourceSheet.PivotTables.Add("A1:B4", "D2", "FruitPivot");
            PivotTable pivot = sourceSheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add the original slicer on the source sheet
            int slicerIdx = sourceSheet.Slicers.Add(pivot, "F2", "Fruit");
            Slicer originalSlicer = sourceSheet.Slicers[slicerIdx];
            originalSlicer.Caption = "Fruit Slicer";
            originalSlicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            originalSlicer.NumberOfColumns = 1;

            // -------------------------------------------------
            // 2. Add a target worksheet where the slicer copy will be placed
            // -------------------------------------------------
            Worksheet targetSheet = workbook.Worksheets.Add("Target");

            // -------------------------------------------------
            // 3. Create a new slicer on the target sheet using the same pivot table and field
            // -------------------------------------------------
            int copySlicerIdx = targetSheet.Slicers.Add(pivot, "F2", "Fruit");
            Slicer copySlicer = targetSheet.Slicers[copySlicerIdx];

            // -------------------------------------------------
            // 4. Clone visual properties from the original slicer to the copy
            // -------------------------------------------------
            copySlicer.Caption = originalSlicer.Caption;
            copySlicer.StyleType = originalSlicer.StyleType;
            copySlicer.NumberOfColumns = originalSlicer.NumberOfColumns;
            copySlicer.WidthPixel = originalSlicer.WidthPixel;
            copySlicer.HeightPixel = originalSlicer.HeightPixel;
            copySlicer.LockedPosition = originalSlicer.LockedPosition;
            copySlicer.ShowCaption = originalSlicer.ShowCaption;
            copySlicer.ShowAllItems = originalSlicer.ShowAllItems;

            // -------------------------------------------------
            // 5. Save the workbook
            // -------------------------------------------------
            workbook.Save("SlicerCloneDemo.xlsx");
        }
    }
}
