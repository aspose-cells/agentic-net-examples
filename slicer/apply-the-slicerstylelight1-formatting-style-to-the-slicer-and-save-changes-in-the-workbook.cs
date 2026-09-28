// Title: How to apply the built‑in SlicerStyleLight1 to an Excel slicer and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds a pivot table, inserts a slicer linked to the pivot, sets the slicer’s StyleType to SlicerStyleLight1, and saves the file. | Show how to format an Excel slicer with the built‑in Light1 style using the Aspose.Cells API in a .NET application.
// Common Searches: Aspose.Cells C# set slicer style to SlicerStyleLight1 | C# example for applying built‑in slicer style in Excel with Aspose | How to change slicer appearance using Aspose.Cells .NET | Save workbook after styling slicer with Aspose.Cells API
// Tags: Aspose.Cells slicer style application | C# slicer StyleType Light1 | Excel slicer formatting with Aspose | save workbook after slicer styling .NET | pivot table slicer creation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerStyleDemo
{
    // // Creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the pivot, applies the built‑in SlicerStyleLight1 to the slicer, and saves the workbook as SlicerStyleLight1Demo.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Fruit");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["B4"].PutValue(20);

            // Add a pivot table based on the sample data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "E3", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity column

            // Add a slicer linked to the pivot table's first base field (Fruit)
            int slicerIdx = sheet.Slicers.Add(pivot, "A1", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Apply the built‑in light style 1 to the slicer
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Save the workbook with the styled slicer
            workbook.Save("SlicerStyleLight1Demo.xlsx");
        }
    }
}
