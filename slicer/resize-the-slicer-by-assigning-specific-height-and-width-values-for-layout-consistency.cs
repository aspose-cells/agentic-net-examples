// Title: Resize an Excel slicer and lock its position using Aspose.Cells for .NET (C#)
// AI Prompts: Set the slicer width to 180 points and height to 120 points, then lock its position in a workbook with Aspose.Cells C#. | Programmatically adjust a pivot table slicer's dimensions and prevent users from moving it using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to change slicer size in points | C# lock slicer position in generated Excel file using Aspose.Cells | Resize pivot table slicer width height Aspose.Cells .NET example | Set slicer dimensions and keep layout consistent Aspose.Cells C#
// Tags: slicer shape width points Aspose.Cells | slicer shape height points Aspose.Cells | prevent slicer movement Aspose.Cells | pivot table slicer size adjustment .NET | excel slicer layout consistency C#

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerResizeDemo
{
    // Creates a workbook, adds a pivot table, inserts a slicer linked to the "Category" field, resizes the slicer to 180 pt width and 120 pt height, locks its position to prevent user changes, and saves the file as SlicerResized.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for a pivot table
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Amount";
                cells["A2"].Value = "Fruit";
                cells["B2"].Value = 120;
                cells["A3"].Value = "Fruit";
                cells["B3"].Value = 150;
                cells["A4"].Value = "Vegetable";
                cells["B4"].Value = 200;

                // Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Add a slicer linked to the pivot table (field name "Category")
                int slicerIdx = sheet.Slicers.Add(pivot, "Category", "F3");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Resize the slicer using Shape properties (points)
                slicer.Shape.WidthPt = 180;   // width in points
                slicer.Shape.HeightPt = 120;  // height in points

                // Optionally lock the slicer position so users cannot move/resize it
                slicer.LockedPosition = true;

                // Save the workbook
                workbook.Save("SlicerResized.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
