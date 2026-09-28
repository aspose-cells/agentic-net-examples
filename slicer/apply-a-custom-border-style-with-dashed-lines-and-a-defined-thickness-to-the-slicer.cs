// Title: Apply a 2‑point dashed border to an Excel slicer with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a pivot table slicer and sets its Shape.Line.Weight to 2 points and Line.DashStyle to Dash using Aspose.Cells. | Show how to enable slicer border visibility and assign a dashed line style through the Shape object in Aspose.Cells. | Provide an example that adjusts slicer dimensions with Shape.WidthPt/HeightPt while applying a custom dashed border.
// Common Searches: C# Aspose.Cells how to set slicer border dash style and thickness | example of customizing Excel slicer line weight with Aspose.Cells for .NET | apply dashed line to pivot table slicer programmatically using Aspose.Cells | Aspose.Cells set slicer shape line properties in C# | change slicer border appearance in Excel workbook with Aspose.Cells API
// Tags: Aspose.Cells slicer border dash style | C# slicer shape line weight | Excel slicer custom line formatting Aspose | pivot table slicer line properties C# | Aspose.Cells shape line dash configuration

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace AsposeCellsSlicerBorderDemo
{
    // Demonstrates creating a workbook with a pivot table, adding a slicer, and using the slicer's Shape object to enable a visible border, set a 2‑point dashed line style, adjust size, and save the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
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

                // Add a pivot table based on the data
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "E3", "FruitPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);
                pivot.AddFieldToArea(PivotFieldType.Data, 1);

                // Add a slicer linked to the pivot table
                int slicerIdx = sheet.Slicers.Add(pivot, "A1", "FruitSlicer");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Access the underlying shape of the slicer
                Shape slicerShape = slicer.Shape;

                // Ensure the line (border) is visible
                slicerShape.HasLine = true;

                // Set custom border: dashed line with defined thickness (weight)
                // Note: Color and visibility properties may not be available in some versions,
                // so they are omitted to maintain compatibility.
                slicerShape.Line.Weight = 2.0;                         // Thickness (points)
                slicerShape.Line.DashStyle = MsoLineDashStyle.Dash;    // Dashed style

                // Optional: set other slicer properties
                slicer.Caption = "Select Fruit";
                slicerShape.WidthPt = 200;   // Use Shape.WidthPt instead of obsolete Slicer.Width
                slicerShape.HeightPt = 120;  // Use Shape.HeightPt instead of obsolete Slicer.Height

                // Save the workbook (lifecycle rule: save)
                workbook.Save("SlicerCustomBorderDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
