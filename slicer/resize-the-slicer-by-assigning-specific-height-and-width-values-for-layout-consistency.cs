// Title: Resize an Aspose.Cells Slicer in C# by Setting WidthPt and HeightPt
// Description: Shows how to create a workbook with a pivot table, add a linked slicer, and programmatically assign specific values to the slicer’s Shape.WidthPt and Shape.HeightPt properties for consistent layout before saving the file.
// Keywords: Aspose.Cells slicer resize | C# slicer width height | .NET Shape.WidthPt | .NET Shape.HeightPt | pivot table slicer size | adjust Excel slicer dimensions | set slicer width points | set slicer height points | Aspose.Cells programmatic slicer | Excel slicer layout control
// Common Searches: Aspose.Cells resize slicer C# | Set slicer width points Aspose.Cells .NET | How to change slicer height in Aspose.Cells | Programmatically adjust Excel slicer size using Aspose | Resize pivot table slicer with Aspose.Cells
// Developer Intent: Assign explicit width and height values to an Aspose.Cells slicer to control its visual size in a generated workbook.
// Use Cases: Standardize slicer dimensions across a batch of automatically generated reports. | Match slicer size to a corporate Excel template for brand‑consistent dashboards. | Dynamically compute slicer size based on worksheet column widths or screen resolution. | Create printable workbooks where slicers have uniform appearance for better readability.
// AI Prompts: Write C# code using Aspose.Cells to set a slicer's WidthPt to 150 and HeightPt to 120. | Show an example that adds a slicer to a pivot table and resizes it with Shape properties. | Explain how to access and modify a slicer's Shape object in Aspose.Cells. | Provide guidance on calculating slicer dimensions from worksheet column widths in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

// Shows how to create a workbook with a pivot table, add a linked slicer, and programmatically assign specific values to the slicer’s Shape.WidthPt and Shape.HeightPt properties for consistent layout before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].Value = "Category";
            worksheet.Cells["A2"].Value = "Fruit";
            worksheet.Cells["A3"].Value = "Fruit";
            worksheet.Cells["A4"].Value = "Vegetable";
            worksheet.Cells["B1"].Value = "Sales";
            worksheet.Cells["B2"].Value = 100;
            worksheet.Cells["B3"].Value = 150;
            worksheet.Cells["B4"].Value = 200;

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a slicer linked to the pivot table (destCellName first, then field name)
            int slicerIndex = worksheet.Slicers.Add(pivotTable, "E1", "Category");
            Slicer slicer = worksheet.Slicers[slicerIndex];
            slicer.AddPivotConnection(pivotTable);

            // Resize the slicer using the Shape's point-based properties
            slicer.Shape.WidthPt = 150;   // Set width to 150 points
            slicer.Shape.HeightPt = 120;  // Set height to 120 points

            // Save the workbook with the resized slicer
            workbook.Save("ResizedSlicer.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
