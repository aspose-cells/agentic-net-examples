// Title: C# – Add a Dashed 2‑Point Border to an Aspose.Cells Slicer
// Description: Demonstrates how to create a workbook, build a pivot table, insert a linked slicer, and use the SlicerShape object to enable a line, set a 2‑point weight and a dash style, then save the file. Ideal for .NET developers who need custom slicer borders in automated Excel reports.
// Keywords: Aspose.Cells slicer border | C# slicer line weight | dash style slicer shape | custom slicer formatting .NET | pivot table slicer appearance | Excel dashboard slicer styling | Aspose.Cells SlicerShape line
// Common Searches: how to set a dashed border on an Aspose.Cells slicer | Aspose.Cells C# slicer line thickness | change slicer border style programmatically | apply custom line dash to Excel slicer using Aspose | set slicer shape line weight in .NET
// Developer Intent: Apply a dashed border with a specific thickness to a slicer created with Aspose.Cells for .NET.
// Use Cases: Generate Excel dashboards where slicers are visually highlighted with a consistent dashed outline. | Standardize slicer appearance across multiple reports to match corporate branding. | Create automated reporting pipelines that adjust slicer borders for better readability.
// AI Prompts: Write C# code to set a solid red 1.5‑point border on an Aspose.Cells slicer. | Show how to detect if the Line object supports a Color property and fall back to the default color. | Provide an example that switches slicer dash styles (Dot, DashDot, LongDash) based on a runtime flag.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace AsposeCellsSlicerBorderDemo
{
    // Demonstrates how to create a workbook, build a pivot table, insert a linked slicer, and use the SlicerShape object to enable a line, set a 2‑point weight and a dash style, then save the file. Ideal for .NET developers who need custom slicer borders in automated Excel reports.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Fruit");
                worksheet.Cells["A3"].PutValue("Fruit");
                worksheet.Cells["A4"].PutValue("Vegetable");
                worksheet.Cells["B1"].PutValue("Amount");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(15);
                worksheet.Cells["B4"].PutValue(8);

                // Add a pivot table based on the data
                int pivotIdx = worksheet.PivotTables.Add("A1:B4", "E1", "PivotTable1");
                PivotTable pivot = worksheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

                // Add a slicer linked to the pivot table's first field (Category)
                int slicerIdx = worksheet.Slicers.Add(pivot, "G1", pivot.BaseFields[0]);
                Slicer slicer = worksheet.Slicers[slicerIdx];

                // Access the underlying shape of the slicer
                SlicerShape slicerShape = slicer.Shape;

                // Apply a custom dashed border with defined thickness
                slicerShape.HasLine = true;                                 // Ensure the line (border) is visible
                // Note: In some Aspose.Cells versions the Line object does not expose a Color property.
                // The default line color will be used if setting the color is not supported.
                slicerShape.Line.Weight = 2.0;                              // Thickness in points
                slicerShape.Line.DashStyle = MsoLineDashStyle.Dash;         // Dashed line style

                // Optional: adjust slicer size and caption for better visibility
                slicer.Caption = "Category Filter";
                slicerShape.WidthPt = 200;   // Width in points
                slicerShape.HeightPt = 120;  // Height in points

                // Save the workbook
                workbook.Save("SlicerCustomBorderDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
