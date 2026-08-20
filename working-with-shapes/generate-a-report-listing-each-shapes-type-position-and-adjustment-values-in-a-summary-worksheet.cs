// Title: Create a Shape Summary Sheet with Type, Position & Adjustment Values using Aspose.Cells for .NET (C#)
// Description: C# program that builds an Excel workbook, adds sample rectangle and chevron shapes, then generates a "Shape Summary" worksheet. It scans every worksheet, extracts each shape's name, type, cell coordinates (upper‑left and lower‑right), and any geometry adjustment values, writes the data to the summary sheet, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | shape summary | shape metadata | shape position | adjustment values | auto shape | ShapeGuide | Excel automation | worksheet report
// Common Searches: Aspose.Cells list all shapes in a workbook | How to get shape coordinates with Aspose.Cells C# | Retrieve auto‑shape adjustment values using Aspose.Cells | Create a summary worksheet for shapes in Excel via .NET | Iterate through worksheets and shapes Aspose.Cells example
// Developer Intent: Generate an Excel file that contains a summary sheet listing each shape’s type, location and adjustment data.
// Use Cases: Document and audit drawing objects across complex spreadsheets | Export shape metadata for downstream processing or reporting | Validate consistency of shape adjustments in multi‑sheet workbooks | Provide a quick reference for designers reviewing Excel drawings
// AI Prompts: Write C# code with Aspose.Cells that adds a summary sheet showing shape names, types, positions and adjustment values for all shapes in a workbook. | Explain how to read ShapeGuide adjustment values from auto shapes and handle shapes without adjustments. | Suggest formatting options (headers, column widths, styles) for the shape summary worksheet created with Aspose.Cells.

using System;
using System.Text;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeSummaryReport
{
    // C# program that builds an Excel workbook, adds sample rectangle and chevron shapes, then generates a "Shape Summary" worksheet. It scans every worksheet, extracts each shape's name, type, cell coordinates (upper‑left and lower‑right), and any geometry adjustment values, writes the data to the summary sheet, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // (Optional) Add some sample shapes to demonstrate
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "DataSheet";

                // Add a rectangle
                Shape rect = sheet1.Shapes.AddRectangle(2, 2, 2, 2, 100, 200);
                rect.Name = "MyRectangle";

                // Add a chevron auto shape with adjustment values
                Shape chevron = sheet1.Shapes.AddAutoShape(AutoShapeType.Chevron, 5, 5, 5, 5, 150, 80);
                // Modify an adjustment value (e.g., roundness)
                if (chevron.Geometry.ShapeAdjustValues.Count > 0)
                {
                    chevron.Geometry.ShapeAdjustValues[0].Value = 0.4;
                }

                // -------------------------------------------------
                // Create a summary worksheet
                // -------------------------------------------------
                int summaryIndex = workbook.Worksheets.Add();
                Worksheet summary = workbook.Worksheets[summaryIndex];
                summary.Name = "Shape Summary";

                // Write header row
                summary.Cells[0, 0].PutValue("Worksheet");
                summary.Cells[0, 1].PutValue("Shape Name");
                summary.Cells[0, 2].PutValue("Shape Type");
                summary.Cells[0, 3].PutValue("Position (UpperLeftRow, UpperLeftColumn, LowerRightRow, LowerRightColumn)");
                summary.Cells[0, 4].PutValue("Adjustment Values");

                int currentRow = 1; // start after header

                // -------------------------------------------------
                // Iterate through all worksheets and their shapes
                // -------------------------------------------------
                for (int wsIdx = 0; wsIdx < workbook.Worksheets.Count; wsIdx++)
                {
                    Worksheet ws = workbook.Worksheets[wsIdx];
                    ShapeCollection shapes = ws.Shapes;

                    for (int i = 0; i < shapes.Count; i++)
                    {
                        Shape shape = shapes[i];

                        // Basic shape information
                        string shapeName = shape.Name;
                        string shapeType = shape.Type.ToString();
                        string position = $"({shape.UpperLeftRow}, {shape.UpperLeftColumn}) - ({shape.LowerRightRow}, {shape.LowerRightColumn})";

                        // Collect adjustment values, if any
                        StringBuilder adjustBuilder = new StringBuilder();
                        foreach (ShapeGuide guide in shape.Geometry.ShapeAdjustValues)
                        {
                            if (adjustBuilder.Length > 0) adjustBuilder.Append("; ");
                            // ShapeGuide may not expose a name property in some versions; output only the value
                            adjustBuilder.Append($"{guide.Value}");
                        }
                        string adjustments = adjustBuilder.Length > 0 ? adjustBuilder.ToString() : "N/A";

                        // Write to summary sheet
                        summary.Cells[currentRow, 0].PutValue(ws.Name);
                        summary.Cells[currentRow, 1].PutValue(shapeName);
                        summary.Cells[currentRow, 2].PutValue(shapeType);
                        summary.Cells[currentRow, 3].PutValue(position);
                        summary.Cells[currentRow, 4].PutValue(adjustments);

                        currentRow++;
                    }
                }

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "ShapeSummaryReport.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while generating the shape summary report:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
