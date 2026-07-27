// Title: Aspose.Cells C# – Create a Shape Summary Worksheet with Types, Positions, and Adjustment Values
// Description: C# example that builds a new workbook, adds sample rectangle and chevron auto‑shapes, then creates a "Summary" sheet. The code writes headers, iterates through every non‑summary worksheet, extracts each shape’s type, upper‑left and lower‑right row/column indices, and any geometry adjustment values, records the data, auto‑fits columns, and saves the file as ShapeSummaryReport.xlsx. Ideal for auditing or exporting shape metadata in Excel using Aspose.Cells.
// Keywords: Aspose.Cells C# | list shape properties | shape coordinates Aspose.Cells | auto shape adjustment values | Excel shape summary report | iterate shapes workbook | shape metadata export | C# generate shape report
// Common Searches: list all shapes in an Excel workbook using Aspose.Cells | retrieve shape coordinates and adjustment guides with Aspose.Cells .NET | create a summary sheet of shape types and positions in C# | export shape metadata to Excel using Aspose.Cells
// Developer Intent: Generate an Excel workbook that includes a summary sheet enumerating each shape’s type, cell coordinates, and any adjustment values.
// Use Cases: Validate placement of shapes in a template before distribution | Document shape metadata for design or compliance audits | Feed shape position data into downstream automation scripts | Compare shape adjustments across multiple worksheets
// AI Prompts: Write C# code with Aspose.Cells that adds a summary worksheet listing each shape’s type, upper‑left row/column, lower‑right row/column, and adjustment values for every shape in a workbook. | Show how to extract adjustment guide values from an auto‑shape and concatenate them into a semicolon‑separated string on a summary sheet using Aspose.Cells. | Demonstrate skipping the summary sheet while looping through worksheets, writing shape data, applying AutoFitColumns, and saving the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeSummaryReport
{
    // C# example that builds a new workbook, adds sample rectangle and chevron auto‑shapes, then creates a "Summary" sheet. The code writes headers, iterates through every non‑summary worksheet, extracts each shape’s type, upper‑left and lower‑right row/column indices, and any geometry adjustment values, records the data, auto‑fits columns, and saves the file as ShapeSummaryReport.xlsx. Ideal for auditing or exporting shape metadata in Excel using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample shapes to the first worksheet for demonstration
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Rectangle shape
            Shape rect = sheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 200);
            rect.Name = "MyRectangle";

            // Auto shape with adjustment values (e.g., Chevron)
            Shape auto = sheet.Shapes.AddAutoShape(AutoShapeType.Chevron, 5, 0, 5, 0, 150, 80);
            // Modify an adjustment value if available
            if (auto.Geometry.ShapeAdjustValues.Count > 0)
            {
                auto.Geometry.ShapeAdjustValues[0].Value = 0.4;
            }

            // Add a summary worksheet
            Worksheet summary = workbook.Worksheets[workbook.Worksheets.Add()];
            summary.Name = "Summary";

            // Write header row
            summary.Cells["A1"].PutValue("Worksheet");
            summary.Cells["B1"].PutValue("Shape Index");
            summary.Cells["C1"].PutValue("Shape Type");
            summary.Cells["D1"].PutValue("Upper Left Row");
            summary.Cells["E1"].PutValue("Upper Left Column");
            summary.Cells["F1"].PutValue("Lower Right Row");
            summary.Cells["G1"].PutValue("Lower Right Column");
            summary.Cells["H1"].PutValue("Adjustment Values");

            int summaryRow = 1; // zero‑based index; row 1 is the second row (after header)

            // Iterate through all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Skip the summary sheet itself
                if (ws == summary) continue;

                ShapeCollection shapes = ws.Shapes;

                // Iterate through each shape in the worksheet
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Basic shape information
                    string shapeType = shape.Type.ToString();
                    int upperLeftRow = shape.UpperLeftRow;
                    int upperLeftColumn = shape.UpperLeftColumn;
                    int lowerRightRow = shape.LowerRightRow;
                    int lowerRightColumn = shape.LowerRightColumn;

                    // Collect adjustment values (if any)
                    string adjustments = "";
                    Geometry geometry = shape.Geometry;
                    if (geometry != null && geometry.ShapeAdjustValues != null && geometry.ShapeAdjustValues.Count > 0)
                    {
                        foreach (ShapeGuide guide in geometry.ShapeAdjustValues)
                        {
                            // Guide.Name may be null; use the formula/value representation
                            adjustments += $"{guide.Value}; ";
                        }
                        // Trim trailing separator
                        adjustments = adjustments.TrimEnd(' ', ';');
                    }
                    else
                    {
                        adjustments = "N/A";
                    }

                    // Write data to the summary sheet
                    summary.Cells[summaryRow, 0].PutValue(ws.Name);
                    summary.Cells[summaryRow, 1].PutValue(i);
                    summary.Cells[summaryRow, 2].PutValue(shapeType);
                    summary.Cells[summaryRow, 3].PutValue(upperLeftRow);
                    summary.Cells[summaryRow, 4].PutValue(upperLeftColumn);
                    summary.Cells[summaryRow, 5].PutValue(lowerRightRow);
                    summary.Cells[summaryRow, 6].PutValue(lowerRightColumn);
                    summary.Cells[summaryRow, 7].PutValue(adjustments);

                    summaryRow++;
                }
            }

            // Auto‑fit columns for better readability
            summary.AutoFitColumns();

            // Save the workbook
            workbook.Save("ShapeSummaryReport.xlsx");
        }
    }
}
