// Title: C# – Generate a Shape Report with Types and Cell Coordinates Using Aspose.Cells
// Description: Loads an Excel workbook, adds a "ShapeReport" worksheet, enumerates every shape on each sheet (except the report), records the shape name, drawing type, upper‑left and lower‑right row/column indices, and saves the workbook with the report.
// Keywords: Aspose.Cells | C# | list shapes in workbook | shape coordinates | MsoDrawingType | Excel shape report | enumerate drawings | shape metadata export | worksheet shapes | export shape data
// Common Searches: Aspose.Cells list all shapes in workbook | C# get shape coordinates Excel | How to export shape types with Aspose.Cells | Create shape inventory sheet Aspose.Cells | Retrieve drawing positions from Excel using .NET
// Developer Intent: Create an automatic worksheet that inventories every shape in a workbook, showing its name, drawing type and the cell range it occupies.
// Use Cases: Audit and document all drawings in an Excel file. | Validate that shapes stay within designated cell boundaries. | Migrate or refactor spreadsheets by extracting shape metadata. | Generate documentation for template designers. | Support automated testing of Excel layouts.
// AI Prompts: Write C# code with Aspose.Cells that adds a summary sheet listing each shape’s name, type, and its upper‑left and lower‑right row/column indices for every worksheet. | Modify the example to also record each shape’s width and height in points alongside its coordinates. | Create a version that excludes pictures (or any specific MsoDrawingType) from the shape report. | Generate code that exports the shape report to a CSV file instead of an Excel worksheet. | Provide a script that adds hyperlinks from the report rows back to the original shapes in the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, adds a "ShapeReport" worksheet, enumerates every shape on each sheet (except the report), records the shape name, drawing type, upper‑left and lower‑right row/column indices, and saves the workbook with the report.
class ShapeReport
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Add a new worksheet to hold the report
        int reportIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportIndex];
        reportSheet.Name = "ShapeReport";

        // Write header row
        reportSheet.Cells["A1"].PutValue("Worksheet");
        reportSheet.Cells["B1"].PutValue("Shape Name");
        reportSheet.Cells["C1"].PutValue("Shape Type");
        reportSheet.Cells["D1"].PutValue("Upper Left Row");
        reportSheet.Cells["E1"].PutValue("Upper Left Column");
        reportSheet.Cells["F1"].PutValue("Lower Right Row");
        reportSheet.Cells["G1"].PutValue("Lower Right Column");

        int row = 1; // zero‑based index for the next data row

        // Iterate through all worksheets (except the report sheet itself)
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Name == "ShapeReport") continue;

            // Iterate through each shape in the current worksheet
            foreach (Shape shape in ws.Shapes)
            {
                // Populate the report with shape details
                reportSheet.Cells[row, 0].PutValue(ws.Name);
                reportSheet.Cells[row, 1].PutValue(shape.Name);
                reportSheet.Cells[row, 2].PutValue(shape.MsoDrawingType.ToString());
                reportSheet.Cells[row, 3].PutValue(shape.UpperLeftRow);
                reportSheet.Cells[row, 4].PutValue(shape.UpperLeftColumn);
                reportSheet.Cells[row, 5].PutValue(shape.LowerRightRow);
                reportSheet.Cells[row, 6].PutValue(shape.LowerRightColumn);
                row++;
            }
        }

        // Save the workbook with the generated report
        workbook.Save("ShapeReport.xlsx");
    }
}
