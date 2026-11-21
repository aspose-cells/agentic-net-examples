// Title: Generate a worksheet report of all shape types and their cell addresses in an Excel workbook using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# list all shapes with their type and cell location in each worksheet | How to create a shape summary sheet that shows shape name, type, and upper‑left cell address using Aspose.Cells | C# code to export shape details from an Excel file to a new report worksheet with Aspose.Cells | Retrieve shape coordinates and types from all worksheets in a workbook with Aspose.Cells for .NET | Generate Excel shape inventory report programmatically using Aspose.Cells API
// Tags: enumerate worksheet shapes Aspose.Cells C# | export shape metadata to report sheet Aspose.Cells | get shape upper‑left cell address Aspose.Cells | create shape type inventory Excel Aspose.Cells | list shape names and locations Aspose.Cells .NET

using Aspose.Cells;
using System;

// The example loads an existing workbook, adds a new worksheet named "Shape Report", writes headers, then iterates through every worksheet (excluding the report sheet) and each shape on those sheets. For each shape it captures the worksheet name, shape name, shape type, and the cell address of the shape's upper‑left corner, writes these details to the report sheet, auto‑fits the columns, and saves the workbook with the added report.
class ShapeReport
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Add a new worksheet that will contain the shape report
        int reportSheetIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
        reportSheet.Name = "Shape Report";

        // Write header row
        reportSheet.Cells[0, 0].PutValue("Worksheet");
        reportSheet.Cells[0, 1].PutValue("Shape Name");
        reportSheet.Cells[0, 2].PutValue("Shape Type");
        reportSheet.Cells[0, 3].PutValue("Cell Address");

        int reportRow = 1; // start writing data from the second row

        // Iterate through all worksheets in the workbook
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Skip the report sheet itself to avoid self‑reference
            if (ws == reportSheet) continue;

            // Iterate through all shapes in the current worksheet
            foreach (Aspose.Cells.Drawing.Shape shape in ws.Shapes)
            {
                // Shape name (if not set, Aspose provides a default name)
                string shapeName = shape.Name;

                // Shape type as string
                string shapeType = shape.Type.ToString();

                // Upper‑left cell coordinates of the shape
                int topRow = shape.UpperLeftRow;
                int leftColumn = shape.UpperLeftColumn;
                string cellAddress = CellsHelper.CellIndexToName(topRow, leftColumn);

                // Write the collected information into the report sheet
                reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                reportSheet.Cells[reportRow, 1].PutValue(shapeName);
                reportSheet.Cells[reportRow, 2].PutValue(shapeType);
                reportSheet.Cells[reportRow, 3].PutValue(cellAddress);

                reportRow++;
            }
        }

        // Adjust column widths for better readability
        reportSheet.AutoFitColumns();

        // Save the workbook with the newly added report (replace with desired output path)
        workbook.Save("output_with_shape_report.xlsx");
    }
}
