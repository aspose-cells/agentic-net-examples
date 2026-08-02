// Title: C# Aspose.Cells: Generate a report of all shapes with their cell coordinates
// Description: Loads an existing workbook, walks through every worksheet and shape, captures each shape's name, MsoDrawingType, upper‑left and lower‑right cell addresses, and writes the data to a new Excel file (ShapeReport.xlsx).
// Keywords: Aspose.Cells | C# | list Excel shapes | shape coordinates | MsoDrawingType | UpperLeftRow | LowerRightColumn | shape report .NET | export shape data | Excel drawing objects
// Common Searches: Aspose.Cells list all shapes in workbook | C# get shape cell address Excel | export shape type and location Aspose | generate shape report with Aspose.Cells | how to retrieve shape coordinates in .NET
// Developer Intent: Create a worksheet that enumerates every shape in a workbook together with its type and bounding cells.
// Use Cases: Audit drawing placement in financial models to ensure compliance with layout standards. | Document visual elements in design workbooks for hand‑off to graphic teams. | Validate that inserted pictures, charts, or SmartArt occupy the intended cell range before publishing.
// AI Prompts: Write C# code using Aspose.Cells that adds each shape's width and height (in points) to the report. | Modify the example to filter and export only picture shapes (MsoDrawingType.Picture). | Create a version that groups shapes by worksheet and adds a summary row with the count per sheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeReportGenerator
{
    // Loads an existing workbook, walks through every worksheet and shape, captures each shape's name, MsoDrawingType, upper‑left and lower‑right cell addresses, and writes the data to a new Excel file (ShapeReport.xlsx).
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("input.xlsx");

            // Create a new workbook to hold the report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Worksheet");
            reportSheet.Cells[0, 1].PutValue("Shape Name");
            reportSheet.Cells[0, 2].PutValue("Shape Type");
            reportSheet.Cells[0, 3].PutValue("Upper Left Cell");
            reportSheet.Cells[0, 4].PutValue("Lower Right Cell");

            int reportRow = 1; // start writing data from the second row

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                // Iterate through each shape in the current worksheet
                foreach (Shape shape in ws.Shapes)
                {
                    // Get cell addresses for the shape's bounding box
                    string upperLeftAddress = ws.Cells[shape.UpperLeftRow, shape.UpperLeftColumn].Name;
                    string lowerRightAddress = ws.Cells[shape.LowerRightRow, shape.LowerRightColumn].Name;

                    // Populate the report sheet
                    reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                    reportSheet.Cells[reportRow, 1].PutValue(shape.Name);
                    reportSheet.Cells[reportRow, 2].PutValue(shape.MsoDrawingType.ToString());
                    reportSheet.Cells[reportRow, 3].PutValue(upperLeftAddress);
                    reportSheet.Cells[reportRow, 4].PutValue(lowerRightAddress);

                    reportRow++;
                }
            }

            // Save the report workbook
            reportWorkbook.Save("ShapeReport.xlsx");
        }
    }
}
