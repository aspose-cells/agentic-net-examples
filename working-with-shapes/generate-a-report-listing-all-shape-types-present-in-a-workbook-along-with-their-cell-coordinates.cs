using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeReportGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source workbook that contains shapes
            Workbook sourceWorkbook = new Workbook("input.xlsx");

            // Create a new workbook to hold the report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "Shape Report";

            // Write header row
            reportSheet.Cells["A1"].PutValue("Worksheet");
            reportSheet.Cells["B1"].PutValue("Shape Name");
            reportSheet.Cells["C1"].PutValue("Shape Type");
            reportSheet.Cells["D1"].PutValue("Upper Left Row");
            reportSheet.Cells["E1"].PutValue("Upper Left Column");
            reportSheet.Cells["F1"].PutValue("Lower Right Row");
            reportSheet.Cells["G1"].PutValue("Lower Right Column");

            int reportRow = 1; // zero‑based index; start after header

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                // Access the collection of shapes in the current worksheet
                ShapeCollection shapes = ws.Shapes;

                // Loop through all shapes
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Write shape information to the report sheet
                    reportSheet.Cells[reportRow, 0].PutValue(ws.Name);                     // Worksheet name
                    reportSheet.Cells[reportRow, 1].PutValue(shape.Name);                 // Shape name
                    reportSheet.Cells[reportRow, 2].PutValue(shape.MsoDrawingType.ToString()); // Shape type
                    reportSheet.Cells[reportRow, 3].PutValue(shape.UpperLeftRow);        // Upper left row
                    reportSheet.Cells[reportRow, 4].PutValue(shape.UpperLeftColumn);     // Upper left column
                    reportSheet.Cells[reportRow, 5].PutValue(shape.LowerRightRow);       // Lower right row
                    reportSheet.Cells[reportRow, 6].PutValue(shape.LowerRightColumn);    // Lower right column

                    reportRow++;
                }
            }

            // Save the report workbook
            reportWorkbook.Save("ShapeReport.xlsx");

            // Optional: also output to console
            Console.WriteLine("Shape report generated successfully.");
        }
    }
}