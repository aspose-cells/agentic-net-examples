// Title: Find Shape‑Only Worksheets in Aspose.Cells for .NET (MaxDataRow = -1, Shapes.Count > 0)
// Description: This C# example creates a workbook with a sheet that contains only a rectangle shape and another sheet with regular cell data. It then scans all worksheets, using Cells.MaxDataRow to detect the absence of cell content (‑1) and Shapes.Count to verify the presence of drawings, and outputs the names of sheets that meet both criteria before saving the file.
// Keywords: Aspose.Cells shape only worksheet | MaxDataRow -1 detection | Shapes.Count Aspose | identify drawing‑only sheets | .NET Excel shape detection
// Common Searches: how to list worksheets that contain only shapes using Aspose.Cells | detect Excel sheets with no data but with drawings in C# | Aspose.Cells find sheets with MaxDataRow -1 and shapes | filter shape‑only worksheets in a workbook | C# code to check for worksheets that have only drawings
// Developer Intent: Locate worksheets that have drawing objects but no cell data in an Aspose.Cells workbook.
// Use Cases: Generate a report of all shape‑only sheets for auditing. | Skip processing of non‑data worksheets to improve performance. | Separate drawing‑only tabs before exporting data to other formats.
// AI Prompts: Write C# code with Aspose.Cells that returns the names of worksheets where Cells.MaxDataRow is -1 and Shapes.Count > 0. | Create a method to remove shape‑only worksheets from a workbook and save the remaining sheets to a new file. | Show how to log each worksheet that contains only shapes in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeOnlyWorksheetDemo
{
    // This C# example creates a workbook with a sheet that contains only a rectangle shape and another sheet with regular cell data. It then scans all worksheets, using Cells.MaxDataRow to detect the absence of cell content (‑1) and Shapes.Count to verify the presence of drawings, and outputs the names of sheets that meet both criteria before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Worksheet 0: contains only shapes, no cell data
                // -------------------------------------------------
                Worksheet shapeOnlySheet = workbook.Worksheets[0];
                shapeOnlySheet.Name = "ShapesOnly";

                // Add a rectangle shape to the worksheet
                // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
                shapeOnlySheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);

                // -------------------------------------------------
                // Worksheet 1: contains regular cell data
                // -------------------------------------------------
                int dataSheetIndex = workbook.Worksheets.Add();
                Worksheet dataSheet = workbook.Worksheets[dataSheetIndex];
                dataSheet.Name = "DataSheet";

                // Populate some cells with data
                dataSheet.Cells["A1"].PutValue("Header");
                dataSheet.Cells["A2"].PutValue(123);

                // -------------------------------------------------
                // Identify worksheets that contain ONLY shapes
                // Criteria:
                //   - No data rows (MaxDataRow == -1)
                //   - At least one shape (Shapes.Count > 0)
                // -------------------------------------------------
                Console.WriteLine("Worksheets containing only shapes:");
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // MaxDataRow returns -1 when there is no data in any cell of the sheet
                    bool hasNoData = sheet.Cells.MaxDataRow == -1;

                    // Shapes.Count gives the number of drawing objects on the sheet
                    bool hasShapes = sheet.Shapes.Count > 0;

                    if (hasNoData && hasShapes)
                    {
                        Console.WriteLine($"- {sheet.Name}");
                    }
                }

                // Save the workbook (lifecycle rule: save)
                workbook.Save("ShapeOnlyWorksheetsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
