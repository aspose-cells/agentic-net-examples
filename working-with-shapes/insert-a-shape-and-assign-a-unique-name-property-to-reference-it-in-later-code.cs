// Title: Insert a rectangle auto shape into an Excel worksheet and assign a unique name for later lookup using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add rectangle auto shape with custom name to worksheet | retrieve shape by name in Aspose.Cells workbook using LINQ | change fill color and text of a named shape in Aspose.Cells .NET | set unique Name property on an auto shape for later modification with Aspose.Cells | C# example of inserting and updating a named shape in Excel using Aspose.Cells
// Tags: add rectangle auto shape Aspose.Cells | assign shape name Aspose.Cells .NET | lookup named shape Aspose.Cells | modify shape fill color Aspose.Cells | update shape text Aspose.Cells workbook

using System;
using System.Drawing;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeInsertDemo
{
    // The example creates a new workbook, inserts a rectangle auto shape on the first worksheet, assigns it a unique Name property, later retrieves the shape by that name using LINQ, updates its fill color and text, and saves the workbook as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Insert an auto shape (Rectangle) into the worksheet
                // Parameters: type, upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
                Shape rectShape = sheet.Shapes.AddAutoShape(
                    AutoShapeType.Rectangle,
                    2,          // upper left row (zero‑based)
                    1,          // upper left column (zero‑based)
                    0,          // row offset in pixels
                    0,          // column offset in pixels
                    150,        // width in pixels
                    80);        // height in pixels

                // Assign a unique name to the shape so it can be referenced later
                string uniqueShapeName = "MyUniqueRectangle";
                rectShape.Name = uniqueShapeName;

                // Optionally set some initial formatting
                rectShape.Fill.SolidFill.Color = Color.LightBlue;
                rectShape.Line.Weight = 2;
                rectShape.Line.DashStyle = MsoLineDashStyle.Solid;

                // -----------------------------------------------------------------
                // Later in the code we retrieve the shape by its unique name
                // -----------------------------------------------------------------
                Shape retrievedShape = sheet.Shapes.FirstOrDefault(s => s.Name == uniqueShapeName);
                if (retrievedShape == null)
                {
                    Console.WriteLine($"Shape with name '{uniqueShapeName}' was not found.");
                }
                else
                {
                    // Modify the retrieved shape – change its fill color and add some text
                    retrievedShape.Fill.SolidFill.Color = Color.Yellow;
                    retrievedShape.Text = "Updated Shape";
                    retrievedShape.Font.Color = Color.Black;
                    retrievedShape.Font.Size = 12;
                }

                // Save the workbook to verify the shape insertion and modifications
                string outputPath = "ShapeInsertDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Report any unexpected errors
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
