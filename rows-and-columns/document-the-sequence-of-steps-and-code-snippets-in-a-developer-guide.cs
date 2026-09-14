// Title: How to add adjustment guides to a RightArrowCallout AutoShape and draw a rectangular freeform shape with ShapePath using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, inserts a RightArrowCallout auto shape, defines four adjustment guides (adj1‑adj4), updates each guide’s value, and saves the workbook. | Show C# steps to construct a rectangular freeform shape using ShapePath methods (MoveTo, LineTo, Close) and add it to a worksheet with AddFreeform. | Provide a C# snippet that iterates over a ShapeGuideCollection to print each guide’s value before and after modification.
// Common Searches: Aspose.Cells .NET example for setting guide values on a callout shape | How to draw a rectangular freeform shape with ShapePath in Aspose.Cells | Changing auto shape adjustment parameters programmatically using Aspose.Cells C# | Saving an Excel file after adding custom shapes with Aspose.Cells
// Tags: add adjustment guides AutoShape Aspose.Cells | modify shape guide values C# | create rectangular freeform shape ShapePath Aspose.Cells | insert freeform shape worksheet Aspose.Cells | save workbook after shape modifications Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDeveloperGuide
{
    // Demonstrates creating a workbook, inserting a RightArrowCallout auto shape, adding and updating its adjustment guides, building a rectangular freeform shape with ShapePath, adding it to the worksheet, and saving the file.
    public class ShapeGuideAndPathDemo
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();                     // create workbook
                Worksheet worksheet = workbook.Worksheets[0];           // get first worksheet

                // 2. Add an AutoShape that supports adjustment guides (e.g., RightArrowCallout)
                Shape shape = worksheet.Shapes.AddAutoShape(
                    AutoShapeType.RightArrowCallout, // shape type
                    2, 0, 2, 0,                      // upper-left row, column, offsetX, offsetY
                    200, 150);                       // width, height

                // 3. Access the ShapeGuideCollection via the shape's geometry
                ShapeGuideCollection guides = shape.Geometry.ShapeAdjustValues;

                // 4. Add adjustment guides using the Add(string name, double value) method
                guides.Add("adj1", 25.5);
                guides.Add("adj2", 30.0);
                guides.Add("adj3", 25.5);
                guides.Add("adj4", 35.0);

                // 5. Read and display the values of the added guides
                Console.WriteLine("Initial guide values:");
                for (int i = 0; i < guides.Count; i++)
                {
                    Console.WriteLine($"Guide {i + 1}: {guides[i].Value}");
                }

                // 6. Modify guide values via the Value property
                guides[0].Value = 20.0;
                guides[1].Value = 20.0;
                guides[2].Value = 20.0;
                guides[3].Value = 20.0;

                // 7. Verify the updated values
                Console.WriteLine("Updated guide values:");
                for (int i = 0; i < guides.Count; i++)
                {
                    Console.WriteLine($"Guide {i + 1}: {guides[i].Value}");
                }

                // 8. Create a custom ShapePath to define a freeform shape
                ShapePath customPath = new ShapePath();                 // instantiate ShapePath
                customPath.MoveTo(100, 100);                            // start point
                customPath.LineTo(200, 100);                            // line to (200,100)
                customPath.LineTo(200, 200);                            // line to (200,200)
                customPath.LineTo(100, 200);                            // line to (100,200)
                customPath.Close();                                     // close the path

                // 9. Add the freeform shape to the worksheet using the custom path
                worksheet.Shapes.AddFreeform(
                    1,          // upper-left row
                    0,          // upper-left column
                    1,          // top offset
                    0,          // left offset
                    200,        // width
                    200,        // height
                    new ShapePath[] { customPath } // array of ShapePath objects
                );

                // 10. Save the workbook to a file
                string outputPath = "ShapeGuideAndPathDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeGuideAndPathDemo.Run();
        }
    }
}
