// Title: Move multiple rectangle shapes together in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# move several shapes as a group in Excel | how to offset multiple rectangle shapes simultaneously using Aspose.Cells | C# code to shift rectangle shapes together in an .xlsx file with Aspose.Cells | grouping rectangle shapes for collective movement in Aspose.Cells .NET | programmatically move multiple shapes in the same direction with Aspose.Cells
// Tags: rectangle shape grouping Aspose.Cells | simultaneous shape repositioning C# | shape left/top offset Aspose.Cells | add rectangle to worksheet programmatically | collective shape movement Excel .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace GroupRectangleShapesDemo
{
    // The example creates a workbook, adds two rectangle shapes to the first worksheet, applies identical Left and Top offsets to both shapes so they move together, and saves the file as GroupedRectangles.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "GroupedRectangles.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Add two independent rectangle shapes
                //    Parameters: upperLeftRow, upperLeftColumn, top, left, width, height
                Shape rect1 = sheet.Shapes.AddRectangle(2, 2, 0, 0, 120, 60);
                rect1.Name = "Rect1";
                rect1.AlternativeText = "First rectangle";

                Shape rect2 = sheet.Shapes.AddRectangle(5, 5, 0, 0, 120, 60);
                rect2.Name = "Rect2";
                rect2.AlternativeText = "Second rectangle";

                // 3. Move both rectangles together to simulate grouping
                //    Adjust the position as needed (values are in pixels)
                const int offsetLeft = 50; // Move 50 pixels to the right
                const int offsetTop = 30;  // Move 30 pixels down

                rect1.Left += offsetLeft;
                rect1.Top  += offsetTop;

                rect2.Left += offsetLeft;
                rect2.Top  += offsetTop;

                // 4. Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
                Console.WriteLine("Both rectangles have been moved together.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing shapes:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
