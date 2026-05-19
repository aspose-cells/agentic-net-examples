using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    public class SetLeftMarginPtDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape that will contain text
            // Parameters: upper left row, upper left column, top (pixels), left (pixels), width (pixels), height (pixels)
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 100);

            // Set the text inside the shape
            shape.Text = "Text aligned with cell borders";

            // Access the text alignment object and set the left margin to 8 points
            ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;
            textAlignment.LeftMarginPt = 8.0; // 8 points

            // Save the workbook to a file
            string outputPath = "SetLeftMarginPtDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}