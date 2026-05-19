using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExamples
{
    public class ResetShapeTextFormatting
    {
        public static void Main()
        {
            try
            {
                Run();
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

            // Add a rectangle shape
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 0);
            shape.Text = "Formatted Text";

            // Apply custom formatting to the shape's text
            shape.Font.IsBold = true;
            shape.Font.Size = 16;
            shape.Font.Color = Color.Red;

            // Reset formatting to default values
            shape.Font.IsBold = false;
            shape.Font.Size = 11;          // default font size
            shape.Font.Color = Color.Black;
            shape.Font.Name = "Calibri";   // default font name

            // Save the workbook
            string outputPath = "ResetShapeTextFormatting.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}