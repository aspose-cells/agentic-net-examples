using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    public class ChangeLatinFontDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape that will contain text
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);
                shape.Text = "Western text with Calibri";

                // Access the TextOptions of the first paragraph in the shape's TextBody
                TextOptions textOptions = shape.TextBody.TextParagraphs[0].TextOptions;

                // Set the Latin (western) font name to Calibri
                textOptions.LatinName = "Calibri";

                // Save the workbook to a file
                string outputPath = "ChangeLatinFontDemo.xlsx";

                // Ensure we can write to the target location
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                workbook.Save(outputPath);

                // Confirmation
                Console.WriteLine($"Latin font set to Calibri. Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ChangeLatinFontDemo.Run();
        }
    }
}