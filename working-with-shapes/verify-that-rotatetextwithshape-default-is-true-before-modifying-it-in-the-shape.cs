using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class VerifyAndModifyRotateTextWithShape
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape
                Shape shape = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 100, 200);
                shape.Text = "Sample Text";

                // Verify the default value of RotateTextWithShape
                bool defaultValue = shape.TextBody.TextAlignment.RotateTextWithShape;
                Console.WriteLine("Default RotateTextWithShape value: " + defaultValue);

                // If the default is true, modify it (set to false as an example)
                if (defaultValue)
                {
                    shape.TextBody.TextAlignment.RotateTextWithShape = false;
                    Console.WriteLine("RotateTextWithShape modified to: " + shape.TextBody.TextAlignment.RotateTextWithShape);
                }

                // Save the workbook
                string outputPath = "VerifyRotateTextWithShape.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            VerifyAndModifyRotateTextWithShape.Run();
        }
    }
}