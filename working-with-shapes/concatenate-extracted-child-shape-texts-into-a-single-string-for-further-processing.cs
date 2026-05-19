using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ConcatenateShapeTexts
    {
        public static void Main(string[] args)
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
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Accumulate text from all shapes
            StringBuilder allTexts = new StringBuilder();

            // Iterate through worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape is GroupShape groupShape)
                    {
                        // Process child shapes of a group
                        foreach (Shape child in groupShape.GetGroupedShapes())
                        {
                            string childText = child.Text ?? child.TextBody?.Text;
                            if (!string.IsNullOrEmpty(childText))
                            {
                                allTexts.Append(childText).Append(' ');
                            }
                        }
                    }
                    else
                    {
                        // Process a regular shape
                        string shapeText = shape.Text ?? shape.TextBody?.Text;
                        if (!string.IsNullOrEmpty(shapeText))
                        {
                            allTexts.Append(shapeText).Append(' ');
                        }
                    }
                }
            }

            // Final concatenated result
            string concatenatedText = allTexts.ToString().Trim();

            Console.WriteLine("Concatenated Shape Texts:");
            Console.WriteLine(concatenatedText);

            // Save the workbook (if any modifications were made)
            workbook.Save(outputPath);
        }
    }
}