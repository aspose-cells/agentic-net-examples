using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    public class PreserveSpaceBeforeDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text box shape
            Shape shape = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 300, 200);
            // Set two paragraphs separated by a line break
            shape.Text = "First paragraph\nSecond paragraph";

            // Access the second paragraph (index 1)
            TextParagraph paragraph = shape.TextBody.TextParagraphs[1];

            // Ensure the paragraph uses point units for spacing
            paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
            paragraph.SpaceBeforeSizeType = LineSpaceSizeType.Points;
            paragraph.SpaceAfterSizeType = LineSpaceSizeType.Points;

            // Read the current SpaceBefore value (in points) and store it
            double originalSpaceBefore = paragraph.SpaceBefore;
            Console.WriteLine("Original SpaceBefore: " + originalSpaceBefore);

            // Modify the SpaceBefore value
            paragraph.SpaceBefore = originalSpaceBefore + 5; // increase by 5 points
            Console.WriteLine("Modified SpaceBefore: " + paragraph.SpaceBefore);

            // Save the workbook to verify persistence
            string outputPath = "PreserveSpaceBeforeDemo.xlsx";
            workbook.Save(outputPath);

            // Reload the workbook to confirm the saved value
            if (File.Exists(outputPath))
            {
                Workbook loadedWorkbook = new Workbook(outputPath);
                Shape loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];
                TextParagraph loadedParagraph = loadedShape.TextBody.TextParagraphs[1];
                Console.WriteLine("Reloaded SpaceBefore: " + loadedParagraph.SpaceBefore);
            }
            else
            {
                Console.WriteLine("Failed to find the saved file: " + outputPath);
            }
        }
    }
}