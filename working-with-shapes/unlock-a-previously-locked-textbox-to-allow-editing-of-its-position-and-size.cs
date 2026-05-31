using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class UnlockTextBoxExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Find the first TextBox shape (if any)
            Shape textBoxShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                // In Aspose.Cells a TextBox is represented by the TextBox class derived from Shape
                if (shape is TextBox)
                {
                    textBoxShape = shape;
                    break;
                }
            }

            // Unlock the TextBox if it was found
            if (textBoxShape != null)
            {
                textBoxShape.IsLocked = false;
            }
            else
            {
                Console.WriteLine("No TextBox shape found in the worksheet.");
            }

            // Allow editing of drawing objects when the sheet is protected
            sheet.Protection.AllowEditingObject = true;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}