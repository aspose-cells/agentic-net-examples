// Title: Reset rotation of TextBox shapes in an Excel worksheet using LINQ with Aspose.Cells for .NET (C#)
// Common Searches: how to reset rotation angle of TextBox shapes in an Excel file using Aspose.Cells C# | LINQ query to find rotated TextBox shapes in a workbook with Aspose.Cells | disable shape rotation programmatically in .NET Excel using Aspose.Cells | C# code to set RotationAngle to zero for all TextBox objects in a worksheet | Aspose.Cells example for clearing text box rotation in Excel workbook
// Tags: reset textbox rotation Aspose.Cells | LINQ shape filtering rotation Aspose.Cells | modify shape rotation Excel .NET | Aspose.Cells shape manipulation C# | text box rotation angle zero

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Linq;

// The example loads an existing Excel workbook, uses LINQ to locate every TextBox shape on the first worksheet whose RotationAngle is not zero, sets each shape's RotationAngle to zero, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Find all TextBox shapes whose rotation angle is not zero
            var textBoxesToReset = sheet.Shapes
                .Cast<Shape>()
                .OfType<TextBox>()
                .Where(tb => tb.RotationAngle != 0)
                .ToList();

            // Reset rotation for each identified TextBox
            foreach (var tb in textBoxesToReset)
            {
                try
                {
                    tb.RotationAngle = 0;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to reset rotation for a shape: {ex.Message}");
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
