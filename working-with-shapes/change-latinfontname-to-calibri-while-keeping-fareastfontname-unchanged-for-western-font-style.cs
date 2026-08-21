// Title: Set a Shape's Latin Font to Calibri in Aspose.Cells for .NET (Keep FarEast Font Intact)
// Description: This example creates a workbook, adds a rectangle shape with text, reads the current Latin and FarEast font names, changes the LatinName to "Calibri" while leaving FarEastName unchanged, and saves the file as ChangeLatinFontDemo.xlsx.
// Keywords: Aspose.Cells shape font | change Latin font Calibri | preserve FarEast font | TextOptions LatinName .NET | shape text formatting Aspose.Cells | C# Aspose.Cells example | multilingual shape text
// Common Searches: Aspose.Cells set shape Latin font to Calibri | how to keep FarEast font when updating shape text font | TextOptions LatinName property usage | change western font of shape text Aspose.Cells | C# example for multilingual shape fonts
// Developer Intent: Apply Calibri to the western (Latin) portion of a shape's text without modifying the existing FarEast font.
// Use Cases: Standardize western text in all shapes to Calibri while retaining Asian fonts for multilingual reports. | Allow users to select a preferred Latin font for newly created shapes in a spreadsheet editor. | Generate invoices where English sections use Calibri and Japanese sections keep their original typeface.
// AI Prompts: Generate C# code that sets TextOptions.LatinName to "Calibri" for a given Shape in Aspose.Cells, preserving TextOptions.FarEastName. | Explain the difference between LatinName and FarEastName in Aspose.Cells shape text rendering. | Create a helper method that receives a Shape and a font name, updates only the Latin font, and returns the modified Shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a rectangle shape with text, reads the current Latin and FarEast font names, changes the LatinName to "Calibri" while leaving FarEastName unchanged, and saves the file as ChangeLatinFontDemo.xlsx.
    public class ChangeLatinFontDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape that will contain text
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);
                shape.Text = "Sample western text";

                // Access the TextOptions of the first paragraph of the shape's text body
                TextOptions textOptions = shape.TextBody.TextParagraphs[0].TextOptions;

                // Display the current Latin and FarEast font names (for demonstration)
                Console.WriteLine("Before change:");
                Console.WriteLine("LatinName: " + textOptions.LatinName);
                Console.WriteLine("FarEastName: " + textOptions.FarEastName);

                // Change the Latin (western) font name to Calibri
                textOptions.LatinName = "Calibri";

                // FarEastName is left unchanged automatically; we can verify it remains the same
                Console.WriteLine("\nAfter change:");
                Console.WriteLine("LatinName: " + textOptions.LatinName);
                Console.WriteLine("FarEastName (unchanged): " + textOptions.FarEastName);

                // Save the workbook to see the effect
                workbook.Save("ChangeLatinFontDemo.xlsx");
                Console.WriteLine("\nWorkbook saved as ChangeLatinFontDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
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
