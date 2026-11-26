// Title: Create a TextBox shape with a horizontal two‑color gradient fill using custom ARGB colors in Aspose.Cells for .NET
// Common Searches: how to add a textbox shape with gradient fill in Aspose.Cells C# | Aspose.Cells set horizontal two color gradient on a shape | C# example of FillFormat.SetTwoColorGradient for Excel textbox | apply custom ARGB colors to a shape gradient in Aspose.Cells | save workbook with gradient-filled textbox using Aspose.Cells .NET
// Tags: aspocells textbox horizontal gradient fill | c# settwocolorgradient fillformat example | custom argb colors gradient aspocells | excel shape gradient fill aspocells .net | fillformat gradientstyle horizontal aspocells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace GradientTextboxDemo
{
    // The sample creates a new workbook, inserts a TextBox shape, assigns text, defines two custom ARGB colors, applies a horizontal two‑color gradient using FillFormat.SetTwoColorGradient, optionally sets the shape border weight, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "GradientTextboxDemo.xlsx";

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a TextBox shape to the worksheet
                // Parameters: drawing type, upper row, upper column, top offset, left offset, width, height
                Shape textBox = sheet.Shapes.AddShape(
                    MsoDrawingType.TextBox, // TextBox shape
                    2,                      // Upper row index
                    1,                      // Upper column index
                    0,                      // Top offset (pixels)
                    0,                      // Left offset (pixels)
                    300,                    // Width (pixels)
                    100);                   // Height (pixels)

                // Set the text inside the textbox
                textBox.Text = "Gradient Filled TextBox";

                // Configure the fill format for a two‑color gradient
                FillFormat fill = textBox.Fill;
                fill.FillType = FillType.Gradient;

                // Define custom colors
                Color colorStart = Color.FromArgb(255, 255, 128, 0); // Orange
                Color colorEnd   = Color.FromArgb(255, 0, 128, 255); // Light Blue

                // Apply the two‑color gradient (horizontal direction)
                fill.SetTwoColorGradient(colorStart, colorEnd, GradientStyleType.Horizontal, 1);

                // Optional: make the textbox border visible
                // Some older Aspose.Cells versions may not expose dash style or color properties,
                // so we set only the weight which is universally supported.
                textBox.Line.Weight = 1.0;

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the gradient textbox:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
