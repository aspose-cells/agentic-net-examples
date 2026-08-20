// Title: Set TextBox Shape Background to Solid Color with Transparency in Aspose.Cells for .NET (C#)
// Description: Shows how to insert a TextBox shape into an Excel worksheet with Aspose.Cells for .NET, apply a solid ARGB fill, adjust its transparency (0‑1), and save the file.
// Keywords: Aspose.Cells | C# textbox fill | solid fill transparency | Excel shape background color | ARGB fill Aspose.Cells | shape fill type solid | Aspose.Cells .NET example | transparent shape fill | Excel textbox background
// Common Searches: Aspose.Cells set textbox background color | C# Aspose.Cells solid fill transparency | How to change textbox shape fill in Excel using Aspose.Cells | Set ARGB color for textbox shape Aspose.Cells | Apply transparency to shape fill Aspose.Cells .NET
// Developer Intent: Add a TextBox shape and apply a solid fill with optional transparency in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Create a report header with a semi‑transparent textbox that highlights the title while showing underlying data. | Overlay a colored, translucent textbox on a chart area to emphasize key metrics without obscuring the chart. | Design a dashboard widget where the textbox background matches a brand color and allows cell content to remain visible.
// AI Prompts: Generate C# code that sets a textbox shape's fill to a solid ARGB color with 50% transparency using Aspose.Cells for .NET. | Explain how the Transparency property and ARGB values work together to control opacity of a shape fill in Aspose.Cells. | Show how to modify an existing textbox shape in a workbook to change its background color and opacity programmatically.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to insert a TextBox shape into an Excel worksheet with Aspose.Cells for .NET, apply a solid ARGB fill, adjust its transparency (0‑1), and save the file.
    public class TextBoxBackgroundFillDemo
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, top, left, width, height
            Shape textBox = sheet.Shapes.AddShape(MsoDrawingType.TextBox, 2, 0, 2, 0, 200, 100);

            // Set the fill type to solid
            textBox.Fill.FillType = FillType.Solid;

            // Configure solid fill color and transparency
            SolidFill solidFill = textBox.Fill.SolidFill;
            solidFill.Color = Color.FromArgb(255, 100, 150, 200); // solid background color
            solidFill.Transparency = 0.3; // 30% transparent (0.0 = opaque, 1.0 = fully transparent)

            // Set sample text and font color
            textBox.Text = "Sample TextBox";
            textBox.Font.Color = Color.White;

            // Determine output path and save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TextBoxBackgroundFillDemo.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
