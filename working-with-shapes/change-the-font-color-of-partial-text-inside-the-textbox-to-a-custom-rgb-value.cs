// Title: Apply a custom RGB font color to a TextBox shape in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets the font color of a TextBox shape to RGB(128,0,255) with Aspose.Cells. | Explain how to change the entire text color of an Aspose.Cells TextBox and why partial text coloring is not supported. | Provide a workaround to simulate partial text coloring in an Aspose.Cells TextBox using multiple shapes or RichText.
// Common Searches: aspnet cells change textbox font color rgb c# | how to set custom text color in Aspose.Cells TextBox shape | Aspose.Cells textbox entire text color change example | partial text formatting limitation Aspose.Cells textbox | C# Aspose.Cells set shape text color to specific RGB value
// Tags: textbox font color Aspose.Cells C# | custom RGB text color Aspose.Cells | Aspose.Cells shape text formatting limitation | C# set textbox text color Aspose.Cells | Aspose.Cells partial text styling workaround

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a TextBox shape to the first worksheet, sets its text to "Hello World", and changes the font color of the entire text to the custom RGB value (128,0,255) using Aspose.Cells for .NET, then saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset (pixels),
            // lower right row, lower right column, lower right offset (pixels)
            TextBox textbox = worksheet.Shapes.AddTextBox(2, 1, 0, 4, 2, 0);

            // Set the textbox text
            textbox.Text = "Hello World";

            // Change the font color of the entire text (partial formatting not supported in this version)
            textbox.Font.Color = Color.FromArgb(128, 0, 255);

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
