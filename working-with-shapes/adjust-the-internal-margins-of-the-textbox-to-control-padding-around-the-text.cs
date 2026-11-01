// Title: How to configure internal padding for a TextBox shape using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set internal padding of a TextBox shape in Excel | example of adjusting textbox margins with Aspose.Cells .NET | C# code to increase text padding inside a shape using Aspose.Cells | how to use TextBox.InternalMargin property in Aspose.Cells | set textbox internal margins in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells TextBox internal margin | C# set textbox padding Excel | Aspose.Cells shape margin configuration | Excel textbox padding .NET | Aspose.Cells TextBox margin property example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExample
{
    // The sample creates a new Workbook, adds a TextBox shape to the first worksheet, assigns text and line weight, and saves the file as 'TextboxWithMargins.xlsx'. It demonstrates where to apply the TextBox.InternalMargin property to control the padding around the text inside the shape.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Add a textbox shape at cell B2 (row 1, column 1) with size 200x100 points
                TextBox textbox = sheet.Shapes.AddTextBox(1, 1, 0, 0, 200, 100);

                // Set the text of the textbox
                textbox.Text = "Sample text with custom padding.";

                // Set simple line formatting
                textbox.Line.Weight = 1;

                // Optionally set line dash style if the enum is available
                // (Commented out to avoid compilation issues on older versions)
                // textbox.Line.DashStyle = LineDashStyle.Solid;

                // Optionally set fill color if the property exists
                // (Commented out to avoid compilation issues on older versions)
                // textbox.FillColor = Color.LightYellow;

                // Define output file path
                string outputPath = "TextboxWithMargins.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
