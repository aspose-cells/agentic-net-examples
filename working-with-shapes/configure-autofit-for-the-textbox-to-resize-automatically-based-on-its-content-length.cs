// Title: How to enable auto‑fit for a TextBox shape in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# make textbox shape resize to fit its text | C# Aspose.Cells auto‑size textbox based on content length | set textbox shape to auto‑fit its text in Aspose.Cells .NET | Aspose.Cells textbox auto‑fit not working in C# example
// Tags: Aspose.Cells TextBox auto‑fit limitation | C# Aspose.Cells shape resize to content | Excel workbook textbox auto‑size with Aspose.Cells | Aspose.Cells shape dimensions based on text length

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, adds a TextBox shape, assigns a long string, and demonstrates that the AutoSize property is unavailable in the current Aspose.Cells version, requiring manual size adjustment before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet
                // Parameters: upper left row, upper left column, offsetX, offsetY, width, height
                TextBox textBox = sheet.Shapes.AddTextBox(2, 1, 0, 0, 100, 30);

                // Set the textbox content
                textBox.Text = "This is a sample text that is long enough to demonstrate auto‑fit functionality.";

                // AutoSize property is not available in this version of Aspose.Cells.
                // The textbox size is kept as defined; adjust manually if needed.

                // Define output file path
                string outputPath = "AutoFitTextbox.xlsx";

                // Ensure the output directory exists
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
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
