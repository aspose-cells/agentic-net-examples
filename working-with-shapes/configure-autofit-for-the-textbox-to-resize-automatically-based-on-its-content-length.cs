// Title: How to enable auto‑fit for a textbox shape so it resizes automatically to its content in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set TextBox AutoSize to true | Resize Excel textbox automatically based on text length using Aspose.Cells | Enable auto‑fit for a shape textbox in a .NET workbook | Retrieve textbox dimensions after auto‑fit with Aspose.Cells | Add textbox to worksheet and auto‑size it with Aspose.Cells API
// Tags: Aspose.Cells TextBox AutoSize property | auto‑fit textbox shape in Excel using C# | Aspose.Cells TextBody.TextAlignment.AutoSize usage | dynamic textbox resizing Aspose.Cells .NET | add and auto‑size textbox shape Aspose.Cells | Excel worksheet textbox auto‑fit with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AutoFitTextboxDemo
{
    // Creates a workbook, adds a textbox shape, assigns long text, enables automatic resizing by setting TextBody.TextAlignment.AutoSize to true, prints the resulting dimensions, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "AutoFitTextbox.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "Demo";

                // 2. Add a textbox shape to the worksheet
                // Overload: AddTextBox(int topRow, int top, int leftColumn, int left, int height, int width)
                TextBox textBox = worksheet.Shapes.AddTextBox(
                    topRow: 2,      // upper‑left row index
                    top: 0,         // vertical offset in pixels
                    leftColumn: 2,  // upper‑left column index
                    left: 0,        // horizontal offset in pixels
                    height: 100,    // height in pixels
                    width: 200);    // width in pixels

                // 3. Set the text content of the textbox
                string longText = "This is a very long piece of text that should cause the textbox to automatically resize itself based on the length of its content. " +
                                  "Aspose.Cells provides an AutoFit property to achieve this behavior.";
                textBox.Text = longText;

                // 4. Enable auto‑fit so the textbox resizes to fit its content
                // The TextBox class exposes AutoSize through its TextBody.TextAlignment object
                textBox.TextBody.TextAlignment.AutoSize = true;

                // 5. (Optional) Display the size after creation
                Console.WriteLine($"Textbox size - Height: {textBox.Height} pixels, Width: {textBox.Width} pixels");

                // 6. Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // 7. Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while configuring the textbox:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
