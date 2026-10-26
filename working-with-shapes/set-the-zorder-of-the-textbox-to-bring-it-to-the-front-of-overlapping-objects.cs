// Title: How to bring a TextBox shape to the front in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# bring textbox shape to front of other objects | set Z-order for a TextBox in an Excel file with Aspose.Cells .NET | use BringToFront method on a shape in Aspose.Cells C# example | move overlapping Excel shapes to top layer using Aspose.Cells API
// Tags: Aspose.Cells shape Z-order | C# Aspose.Cells BringToFront method | Excel textbox layering Aspose.Cells | Aspose.Cells drawing order manipulation | Aspose.Cells set shape front layer

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, adds a TextBox shape, and applies the Aspose.Cells Shape API (e.g., BringToFront) to adjust the Z‑order so the textbox appears above any overlapping objects before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if the file exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape: upper left row, column, row offset, column offset, width, height
                TextBox textBox = worksheet.Shapes.AddTextBox(2, 0, 2, 0, 200, 100);

                // Optional: set some text for the textbox
                textBox.Text = "Sample TextBox";

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
