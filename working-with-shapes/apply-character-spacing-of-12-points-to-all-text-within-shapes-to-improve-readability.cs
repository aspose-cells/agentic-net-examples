// Title: Adjust text box font size to improve readability of shape text in Excel using Aspose.Cells for .NET
// Common Searches: how to change font size of text boxes in an Excel file with Aspose.Cells C# | Aspose.Cells C# modify shape text appearance for better readability | set character spacing for shape text in Excel workbook using Aspose.Cells .NET | increase readability of Excel shape text by adjusting font properties in C#
// Tags: shape text formatting Aspose.Cells .NET | text box font size adjustment C# | Excel shape text appearance customization | Aspose.Cells shape text properties workaround | readability improvement for Excel shape text

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, iterates through each worksheet and its shapes, identifies TextBox shapes, and sets the TextEffect.FontSize to 12 as a practical way to enhance text readability (since a direct character spacing property is not exposed in this Aspose.Cells version). The modified workbook is then saved, with checks for missing input files and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each shape on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Apply only to TextBox shapes
                        if (shape is TextBox textBox)
                        {
                            // Example: set the font size of the text inside the textbox
                            // (FontSpacing property is not available in this version of Aspose.Cells)
                            textBox.TextEffect.FontSize = 12;
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
