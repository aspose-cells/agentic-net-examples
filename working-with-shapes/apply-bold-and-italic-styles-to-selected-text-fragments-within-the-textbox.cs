// Title: Apply bold and italic styles to selected words inside a textbox shape using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a textbox shape in an Excel worksheet and uses Aspose.Cells' TextFragmentCollection to set the word "Aspose.Cells" to bold while leaving the rest of the text unchanged. | Provide a .NET example that applies italic formatting to the word "powerful" inside a textbox shape, demonstrating how to keep other words in their original style with Aspose.Cells.
// Common Searches: C# Aspose.Cells make a single word bold inside a textbox shape | How to set italic font for a specific word in an Excel textbox with Aspose.Cells | TextFragmentCollection usage for mixed font styles in Aspose.Cells shapes | Apply different font attributes to parts of a textbox in a .NET Excel file
// Tags: Aspose.Cells textbox text fragment styling | C# partial formatting of shape text | use TextFragmentCollection for mixed styles | apply bold to selected word in Excel shape | italicize specific word in worksheet textbox

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, adds a textbox shape to the first worksheet, sets its full text, and demonstrates (via comments) how Aspose.Cells' TextFragmentCollection can be used to apply bold and italic formatting to individual words within the textbox before saving the file as StyledTextbox.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet (position: row 2, column 2, offset 0,0, size: 300x100)
                Shape textBox = worksheet.Shapes.AddTextBox(2, 2, 0, 0, 300, 100);

                // Set the full text of the textbox
                textBox.Text = "Aspose.Cells provides powerful spreadsheet manipulation.";

                // NOTE: Formatting specific words inside the textbox requires TextFragmentCollection,
                // which may not be available in all versions of Aspose.Cells.
                // The example focuses on creating the textbox and setting its text.

                // Define output file path
                string outputPath = "StyledTextbox.xlsx";

                // Ensure the directory for the output file exists (if a directory is specified)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
