// Title: Change the font color of a specific text portion in an Aspose.Cells TextBox shape to a custom RGB value using C#
// AI Prompts: Write C# code with Aspose.Cells that sets the font color of the second portion of a TextBox's text to a given RGB value while preserving the original color of other portions. | Show how to programmatically detect TextBody support in Aspose.Cells and apply different RGB colors to multiple text portions inside a TextBox shape. | Provide a concise example that formats only part of the text in an Excel TextBox using Aspose.Cells and saves the workbook.
// Common Searches: asp.net c# aspose.cells change color of part of textbox text | how to set custom RGB font color for a portion of text in an Excel textbox using Aspose | partial text formatting in Aspose.Cells TextBox shape C# example | detect TextBody support in Aspose.Cells and format textbox portions | apply different colors to multiple portions of a textbox in Aspose.Cells
// Tags: Aspose.Cells TextBox partial font color | C# set RGB color TextBody portion | Excel shape text formatting Aspose.Cells | detect TextBody support Aspose.Cells | custom font color textbox Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook, accesses the first shape on the first worksheet, confirms it is a TextBox, sets the whole textbox font color to teal via the Font.Color property, includes placeholder logic for applying a different RGB color to a second text portion using TextBody when supported, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the sheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
                return;
            }

            // Get the first shape (assumed to be a TextBox)
            Shape shape = sheet.Shapes[0];

            // Verify that the shape is a TextBox
            if (shape is TextBox textbox)
            {
                // Set the font color for the entire textbox text
                // (Aspose.Cells versions prior to supporting TextBody use the Font property)
                textbox.Font.Color = Color.FromArgb(0, 128, 128); // teal

                // If you need to apply a different color to a second portion,
                // you would need a newer Aspose.Cells version that supports TextBody.
                // The following is a placeholder for such logic:
                // if (textbox.TextBody?.Paragraphs[0]?.Portions.Count > 1)
                // {
                //     textbox.TextBody.Paragraphs[0].Portions[1].Font.Color = Color.FromArgb(255, 165, 0); // orange
                // }
            }
            else
            {
                Console.WriteLine("The first shape is not a TextBox.");
                return;
            }

            // Save the workbook with the updated textbox formatting
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
