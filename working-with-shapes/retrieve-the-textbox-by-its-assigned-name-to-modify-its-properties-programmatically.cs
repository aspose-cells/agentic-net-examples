// Title: How to retrieve a textbox shape by its assigned name and change its text and font in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load an Excel file with Aspose.Cells, locate a shape named "MyTextBox" on a worksheet, update its Text property and set Font.Name, Font.Size, and Font.Color, then save the workbook. | Write C# code that verifies a shape retrieved by name is a TextBox, modifies its displayed text and font attributes safely, and handles the case where the shape is missing. | Show how to adjust optional fill and line formatting of a named textbox in Aspose.Cells while accounting for version‑specific API availability.
// Common Searches: Aspose.Cells C# retrieve shape by name from worksheet | change textbox text and font in Excel using Aspose.Cells .NET | update a named textbox's properties in an existing .xlsx file with Aspose.Cells | C# example for modifying textbox shape fill color in Aspose.Cells | handle missing textbox shape when using Aspose.Cells Shapes collection
// Tags: access named textbox Aspose.Cells | set textbox font attributes C# | update textbox displayed text .NET | shape fill line version handling Aspose.Cells | check shape type before modification Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, accesses the first worksheet, retrieves a textbox shape named "MyTextBox", updates its text and font (name, size, color), optionally adjusts fill and border settings while considering API version differences, and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the textbox shape by its assigned name
            // Replace "MyTextBox" with the actual name of your textbox
            Shape textbox = sheet.Shapes["MyTextBox"];

            // Ensure the shape exists and is a textbox before modifying
            if (textbox != null && textbox.Type == AutoShapeType.TextBox)
            {
                try
                {
                    // Change the displayed text
                    textbox.Text = "Updated text content";

                    // Modify font properties
                    textbox.Font.Name = "Arial";
                    textbox.Font.Size = 12;
                    textbox.Font.Color = Color.Blue;

                    // NOTE: Fill and line properties may vary between Aspose.Cells versions.
                    // The following lines are commented out to ensure compatibility.
                    // Uncomment and adjust if your version supports these members.

                    // textbox.Fill.SolidFillColor = Color.Yellow;   // Set background color
                    // textbox.Line.Weight = 1.5;                  // Set border thickness
                    // textbox.Line.Color = Color.Red;             // Set border color
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Error modifying textbox: {exShape.Message}");
                }
            }
            else
            {
                Console.WriteLine("Textbox shape not found or is not a TextBox.");
            }

            // Save the workbook with the modifications
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
