// Title: How to add exception handling for a missing named TextBox shape in Aspose.Cells C#
// AI Prompts: Generate C# code using Aspose.Cells that attempts to retrieve a TextBox shape by its name and logs a custom error if the shape is not found. | Modify the example to differentiate between a missing TextBox shape and other runtime errors, returning a specific message for each case. | Show how to safely update the text of a named TextBox in an Excel worksheet with Aspose.Cells while wrapping the shape access in a try‑catch block.
// Common Searches: Aspose.Cells C# check for existence of a TextBox shape before updating its text | C# catch exception when worksheet.Shapes["MyTextBox"] throws index out of range | How to handle missing shape name error in Aspose.Cells workbook | Safely update a named TextBox in Excel using Aspose.Cells with error handling | Aspose.Cells exception handling for non‑existent shapes in C#
// Tags: Aspose.Cells exception handling for missing shape | C# retrieve TextBox shape by name | worksheet.Shapes index out of range protection | update TextBox text safely Aspose.Cells | validate shape existence Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, attempts to locate a TextBox named 'MyTextBox' on the first worksheet, updates its text, and uses try‑catch blocks to handle missing input files, absent TextBox shapes, and save errors, providing clear error messages for each scenario.
class TextBoxHandler
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string textBoxName = "MyTextBox";

        Workbook workbook = null;

        // Load the workbook safely
        try
        {
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve and modify the TextBox shape
        try
        {
            Shape textBox = worksheet.Shapes[textBoxName];
            textBox.Text = "Updated text content";
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: TextBox named '{textBoxName}' was not found in the worksheet.");
            Console.WriteLine($"Exception details: {ex.Message}");
        }

        // Save the workbook safely
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
