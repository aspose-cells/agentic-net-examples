// Title: Unlock a locked TextBox shape in an Excel .xlsx file with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to iterate over Worksheet.Shapes, detect TextBox objects, and set their IsLocked property to false in C#. | Load an existing .xlsx workbook, unlock all TextBox shapes to enable moving and resizing, then save the file using Aspose.Cells. | Programmatically change the lock state of Excel TextBox shapes by accessing the Shape.IsLocked property via the Aspose.Cells API.
// Common Searches: C# Aspose.Cells how to unlock TextBox shape in Excel workbook | set IsLocked false for TextBox objects using Aspose.Cells .NET | modify locked textbox position and size programmatically in .xlsx with Aspose | unlock Excel shape lock property Aspose.Cells example | iterate worksheet shapes to find TextBox and change lock state in C#
// Tags: Aspose.Cells unlock TextBox shape C# | Shape.IsLocked property usage | Iterate worksheet shapes Aspose.Cells | Edit Excel textbox position programmatically | Load and save .xlsx with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing .xlsx workbook, loops through all shapes on the first worksheet, unlocks any TextBox shapes by setting Shape.IsLocked = false, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes and unlock any TextBox shapes
            foreach (Shape shape in worksheet.Shapes)
            {
                // Check if the shape is a TextBox
                if (shape is TextBox)
                {
                    // Unlock the shape to allow editing of position and size
                    shape.IsLocked = false;
                }
            }

            // Save the workbook with the unlocked textbox
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
