// Title: Add a custom hover tooltip to an Excel shape named "InfoBox" using Aspose.Cells for .NET (C#)
// AI Prompts: Assign a custom string to the AlternativeText attribute of a shape called "InfoBox" after loading the workbook with Aspose.Cells. | Write C# code that verifies an input Excel file, locates a shape by its name, sets hover help text, and saves the workbook using Aspose.Cells. | Generate a snippet that demonstrates adding a tooltip to an Excel shape via the Shape.AlternativeText property in Aspose.Cells for .NET.
// Common Searches: aspnet add tooltip to specific shape in Excel using Aspose.Cells C# | set alternative text for shape named InfoBox in Aspose.Cells workbook | C# Aspose.Cells find shape by name and assign hover text | how to display custom help text when hovering over an Excel shape with Aspose.Cells | Aspose.Cells shape tooltip example C#
// Tags: Aspose.Cells shape AlternativeText | C# add tooltip to Excel shape | Aspose.Cells locate shape by name | Excel shape hover help text .NET | modify shape tooltip Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing workbook, searches the first worksheet for a shape named "InfoBox", sets its AlternativeText property to a custom message that appears as a hover tooltip, and saves the updated file, including file‑existence checks and error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the shape.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0) where the shape resides.
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the shape named "InfoBox".
            Shape infoBox = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape.Name == "InfoBox")
                {
                    infoBox = shape;
                    break;
                }
            }

            if (infoBox != null)
            {
                // Assign custom tooltip text that appears on hover.
                infoBox.AlternativeText = "Enter your custom help text here.";
            }
            else
            {
                Console.WriteLine("Shape 'InfoBox' not found.");
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
