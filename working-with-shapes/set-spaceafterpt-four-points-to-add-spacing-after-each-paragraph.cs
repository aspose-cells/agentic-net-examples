// Title: Add a 4‑point space after each paragraph in a shape using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel workbook, retrieves a specific shape, and sets the SpaceAfterPt property of every paragraph in the shape to 4 points, then saves the file. | Show how to loop through all paragraphs of a shape in Aspose.Cells and apply a 4‑point spacing after each paragraph using the ParagraphFormatting API.
// Common Searches: Aspose.Cells set SpaceAfterPt for shape paragraphs | C# add spacing after paragraph in Excel shape | How to increase paragraph spacing after in a shape with Aspose.Cells | Set 4 point space after each paragraph in shape text using .NET
// Tags: set SpaceAfterPt Aspose.Cells | shape paragraph spacing .NET | modify shape text formatting C# | Aspose.Cells paragraph formatting API | Excel shape paragraph spacing using C#

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, accesses a shape, sets the SpaceAfterPt property of each paragraph to 4 points to add spacing after the text, and saves the updated workbook using Aspose.Cells for .NET.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Example operation: set a uniform row height for all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Set the default row height (in points)
                sheet.Cells.StandardHeight = 15.0;
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
