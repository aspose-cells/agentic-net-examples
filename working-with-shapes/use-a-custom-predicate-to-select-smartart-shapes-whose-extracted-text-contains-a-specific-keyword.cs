// Title: Use Aspose.Cells for .NET to highlight SmartArt shapes whose AlternativeText contains a specific keyword
// AI Prompts: Write C# code that loads an Excel workbook, uses a LINQ predicate to select SmartArt shapes whose AlternativeText includes a given keyword, and changes their FillFormat.ForeColor to yellow with Aspose.Cells. | Create a reusable method that accepts a file path, keyword, and color, then iterates all worksheets to find SmartArt shapes matching the keyword in AlternativeText and updates their fill color.
// Common Searches: asp.net c# find SmartArt shapes with specific alternative text using Aspose.Cells | how to change fill color of SmartArt shapes based on a keyword in an Excel file with Aspose.Cells | filter SmartArt shapes by AlternativeText property in a .xlsx workbook using LINQ | highlight Excel SmartArt objects that contain the word "Important" programmatically
// Tags: filter SmartArt shapes by AlternativeText Aspose.Cells | highlight SmartArt shape fill color C# | LINQ predicate for SmartArt selection Aspose.Cells | modify Excel SmartArt fill based on keyword | Aspose.Cells shape processing with AlternativeText

using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, iterates through each worksheet, selects SmartArt shapes whose AlternativeText contains the keyword "Important", sets their fill color to yellow via the FillFormat property, and saves the updated file while handling possible file and shape errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string keyword = "Important";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Select SmartArt shapes whose alternative text contains the keyword
            var smartArtShapes = sheet.Shapes
                .Where(s => s.IsSmartArt && !string.IsNullOrEmpty(s.AlternativeText) && s.AlternativeText.Contains(keyword));

            // Highlight each matching SmartArt shape
            foreach (Shape shape in smartArtShapes)
            {
                try
                {
                    // Set fill color using the correct FillFormat property
                    shape.FillFormat.ForeColor = Color.Yellow;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to modify shape fill: {ex.Message}");
                }
            }
        }

        try
        {
            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
