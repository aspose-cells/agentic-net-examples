// Title: C# – Replace Excel Header WordArt with a New Styled WordArt using Aspose.Cells
// Description: Loads an existing workbook (or creates a new one), removes any WordArt shapes on the first worksheet, adds a new WordArt header with a preset style and custom font settings, and saves the updated file. Demonstrates shape manipulation, text‑effect formatting, and error handling with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Excel | WordArt | header replacement | shape collection | add WordArt | remove WordArt | preset WordArt style | text effect format | load workbook | save workbook | sample code | GitHub example
// Common Searches: how to replace an Excel WordArt header with Aspose.Cells | C# code to delete and add WordArt in a worksheet | Aspose.Cells remove WordArt shapes example | add styled WordArt header to Excel using .NET | replace Excel header WordArt programmatically
// Developer Intent: Programmatically swap an existing WordArt header in an Excel file for a new styled WordArt and persist the changes.
// Use Cases: Standardize report templates by removing old WordArt headers and inserting a brand‑compliant header. | Automate Excel workbook generation where a custom WordArt title must appear on the first sheet. | Update legacy spreadsheets that contain WordArt placeholders with a new design without altering other data.
// AI Prompts: Write C# code with Aspose.Cells that finds all WordArt shapes on the first worksheet, deletes them, and inserts a new WordArt header using a specified preset style and font properties. | Explain step‑by‑step how to handle missing source files while replacing a WordArt header in an Excel workbook with Aspose.Cells for .NET. | Provide best practices for maintaining Excel templates that use WordArt headers, ensuring the replacement works even when no WordArt shapes are present.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing workbook (or creates a new one), removes any WordArt shapes on the first worksheet, adds a new WordArt header with a preset style and custom font settings, and saves the updated file. Demonstrates shape manipulation, text‑effect formatting, and error handling with Aspose.Cells for .NET.
class ReplaceHeaderWordArt
{
    static void Main()
    {
        // Paths to the source and destination Excel files
        string sourcePath = "HeaderTemplate.xlsx";
        string destinationPath = "HeaderUpdated.xlsx";

        try
        {
            Workbook workbook;

            // Load the existing workbook if it exists; otherwise create a new one
            if (File.Exists(sourcePath))
            {
                workbook = new Workbook(sourcePath);
            }
            else
            {
                Console.WriteLine($"Source file not found: {sourcePath}. A new workbook will be created.");
                workbook = new Workbook();
                workbook.Worksheets.Add("Sheet1");
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of shapes on the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Remove any existing WordArt shapes (assumed to be used as header)
            for (int i = shapes.Count - 1; i >= 0; i--)
            {
                Shape shape = shapes[i];
                if (shape.IsWordArt)
                {
                    shapes.RemoveAt(i);
                }
            }

            // Add a new WordArt shape that will serve as the header
            // Parameters: style, text, topRow, top, leftColumn, left, height, width
            Shape headerWordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle5,   // New preset style
                "New Header Text",                  // Header text
                0,    // topRow (first row)
                5,    // top offset in pixels
                0,    // leftColumn (first column)
                5,    // left offset in pixels
                60,   // height in pixels
                400   // width in pixels
            );

            // Optional: further customize the appearance via TextEffectFormat
            TextEffectFormat textEffect = headerWordArt.TextEffect;
            textEffect.FontBold = true;
            textEffect.FontItalic = false;
            textEffect.FontName = "Calibri";
            textEffect.FontSize = 28;

            // Save the modified workbook
            workbook.Save(destinationPath);
            Console.WriteLine($"Workbook saved successfully to {destinationPath}.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
