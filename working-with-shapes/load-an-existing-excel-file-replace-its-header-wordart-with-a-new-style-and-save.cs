// Title: Replace Excel Header WordArt with a New Preset Style using Aspose.Cells for .NET (C#)
// Description: Loads a workbook, locates the first WordArt shape on a worksheet, records its geometry and text, removes the original shape, inserts a new WordArt with a chosen PresetWordArtStyle (e.g., WordArtStyle5), and saves the updated file.
// Keywords: Aspose.Cells | C# | WordArt | replace WordArt | Excel header | preset WordArt style | shape manipulation | worksheet shapes | add WordArt | remove WordArt | Aspose.Cells for .NET | Excel automation | document styling
// Common Searches: Aspose.Cells replace WordArt header C# | Change WordArt style in Excel with Aspose | How to update WordArt shape programmatically .NET | Add preset WordArt to worksheet using Aspose.Cells | Remove and insert WordArt preserving position | C# code to modify Excel shapes Aspose
// Developer Intent: Programmatically replace an existing WordArt header in an Excel workbook with a different preset style while preserving its original text and layout.
// Use Cases: Refresh corporate report templates by applying a new branded WordArt header across all generated workbooks. | Migrate legacy Excel files that contain outdated WordArt headers to a standardized style for visual consistency. | Create invoices where the header WordArt is automatically restyled based on client‑specific branding rules.
// AI Prompts: Generate C# code using Aspose.Cells that finds the first WordArt shape on the first worksheet, captures its position and text, deletes it, adds a new WordArt with PresetWordArtStyle.WordArtStyle5, and saves the workbook. | Provide a reusable method for Aspose.Cells that accepts input and output file paths, a worksheet index, and a PresetWordArtStyle value, then replaces any WordArt header while handling missing files and preserving geometry. | Write a robust Aspose.Cells script that iterates through all shapes, identifies WordArt headers, replaces each with a specified preset style, and logs the changes for batch processing of multiple Excel files.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, locates the first WordArt shape on a worksheet, records its geometry and text, removes the original shape, inserts a new WordArt with a chosen PresetWordArtStyle (e.g., WordArtStyle5), and saves the updated file.
class ReplaceHeaderWordArt
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            for (int i = 0; i < worksheet.Shapes.Count; i++)
            {
                Shape shape = worksheet.Shapes[i];

                // Identify WordArt shapes (used as header in this scenario)
                if (shape.IsWordArt)
                {
                    // Preserve original geometry and text
                    int upperLeftRow = shape.UpperLeftRow;
                    int upperLeftColumn = shape.UpperLeftColumn;
                    int top = shape.Top;
                    int left = shape.Left;
                    int height = shape.Height;
                    int width = shape.Width;
                    string text = shape.Text;

                    // Remove the existing WordArt shape
                    worksheet.Shapes.RemoveAt(i);

                    // Add a new WordArt shape with the desired preset style
                    // Example: using WordArtStyle5 (Fill - Gold, Accent 4, Soft Bevel)
                    worksheet.Shapes.AddWordArt(
                        PresetWordArtStyle.WordArtStyle5,
                        text,
                        upperLeftRow,
                        top,
                        upperLeftColumn,
                        left,
                        height,
                        width);

                    // Since we replaced the header, exit the loop (assuming only one header WordArt)
                    break;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
