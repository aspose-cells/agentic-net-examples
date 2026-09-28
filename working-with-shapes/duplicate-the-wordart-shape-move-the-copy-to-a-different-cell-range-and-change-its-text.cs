// Title: Copy a WordArt (TextEffect) shape, place it in a new cell range, and change its text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to duplicate the first TextEffect shape, position the copy at row 6 column C while keeping its original dimensions, and set the copied shape’s Text to a new string. | Show how to use the AddCopy overload to copy a WordArt shape to a target cell range and then update its Text property in a .NET workbook.
// Common Searches: Aspose.Cells C# copy WordArt shape to specific cells | How to duplicate a TextEffect shape and change its text in Aspose.Cells | Move copied WordArt to cell C6 while preserving size using Aspose.Cells .NET | Shapes.AddCopy example for WordArt in Aspose.Cells C# | Change text of a duplicated shape in an Excel file with Aspose.Cells
// Tags: Aspose.Cells copy WordArt shape C# | Shapes.AddCopy method Aspose.Cells | set Text property TextEffect shape | copy shape while keeping size Aspose.Cells | move shape to cell range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtExample
{
    // The example creates or loads a workbook, copies the first WordArt (TextEffect) shape to cell C6 using Shapes.AddCopy while preserving its dimensions, updates the copied shape's Text, and saves the result to output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; if not, create a simple workbook with a WordArt shape.
                if (!File.Exists(inputPath))
                {
                    var newWb = new Workbook();
                    var newWs = newWb.Worksheets[0];

                    // Add a WordArt shape (TextEffect) as the original shape.
                    // Parameters: preset, text, font name, font size, bold, italic,
                    // upper left row, upper left column, lower right row, lower right column, height, width
                    newWs.Shapes.AddTextEffect(
                        MsoPresetTextEffect.TextEffect1,
                        "Original WordArt",
                        "Arial",
                        24,
                        false,
                        false,
                        0,
                        0,
                        0,
                        0,
                        200,
                        50);

                    newWb.Save(inputPath);
                }

                // Load the existing workbook.
                var workbook = new Workbook(inputPath);
                var worksheet = workbook.Worksheets[0];

                // Verify that there is at least one shape.
                if (worksheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found in the worksheet.");
                    return;
                }

                Shape originalWordArt = worksheet.Shapes[0];

                // Target cell for the copy (e.g., cell C6).
                int targetRow = 5;    // zero‑based index (row 6)
                int targetColumn = 2; // zero‑based index (column C)

                // Calculate the lower‑right cell to preserve the original shape size.
                int rowSpan = originalWordArt.LowerRightRow - originalWordArt.UpperLeftRow;
                int colSpan = originalWordArt.LowerRightColumn - originalWordArt.UpperLeftColumn;
                int targetLowerRightRow = targetRow + rowSpan;
                int targetLowerRightColumn = targetColumn + colSpan;

                // Duplicate the WordArt shape to the specified location.
                // AddCopy overload: (sourceShape, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn)
                Shape copiedWordArt = worksheet.Shapes.AddCopy(
                    originalWordArt,
                    targetRow,
                    targetColumn,
                    targetLowerRightRow,
                    targetLowerRightColumn);

                // Change the text of the duplicated WordArt.
                copiedWordArt.Text = "New WordArt Text";

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the changes.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
