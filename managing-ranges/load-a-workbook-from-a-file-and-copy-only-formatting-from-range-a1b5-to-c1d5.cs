// Title: Copy Formatting Only from A1:B5 to C1:D5 with Aspose.Cells (.NET C#)
// Description: The sample loads (or creates) an Excel workbook, applies a light‑blue solid fill to cells A1:B5, and then moves just the style data to cells C1:D5 using Aspose.Cells' CopyStyle method, leaving any existing values untouched.
// Keywords: Aspose.Cells C# copy formatting | CopyStyle method | transfer cell style .NET | Excel range style copy | copy only formatting Aspose | range A1:B5 to C1:D5 | apply style without data
// Common Searches: Aspose.Cells copy only cell style C# | How to use CopyStyle in Aspose.Cells | Copy formatting between ranges in .NET Excel | Transfer Excel range styling without values | C# example for copying styles with Aspose.Cells
// Developer Intent: Move the visual styling of the source range to the destination range while preserving the destination's cell contents.
// Use Cases: Standardize header appearance across multiple table sections | Reuse a predefined theme for a new data block without overwriting values | Migrate conditional‑formatting rules to another area of the sheet | Create a template where only styles are propagated to fresh data
// AI Prompts: Write a C# program that uses Aspose.Cells to copy only the style from range A1:B5 to C1:D5, keeping existing cell values unchanged. | Explain step‑by‑step how the CopyStyle method works in Aspose.Cells and show how to verify that only formatting was transferred.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The sample loads (or creates) an Excel workbook, applies a light‑blue solid fill to cells A1:B5, and then moves just the style data to cells C1:D5 using Aspose.Cells' CopyStyle method, leaving any existing values untouched.
class CopyFormattingExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                Workbook tempWb = new Workbook();
                Worksheet tempWs = tempWb.Worksheets[0];
                Cells tempCells = tempWs.Cells;

                // Populate source range with sample data.
                tempCells["A1"].PutValue("Header1");
                tempCells["B1"].PutValue("Header2");
                for (int i = 2; i <= 5; i++)
                {
                    tempCells[$"A{i}"].PutValue($"R{i - 1}C1");
                    tempCells[$"B{i}"].PutValue($"R{i - 1}C2");
                }

                // Apply a simple style to the source range.
                Style style = tempWb.CreateStyle();
                style.ForegroundColor = Color.LightBlue;
                style.Pattern = BackgroundType.Solid;
                Aspose.Cells.Range srcRange = tempCells.CreateRange("A1:B5");
                srcRange.ApplyStyle(style, new StyleFlag { All = true });

                tempWb.Save(inputPath);
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define source and destination ranges (fully qualified to avoid ambiguity).
            Aspose.Cells.Range sourceRange = cells.CreateRange("A1:B5");
            Aspose.Cells.Range destinationRange = cells.CreateRange("C1:D5");

            // Copy only the formatting (styles) from source to destination.
            destinationRange.CopyStyle(sourceRange);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
