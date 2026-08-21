// Title: Copy range formatting only with Aspose.Cells for .NET (C#) using CopyStyle
// Description: Shows how to copy just the formatting from a source range (A1:C3) to a destination range (E1:G3) in an Excel workbook with Aspose.Cells for .NET. The sample creates a workbook, fills values, applies a bold Calibri style with a light‑blue background, and transfers the style via the CopyStyle method while leaving cell data untouched.
// Keywords: Aspose.Cells | CopyStyle | C# | .NET | range formatting | copy cell style | preserve cell values | Excel formatting | Aspose.Cells example | copy style between ranges
// Common Searches: Aspose.Cells copy only formatting | CopyStyle method C# example | How to copy cell style without values Aspose.Cells | Transfer range formatting in .NET | Copy formatting between ranges Aspose.Cells
// Developer Intent: Copy the formatting of a source range to a destination range while keeping the cell values unchanged.
// Use Cases: Replicate a styled template block across multiple sections of a report without overwriting existing data. | Create a print‑ready layout on a separate sheet by copying visual formatting while preserving original calculations. | Synchronize visual styles across worksheets in a workbook without modifying the underlying cell contents.
// AI Prompts: Generate C# code that copies only the formatting from range A1:C3 to E1:G3 using Aspose.Cells, leaving the cell values intact. | Provide an Aspose.Cells example that copies a style between two ranges and then changes the font color of the destination range. | Explain the CopyStyle method, including how to handle exceptions and apply a custom Style object in a .NET workbook.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Entry point for the console application
    // Shows how to copy just the formatting from a source range (A1:C3) to a destination range (E1:G3) in an Excel workbook with Aspose.Cells for .NET. The sample creates a workbook, fills values, applies a bold Calibri style with a light‑blue background, and transfers the style via the CopyStyle method while leaving cell data untouched.
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                CopyFormattingOnly.Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class CopyFormattingOnly
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Create source range and fill it with data
                Aspose.Cells.Range srcRange = cells.CreateRange("A1:C3");
                for (int i = 0; i < 3; i++)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        srcRange[i, j].PutValue($"R{i + 1}C{j + 1}");
                    }
                }

                // Define a style and apply it to the source range
                Style style = workbook.CreateStyle();
                style.Font.Name = "Calibri";
                style.Font.Size = 12;
                style.Font.IsBold = true;
                style.ForegroundColor = Color.LightBlue;
                style.Pattern = BackgroundType.Solid;
                srcRange.SetStyle(style);

                // Create destination range (same dimensions)
                Aspose.Cells.Range destRange = cells.CreateRange("E1:G3");

                // Copy only the formatting from source to destination
                destRange.CopyStyle(srcRange);

                // Save the workbook
                string outputPath = "CopyFormattingOnly.xlsx";
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Propagate exception to caller
                throw new ApplicationException("Failed to copy formatting.", ex);
            }
        }
    }
}
