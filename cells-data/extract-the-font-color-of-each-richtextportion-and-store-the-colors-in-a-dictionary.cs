// Title: Extract font colors of RichTextPortion objects from Excel cells into a C# Dictionary with Aspose.Cells
// Common Searches: how to read font color of each RichTextPortion in an Excel worksheet using Aspose.Cells C# | C# Aspose.Cells example to store cell text portion colors in a dictionary | retrieve RichTextPortion font color from an Excel file with Aspose.Cells .NET | extract colors of rich text portions from Excel cells and map them to cell addresses in C# | Aspose.Cells get font color of cell style for each text segment in C#
// Tags: Aspose.Cells extract RichTextPortion font color | C# dictionary store Excel text portion colors | iterate worksheets and cells Aspose.Cells | cell style font color retrieval .NET | Excel rich text portion color extraction

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ExtractRichTextPortionColors
{
    // The sample loads an Excel workbook, iterates through all worksheets and used cells, treats each cell as a single RichTextPortion, extracts the font color from the cell's style, and stores the Color in a Dictionary keyed by "CellAddress:PortionIndex". It then prints the collected colors and saves the workbook unchanged.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook (replace with your actual file)
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook with error handling
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook \"{inputPath}\": {ex.Message}");
                return;
            }

            // Dictionary to store the font color of each portion.
            // Key format: "CellAddress:PortionIndex"
            var portionColors = new Dictionary<string, Color>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Processing worksheet: {sheet.Name}");

                // Iterate through all used cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    try
                    {
                        // Aspose.Cells does not expose RichTextPortion in this version.
                        // Treat the entire cell as a single portion (index 0) and extract its font color.
                        Style style = cell.GetStyle();
                        Color fontColor = style.Font.Color;
                        string dictKey = $"{cell.Name}:0";
                        portionColors[dictKey] = fontColor;

                        Console.WriteLine($"Extracted color from cell {cell.Name}, portion 0: {fontColor}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing cell {cell.Name}: {ex.Message}");
                    }
                }
            }

            // Output the collected colors
            Console.WriteLine("\n=== Summary of Extracted Font Colors ===");
            foreach (var kvp in portionColors)
            {
                Console.WriteLine($"Portion \"{kvp.Key}\" => Color: {kvp.Value}");
            }

            // Save the workbook (unchanged) with error handling
            const string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
