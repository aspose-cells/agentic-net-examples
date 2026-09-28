// Title: Highlight duplicate values in a named range of an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to iterate over a named range, identify cells with duplicate values, and apply a yellow background style to those cells. | Adjust the duplicate‑highlighting example to use a red fill and perform case‑insensitive string comparison when detecting duplicates. | Extend the sample to log the addresses of all duplicate cells to a text file while keeping the visual highlighting intact.
// Common Searches: asp.net locate duplicate cells inside a defined named range with Aspose.Cells | c# code to mark repeated values in an Excel named range using Aspose.Cells library | apply background color to duplicate entries in a specific range with Aspose.Cells for .NET | save workbook after highlighting duplicate cells in an Excel file using Aspose.Cells
// Tags: duplicate detection in named range Aspose.Cells | apply background style to duplicate cells C# | map cell values to list for duplicate identification Aspose.Cells | highlight repeated entries Excel workbook Aspose.Cells | create custom style for duplicate highlighting .NET

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, retrieves a named range called "MyRange", builds a dictionary that maps each distinct cell value to the cells containing it, creates a yellow solid‑fill style, and applies this style to every cell whose value appears more than once. It ensures the output directory exists, saves the modified workbook as a new file, and reports any errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range (use GetRangeByName)
            Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName("MyRange");
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Map each distinct value to the list of cells containing it
            Dictionary<object, List<Cell>> valueMap = new Dictionary<object, List<Cell>>();

            foreach (Cell cell in namedRange)
            {
                object val = cell.Value;
                if (val == null) continue; // Skip empty cells

                if (!valueMap.ContainsKey(val))
                    valueMap[val] = new List<Cell>();

                valueMap[val].Add(cell);
            }

            // Create a style for highlighting duplicate values
            Style dupStyle = workbook.CreateStyle();
            dupStyle.ForegroundColor = Color.Yellow;
            dupStyle.Pattern = BackgroundType.Solid;
            StyleFlag styleFlag = new StyleFlag { CellShading = true };

            // Apply the highlight style to cells that have duplicate values
            foreach (var entry in valueMap)
            {
                if (entry.Value.Count > 1) // Duplicate found
                {
                    foreach (Cell dupCell in entry.Value)
                    {
                        dupCell.SetStyle(dupStyle, styleFlag);
                    }
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
