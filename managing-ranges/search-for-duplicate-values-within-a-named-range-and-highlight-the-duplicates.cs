// Title: C# Aspose.Cells example: Highlight duplicate values inside a named range
// Common Searches: C# Aspose.Cells code to mark duplicate cells in a specific named range | How to find and style repeated values in an Excel named range using Aspose.Cells .NET | Programmatically highlight duplicate entries within a defined range with Aspose.Cells | Aspose.Cells .NET example for detecting duplicate data in a workbook range | Apply background color to duplicate cells in a named range using C# and Aspose.Cells
// Tags: duplicate cell highlighting Aspose.Cells C# | named range duplicate detection Aspose.Cells | apply background style to duplicate values Aspose.Cells | Aspose.Cells repeat value identification .NET | style duplicate entries in Excel range Aspose.Cells | C# workbook duplicate value highlighting Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, retrieves the named range "MyRange", builds a case‑insensitive map of cell values, applies a light‑yellow background style to every cell that appears more than once, and saves the result as output.xlsx.
class DuplicateHighlighter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "MyRange"
            Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName("MyRange");
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Map each cell value to the list of cells containing that value
            Dictionary<string, List<Cell>> valueMap = new Dictionary<string, List<Cell>>(StringComparer.OrdinalIgnoreCase);

            foreach (Cell cell in namedRange)
            {
                string cellValue = cell.StringValue ?? string.Empty;
                if (string.IsNullOrWhiteSpace(cellValue))
                    continue;

                if (!valueMap.TryGetValue(cellValue, out List<Cell> cellList))
                {
                    cellList = new List<Cell>();
                    valueMap[cellValue] = cellList;
                }
                cellList.Add(cell);
            }

            // Create a style for highlighting duplicates (light yellow background)
            Style duplicateStyle = workbook.CreateStyle();
            duplicateStyle.ForegroundColor = Color.Yellow;
            duplicateStyle.Pattern = BackgroundType.Solid;

            // Apply the style to cells that have duplicate values
            foreach (var entry in valueMap)
            {
                if (entry.Value.Count > 1)
                {
                    foreach (Cell dupCell in entry.Value)
                    {
                        dupCell.SetStyle(duplicateStyle);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
