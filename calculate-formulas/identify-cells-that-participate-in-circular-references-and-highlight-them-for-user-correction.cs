// Title: Highlight Circular Reference Cells in an Excel Workbook Using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to highlight cells involved in circular references | detect and mark circular reference errors while calculating formulas with Aspose.Cells | programmatically apply yellow background to Excel cells that cause circular reference exceptions | iterate through formula cells and catch circular reference exceptions using Aspose.Cells .NET | find circular reference cells in a workbook and save highlighted version with Aspose.Cells
// Tags: circular reference detection Aspose.Cells | highlight circular reference cells C# | formula calculation exception handling Aspose.Cells | apply solid yellow style to Excel cells Aspose.Cells | iterate formula cells workbook Aspose.Cells | save highlighted workbook Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace CircularReferenceHighlighter
{
    // The example loads an Excel file, walks through each worksheet and formula cell, attempts to calculate each formula with CalculationOptions, catches exceptions that contain "circular reference", and highlights the offending cells by applying a solid yellow background style before saving the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Options for formula calculation
            var calcOptions = new CalculationOptions();

            // Iterate through each worksheet and examine formula cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                foreach (Cell cell in cells)
                {
                    if (!cell.IsFormula) continue;

                    try
                    {
                        // Calculate the formula; if a circular reference exists an exception is thrown
                        cell.Calculate(calcOptions);
                    }
                    catch (Exception ex)
                    {
                        // Detect circular reference by inspecting the exception message
                        if (ex.Message.IndexOf("circular reference", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Highlight cells involved in a circular reference
                            Style style = cell.GetStyle();
                            style.ForegroundColor = Color.Yellow;
                            style.Pattern = BackgroundType.Solid;
                            cell.SetStyle(style);
                        }
                        else
                        {
                            // Log other calculation errors without stopping the process
                            Console.WriteLine($"Error calculating cell {cell.Name}: {ex.Message}");
                        }
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
