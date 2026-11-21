// Title: Identify and highlight circular‑reference cells in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# code to find and highlight circular reference cells with Aspose.Cells | how to automatically resolve circular references in an Excel file using Aspose.Cells | Aspose.Cells CalculateFormula catch circular reference exception and mark cells | extract offending cell address from Aspose.Cells circular reference error message | replace circular reference formula with its evaluated value in .NET
// Tags: detect circular references Aspose.Cells .NET | apply yellow fill style Aspose.Cells | replace formula with value to break circular reference C# | parse cell address from Aspose.Cells exception | loop CalculateFormula until no circular references Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, creates a yellow fill style, and repeatedly calls CalculateFormula while catching circular‑reference exceptions. It parses the offending cell address from the exception message, applies the highlight style, clears the formula by inserting its evaluated value, and saves the corrected workbook.
class CircularReferenceHighlighter
{
    static void Main()
    {
        const string inputPath = @"Input.xlsx";
        const string outputPath = @"Output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
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
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Create a style for highlighting circular‑reference cells (yellow fill)
        Style highlightStyle = workbook.CreateStyle();
        highlightStyle.ForegroundColor = Color.Yellow;
        highlightStyle.Pattern = BackgroundType.Solid;
        StyleFlag styleFlag = new StyleFlag { CellShading = true };

        // Repeatedly calculate formulas, removing circular references one by one
        while (true)
        {
            try
            {
                workbook.CalculateFormula();
                // No exception → all circular references resolved
                break;
            }
            // Catch circular reference errors by inspecting the exception message
            catch (Exception ex) when (ex.Message != null && ex.Message.IndexOf("circular reference", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Extract cell address from the exception message
                string message = ex.Message;
                string cellName = string.Empty;

                int start = message.IndexOf("cell ", StringComparison.OrdinalIgnoreCase);
                if (start >= 0)
                {
                    start += 5; // length of "cell "
                    int end = message.IndexOfAny(new[] { '.', ' ', '\n', '\r' }, start);
                    if (end > start)
                        cellName = message.Substring(start, end - start);
                }

                if (string.IsNullOrEmpty(cellName))
                {
                    Console.WriteLine($"Unable to parse cell name from exception: {message}");
                    break;
                }

                // Locate the cell (first worksheet is used here)
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet?.Cells[cellName];
                if (cell == null)
                {
                    Console.WriteLine($"Cell {cellName} not found.");
                    break;
                }

                // Apply highlight style
                cell.SetStyle(highlightStyle, styleFlag);

                // Break the circular reference by replacing the formula with its current value
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    object evaluatedValue = cell.Value;
                    cell.Formula = string.Empty;
                    cell.PutValue(evaluatedValue);
                }
                else
                {
                    // No formula to clear – exit to avoid infinite loop
                    break;
                }
            }
            catch (Exception ex)
            {
                // Unexpected exception – report and exit
                Console.WriteLine($"Unexpected error during calculation: {ex.Message}");
                break;
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
