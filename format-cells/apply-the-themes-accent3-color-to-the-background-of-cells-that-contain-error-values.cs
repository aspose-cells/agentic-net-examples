// Title: Highlight cells with error values using the workbook’s Accent3 theme color in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, retrieves the Accent3 theme color, and applies it as a solid fill to every cell whose value type is IsError. | Show how to iterate through all worksheets and their used ranges, detect error cells, and set the cell style to use the workbook’s Accent3 background color. | Provide a complete example that saves the modified workbook after highlighting error cells with the Accent3 theme color, including handling missing input files and creating the output directory.
// Common Searches: Aspose.Cells C# highlight #N/A and #DIV/0! errors with theme accent color | How to set background color of error cells using workbook theme in Aspose.Cells .NET | Retrieve Accent3 color from workbook theme and apply to cells with CellValueType.IsError | C# Aspose.Cells change cell style based on error value type | Apply solid fill to error cells in Excel using Aspose.Cells theme colors
// Tags: apply accent3 theme color to error cells Aspose.Cells | retrieve workbook theme color C# Aspose.Cells | set solid background for CellValueType.IsError Aspose.Cells | iterate worksheets used range Aspose.Cells C# | save workbook after styling error values Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example loads an input workbook, obtains the Accent3 color from the workbook's theme, iterates every worksheet and its used range, and applies a solid Accent3 background to cells whose type is IsError. It ensures the output directory exists, saves the modified workbook, and reports success or errors.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the Accent3 color from the workbook's theme
            Color accent3Color = workbook.GetThemeColor(ThemeColorType.Accent3);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet (explicit Aspose.Cells.Range to avoid ambiguity)
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Loop through each cell in the used range
                for (int row = usedRange.FirstRow; row < usedRange.FirstRow + usedRange.RowCount; row++)
                {
                    for (int col = usedRange.FirstColumn; col < usedRange.FirstColumn + usedRange.ColumnCount; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // If the cell contains an error value, apply Accent3 background
                        if (cell.Type == CellValueType.IsError)
                        {
                            Style style = cell.GetStyle();
                            style.ForegroundColor = accent3Color;
                            style.Pattern = BackgroundType.Solid;
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
