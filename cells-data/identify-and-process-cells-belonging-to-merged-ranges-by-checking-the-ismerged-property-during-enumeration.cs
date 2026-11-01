// Title: Highlight merged cells with a yellow background using Aspose.Cells for .NET
// Common Searches: how to detect merged cells in an Excel worksheet with Aspose.Cells C# | apply a solid background color to a merged range using Aspose.Cells .NET | enumerate worksheet cells and check the IsMerged property in Aspose.Cells | retrieve the full merged range of a cell with Aspose.Cells C# example | set style for merged cells in an Excel file using Aspose.Cells
// Tags: detect merged cells Aspose.Cells | apply style to merged range Aspose.Cells | highlight merged cells Excel .NET | enumerate cells IsMerged Aspose.Cells | set background color merged range C#

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an existing Excel workbook, iterates through each worksheet and all populated cells, identifies cells that belong to merged ranges via the IsMerged flag, obtains the complete merged range, applies a yellow solid background style to the entire range, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Loop through all cells that contain data
                for (int row = 0; row <= cells.MaxDataRow; row++)
                {
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        Cell cell = cells[row, col];

                        // Identify cells that are part of a merged range
                        if (cell.IsMerged)
                        {
                            Console.WriteLine($"Merged cell at {cell.Name} in sheet '{sheet.Name}'");

                            // Retrieve the entire merged range (use fully qualified type to avoid ambiguity)
                            Aspose.Cells.Range mergedRange = cell.GetMergedRange();

                            // Apply a yellow background to the whole merged range
                            Style style = workbook.CreateStyle();
                            style.ForegroundColor = Color.Yellow;
                            style.Pattern = BackgroundType.Solid;

                            StyleFlag flag = new StyleFlag();
                            flag.CellShading = true;

                            mergedRange.ApplyStyle(style, flag);
                        }
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
