// Title: Aspose.Cells for .NET: Ungroup Columns and Keep Uniform Widths
// Description: Loads or creates an Excel workbook, groups columns 2‑4, immediately ungroups them, captures the width of the first column, applies that width to every column up to the sheet's last used column, and saves the result.
// Keywords: Aspose.Cells ungroup columns C# | preserve column width Aspose.Cells | set uniform column width .NET | Excel column operations Aspose | C# worksheet column formatting
// Common Searches: Aspose.Cells how to ungroup columns and retain width | C# set same width for all columns after grouping | maintain column widths when removing outlines Aspose | reset column widths across worksheet .NET
// Developer Intent: Remove a column group while ensuring every column retains the same width as the first column.
// Use Cases: Standardize column widths after temporary grouping for clean export or printing. | Clear column outlines in an imported workbook before further data processing. | Apply a consistent column width across an entire sheet after performing grouping/ungrouping actions.
// AI Prompts: Write C# code with Aspose.Cells that ungroups columns 2‑4 and sets all column widths to match column A. | Show an Aspose.Cells snippet that groups columns, then ungroups them and restores the original width for every column. | Explain how to retrieve the maximum column index in a worksheet and apply a uniform width using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads or creates an Excel workbook, groups columns 2‑4, immediately ungroups them, captures the width of the first column, applies that width to every column up to the sheet's last used column, and saves the result.
    public class UngroupColumnsAndMaintainWidths
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Group columns 1 to 3 (zero‑based indexes)
            cells.GroupColumns(1, 3);

            // Ungroup the previously grouped columns
            cells.UngroupColumns(1, 3);

            // Ensure column widths remain consistent across the worksheet
            double referenceWidth = cells.GetColumnWidth(0);
            int lastColumn = cells.MaxColumn; // zero‑based

            for (int colIndex = 0; colIndex <= lastColumn; colIndex++)
            {
                cells.Columns[colIndex].Width = referenceWidth;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
