// Title: How to change the entire workbook’s font to Arial (sans‑serif) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change default workbook font to Arial | Set all cells to a sans‑serif font in an existing Excel file with Aspose | Update the theme font scheme globally in a .NET Excel workbook using Aspose.Cells | Programmatically apply Arial font to every worksheet in Aspose.Cells | Modify default style font before saving a workbook with Aspose.Cells
// Tags: Aspose.Cells default style font | set workbook font Arial C# | apply global font to Excel workbook .NET | iterate worksheets cells set font Aspose | save modified workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX file, changes the workbook’s default style font to Arial, iterates through each worksheet and cell to enforce the Arial font, ensures the output directory exists, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Set default style font to Arial
            Style defaultStyle = workbook.DefaultStyle;
            defaultStyle.Font.Name = "Arial";
            workbook.DefaultStyle = defaultStyle;

            // Optionally update all cells' style to use Arial
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell != null)
                        {
                            Style style = cell.GetStyle();
                            style.Font.Name = "Arial";
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
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
