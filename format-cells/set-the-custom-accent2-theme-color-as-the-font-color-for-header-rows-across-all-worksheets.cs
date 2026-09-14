// Title: Apply the workbook's Accent2 theme color as the header row font color across all worksheets with Aspose.Cells for .NET
// AI Prompts: Retrieve the workbook's Accent2 theme color and assign it to the Font.Color of every cell in the first row of each worksheet using Aspose.Cells. | Loop through all worksheets in an Excel file and set the header row's font to the workbook's custom Accent2 theme color in C#. | Use Aspose.Cells to get the Accent2 theme color from a workbook and apply it as the font color for header rows on every sheet.
// Common Searches: Aspose.Cells C# set header row font color to workbook Accent2 theme color | How to apply custom theme color to first row of all sheets using Aspose.Cells | Retrieve Accent2 theme color from Excel workbook with Aspose.Cells .NET | Change header font color to theme Accent2 for multiple worksheets in C# | Apply workbook theme colors to header rows across worksheets Aspose.Cells
// Tags: set header row font color using workbook theme Aspose.Cells | retrieve Accent2 theme color .NET workbook | apply theme color to cells across worksheets | format header cells with custom theme color C# | iterate worksheets and update style Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, obtains its Accent2 theme color, and sets that color as the font color for every cell in the header row of each worksheet before saving the file.
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
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the custom Accent2 theme color using the correct API
            Color accent2Color = workbook.GetThemeColor(ThemeColorType.Accent2);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Header row index (0‑based). Adjust if needed.
                int headerRowIndex = 0;

                // Determine the last used column in the sheet
                int lastColumn = sheet.Cells.MaxDataColumn;

                // Apply the Accent2 color to each cell in the header row
                for (int col = 0; col <= lastColumn; col++)
                {
                    // Get the current style of the cell
                    Style cellStyle = sheet.Cells[headerRowIndex, col].GetStyle();

                    // Set the font color to the Accent2 theme color
                    cellStyle.Font.Color = accent2Color;

                    // Apply the modified style back to the cell
                    sheet.Cells[headerRowIndex, col].SetStyle(cellStyle);
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
