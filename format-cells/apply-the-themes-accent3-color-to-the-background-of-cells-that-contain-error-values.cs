// Title: Highlight Excel error cells with the workbook's Accent3 theme color using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set background of cells that contain error values to theme Accent3 | apply workbook theme Accent3 color to error cells in Excel with Aspose.Cells | retrieve Accent3 theme color and format error cells using Aspose.Cells .NET | iterate through worksheets and color cells with #N/A or other errors using Aspose.Cells | Aspose.Cells change cell style based on CellValueType.IsError in C#
// Tags: Aspose.Cells theme Accent3 background for error cells | C# style cells with CellValueType.IsError | retrieve workbook theme color using Aspose.Cells | apply solid fill to cells based on error value | iterate worksheets to format error cells Aspose

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, obtains the Accent3 color from the workbook's theme (fallback to LightBlue), scans every worksheet's used range, detects cells whose value type is an error, and applies a solid foreground fill with the Accent3 color before saving the updated file.
class ApplyAccent3ToErrorCells
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the Accent3 color from the workbook's theme (fallback to LightBlue if unavailable)
            Color accent3Color = Color.LightBlue;
            try
            {
                accent3Color = workbook.GetThemeColor(ThemeColorType.Accent3);
            }
            catch
            {
                // Keep default fallback color
            }

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Loop through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Check if the cell contains an error value
                        if (cell.Type == CellValueType.IsError)
                        {
                            // Get the current style of the cell
                            Style style = cell.GetStyle();

                            // Set the background (foreground) color to Accent3
                            style.ForegroundColor = accent3Color;
                            style.Pattern = BackgroundType.Solid;

                            // Apply the modified style back to the cell
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
