// Title: Load an Excel workbook with Aspose.Cells for .NET, apply a red foreground and yellow background solid fill to every used cell, verify the fill pattern, and save the result
// AI Prompts: Write a C# program that opens an existing .xlsx file using Aspose.Cells, creates a style with a red foreground and yellow background solid fill, applies the style to the entire used range of the first worksheet, checks that each cell’s BackgroundType is Solid, and saves the modified workbook. | Modify the program to log the address of any cell whose fill pattern is not solid and output a summary indicating how many cells passed or failed verification. | Add a try‑catch block around the workbook.Save call to capture and display detailed error information if the file cannot be written.
// Common Searches: how to apply a solid fill style to all used cells in an Excel file using Aspose.Cells C# | verify cell background pattern after applying style with Aspose.Cells .NET | Aspose.Cells C# change cell foreground and background colors for a range | error handling when saving a workbook with Aspose.Cells in C# | iterate over used range of worksheet to set style Aspose.Cells example
// Tags: solid fill for used cells Aspose.Cells | verify cell background pattern C# | load and save workbook with custom style Aspose.Cells | error handling for workbook.Save Aspose.Cells | iterate over worksheet cells Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, creates a style with red foreground and yellow background solid fill, applies it to every cell in the worksheet's used range, verifies that each cell's fill pattern is solid, and saves the workbook as output.xlsx using Aspose.Cells for .NET.
class ThemeAccentGradientExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed).
            Worksheet sheet = workbook.Worksheets[0];

            // Create a solid style that will replace existing fills.
            Style replaceStyle = workbook.CreateStyle();
            replaceStyle.Pattern = BackgroundType.Solid;          // Solid fill
            replaceStyle.ForegroundColor = Color.Red;            // Red foreground
            replaceStyle.BackgroundColor = Color.Yellow;         // Yellow background

            // Get the used range of the worksheet.
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // Apply the replacement style to all used cells.
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    cell.SetStyle(replaceStyle);
                }
            }

            // Verify that the style was applied.
            bool verificationPassed = true;
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    Style cellStyle = cell.GetStyle();

                    if (cellStyle.Pattern != BackgroundType.Solid)
                    {
                        verificationPassed = false;
                        Console.WriteLine($"Verification failed at cell {cell.Name}");
                    }
                }
            }

            Console.WriteLine(verificationPassed
                ? "All cells successfully updated to solid fill."
                : "Some cells were not updated correctly.");

            // Save the modified workbook.
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
