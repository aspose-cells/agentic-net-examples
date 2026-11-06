// Title: Programmatically set a monospaced Consolas font for the workbook theme and all cells with Aspose.Cells for .NET
// Common Searches: how to change the default font to Consolas in an Excel workbook using Aspose.Cells C# | Aspose.Cells C# set monospaced font for all worksheets and existing cells | update Excel theme font to Consolas programmatically with Aspose.Cells .NET | apply a monospaced font to every cell in an existing .xlsx file using Aspose.Cells | C# code to enforce Consolas font across an entire workbook with Aspose.Cells
// Tags: Aspose.Cells default workbook font | monospaced font for Excel worksheets C# | set cell style font name Aspose.Cells | programmatic Excel theme font change .NET | apply Consolas font to all cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, sets the workbook's DefaultStyle font to "Consolas", updates each worksheet's default style, and iterates through all populated cells to apply the same font before saving the file.
class UpdateThemeFont
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Set a monospaced font (e.g., Consolas) as the default font for the workbook
            workbook.DefaultStyle.Font.Name = "Consolas";

            // Apply the same font to each worksheet's default cell style
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Update the worksheet's default style
                    Style wsStyle = sheet.Cells.Style;
                    wsStyle.Font.Name = "Consolas";
                    sheet.Cells.Style = wsStyle;

                    // Optionally, update all existing cells that already have a style
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            Style cellStyle = cell.GetStyle();
                            cellStyle.Font.Name = "Consolas";
                            cell.SetStyle(cellStyle);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing worksheet '{sheet.Name}': {ex.Message}");
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
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display the message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
