using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class HeaderFillPatternDemo
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate header cells with sample text
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Score");

                // Create a style for the header cells
                Style headerStyle = workbook.CreateStyle();

                // Apply a dense hatch pattern (DiagonalCrosshatch) to the style
                headerStyle.Pattern = BackgroundType.DiagonalCrosshatch;

                // Set foreground (text) and background colors for the pattern
                headerStyle.ForegroundColor = Color.White;
                headerStyle.BackgroundColor = Color.DarkBlue;

                // Enable cell shading (pattern application)
                StyleFlag flag = new StyleFlag { CellShading = true };

                // Define the range that covers the header row (A1:C1)
                int startRow = 0;
                int startColumn = 0;
                int totalRows = 1;
                int totalColumns = 3;

                // Resolve ambiguity between Aspose.Cells.Range and System.Range
                Aspose.Cells.Range headerRange = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);
                headerRange.ApplyStyle(headerStyle, flag);

                // Determine output file path
                string outputPath = "HeaderFillPattern.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}