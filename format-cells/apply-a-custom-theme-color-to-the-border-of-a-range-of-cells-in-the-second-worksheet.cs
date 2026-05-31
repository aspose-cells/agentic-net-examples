using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class ApplyCustomThemeBorderToSecondWorksheet
    {
        public static void Main()
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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a second worksheet (index 1)
            workbook.Worksheets.Add();
            Worksheet secondSheet = workbook.Worksheets[1];

            // Define the range where the border will be applied
            AsposeRange range = secondSheet.Cells.CreateRange("B2:D5");

            // Create a CellsColor object and set its theme color (e.g., Accent2)
            CellsColor themeBorderColor = workbook.CreateCellsColor();
            themeBorderColor.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);

            // Apply an outline border with the theme color to the entire range
            // Using Thin style for all four edges
            range.SetOutlineBorders(CellBorderType.Thin, themeBorderColor);

            // Optionally, put some sample data to visualize the border
            for (int row = 1; row <= 4; row++)          // rows 2-5 (0‑based index)
            {
                for (int col = 1; col <= 3; col++)      // columns B-D (0‑based index)
                {
                    secondSheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Save the workbook
            string outputPath = "CustomThemeBorder.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}