// Title: C# – Set Column Width in Pixels After AutoFit with Aspose.Cells for .NET
// Description: Creates a workbook, fills columns A‑E, auto‑fits them, then rounds each column's pixel width to the nearest 5 pixels using GetColumnWidthPixel and SetColumnWidthPixel, and saves the file.
// Keywords: Aspose.Cells | SetColumnWidthPixel | GetColumnWidthPixel | AutoFitColumns | column width pixel | C# example | Excel column sizing | precise pixel alignment | workbook manipulation | pixel‑perfect columns
// Common Searches: Aspose.Cells set column width pixel after autofit | C# round column width to nearest 5 pixels | GetColumnWidthPixel usage example | AutoFitColumns then SetColumnWidthPixel | pixel‑perfect column sizing Aspose.Cells for .NET
// Developer Intent: Adjust column widths to a specific pixel increment after auto‑fitting them.
// Use Cases: Design print‑ready reports where column spacing must follow a fixed pixel grid. | Generate web‑based spreadsheets that require consistent pixel‑based column widths for UI alignment. | Standardize column dimensions across multiple worksheets before exporting to PDF or image formats.
// AI Prompts: Generate C# code that auto‑fits a range of columns with Aspose.Cells and then sets each column width to the nearest 10 pixels. | Explain how GetColumnWidthPixel and SetColumnWidthPixel can be combined to achieve pixel‑perfect column alignment after AutoFit in Aspose.Cells. | Create a reusable C# method that accepts a pixel step value and applies it to a given column range after AutoFit using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills columns A‑E, auto‑fits them, then rounds each column's pixel width to the nearest 5 pixels using GetColumnWidthPixel and SetColumnWidthPixel, and saves the file.
    public class SetColumnWidthPixelAfterAutoFitDemo
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in columns A to E (indices 0‑4)
            for (int col = 0; col < 5; col++)
            {
                for (int row = 0; row < 10; row++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1} - Some long text");
                }
            }

            // Auto‑fit the columns based on the data just added
            worksheet.AutoFitColumns(0, 4);

            // Desired pixel alignment: round each column width to the nearest 5 pixels
            const int alignmentStep = 5;

            // Apply precise pixel widths after auto‑fit
            for (int col = 0; col < 5; col++)
            {
                // Get the current width in pixels after auto‑fit
                int currentPixels = cells.GetColumnWidthPixel(col);

                // Calculate the aligned width (nearest multiple of alignmentStep)
                int alignedPixels = ((currentPixels + alignmentStep / 2) / alignmentStep) * alignmentStep;

                // Set the column width to the aligned pixel value
                cells.SetColumnWidthPixel(col, alignedPixels);
            }

            // Save the workbook
            workbook.Save("SetColumnWidthPixelAfterAutoFitDemo.xlsx");
        }
    }
}
