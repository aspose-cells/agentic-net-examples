using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class SetSummaryBlockOutline
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the summary block range (example: B2:E6)
            AsposeRange summaryRange = worksheet.Cells.CreateRange("B2", "E6");

            // Fill the range with sample data so the borders are visible
            for (int row = 1; row <= 5; row++)          // rows 2‑6 (zero‑based index)
            {
                for (int col = 1; col <= 4; col++)      // columns B‑E (zero‑based index)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Apply a double‑line outline border with black color to the outer edges of the range
            summaryRange.SetOutlineBorders(CellBorderType.Double, Color.Black);

            // Save the workbook
            string outputPath = "SummaryBlockOutline.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}